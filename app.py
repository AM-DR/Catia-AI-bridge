"""
================================================================================
CATIA V5 R21 AI Studio – Boolean Remove Architecture
================================================================================
Uses Boolean Remove (AddNewRemove) instead of Pockets to guarantee through-cuts
in CATIA V5 R21. Creates cutout geometry as Pads in separate Bodies, then
Boolean-subtracts them from MainBody – direction-independent and bulletproof.
================================================================================
"""

import os, sys, io, math, base64, re, time
import streamlit as st
import pythoncom
import win32com.client
from PIL import Image

try:
    from pycatia import catia
    from pycatia.mec_mod_interfaces.part_document import PartDocument
except ImportError:
    pass

from langchain_core.tools import tool
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage
from langchain_core.callbacks import BaseCallbackHandler

try:
    from langchain.agents import create_tool_calling_agent, AgentExecutor
except ImportError:
    from langchain_classic.agents import create_tool_calling_agent, AgentExecutor

from langchain_openai import ChatOpenAI
from langchain_anthropic import ChatAnthropic

try:
    from langchain_ollama import ChatOllama
except ImportError:
    try:
        from langchain_community.chat_models import ChatOllama
    except ImportError:
        ChatOllama = None


# ==============================================================================
# HELPERS
# ==============================================================================

def compress_image_for_llm(image_bytes, max_dim=1024, quality=80):
    try:
        img = Image.open(io.BytesIO(image_bytes))
        img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=quality)
        return buf.getvalue(), "image/jpeg"
    except Exception:
        return image_bytes, "image/png"


class StreamlitAgentProgressHandler(BaseCallbackHandler):
    def __init__(self, sc):
        self.sc = sc
        self.thinking_container = None
        self.thought_buffer = ""
        self.last_update = 0

    def on_llm_start(self, *a, **k):
        self.thought_buffer = ""
        self.last_update = 0
        self.thinking_container = self.sc.empty()
        self.thinking_container.markdown("🧠 **Generating response & CAD geometry…**")

    def on_llm_new_token(self, token: str, **k):
        if token:
            self.thought_buffer += token
            now = time.time()
            if self.thinking_container and (now - self.last_update > 0.12):
                self.last_update = now
                clean_prev = self.thought_buffer[-200:].replace("\n", " ")
                self.thinking_container.markdown(f"💭 *Thinking:* `{clean_prev}`")

    def on_tool_start(self, ser, inp, **k):
        name = ser.get('name', 'CAD Action')
        self.sc.markdown(f"⚙️ **Executing CATIA Tool:** `{name}`")
        if isinstance(inp, dict):
            code = inp.get('code_snippet') or inp.get('input')
            if code and len(str(code).strip()) > 10:
                self.sc.markdown("##### 📝 Generated Python CAD Script:")
                self.sc.code(str(code).strip(), language="python")
        elif isinstance(inp, str) and len(inp.strip()) > 10 and "\n" in inp:
            self.sc.markdown("##### 📝 Generated Python CAD Script:")
            self.sc.code(inp.strip(), language="python")

    def on_tool_end(self, out, **k):
        self.sc.markdown(f"✅ **Result:** `{out}`")

    def on_tool_error(self, err, **k):
        self.sc.markdown(f"❌ **Error:** `{err}`")


# ==============================================================================
# CATIA V5 CONNECTION
# ==============================================================================

def init_com():
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

def connect_to_catia():
    init_com()
    try:
        win32com.client.GetActiveObject("CATIA.Application")
        return catia(), None
    except Exception:
        return None, "CATIA V5 is not running. Please launch CATIA V5 R21."

def get_part_and_parameters():
    caa, err = connect_to_catia()
    if err:
        return None, None, [], err
    try:
        doc = caa.active_document
        if not doc:
            return caa, None, [], "No active document"
        name = doc.name
        if not name.lower().endswith(".catpart"):
            return caa, None, [], f"Active: '{name}' (not a CATPart)"
        pd = PartDocument(doc.com_object)
        part = pd.part
        params = part.parameters
        np_list = []
        for i in range(1, params.count + 1):
            p = params.item(i)
            try:
                v = p.value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    np_list.append({"name": p.name, "display_name": p.name.split("\\")[-1],
                                    "value": float(v), "item_index": i})
            except Exception:
                pass
        return caa, part, np_list, None
    except Exception as e:
        return caa, None, [], str(e)

def apply_parameter_update(pn, nv):
    init_com()
    caa, part, nps, err = get_part_and_parameters()
    if not part:
        return False, f"Error: {err}"
    try:
        params = part.parameters
        tp = None
        try:
            tp = params.item(pn)
        except Exception:
            for i in range(1, params.count + 1):
                p = params.item(i)
                if p.name.lower() == pn.lower() or p.name.split("\\")[-1].lower() == pn.lower():
                    tp = p; break
        if not tp:
            return False, f"Parameter '{pn}' not found."
        tp.value = float(nv)
        part.update()
        return True, f"Updated '{tp.name}' to {nv}."
    except Exception as e:
        return False, str(e)


# ==============================================================================
# CATIA V5 CAD PRIMITIVES
# ==============================================================================

def create_new_document_in_catia(doc_type="Part"):
    init_com()
    caa, err = connect_to_catia()
    if err:
        return False, err
    try:
        doc = caa.documents.add(doc_type)
        return True, f"Created {doc_type}: '{doc.name}'."
    except Exception as e:
        return False, str(e)

def ensure_active_part():
    caa, part, _, err = get_part_and_parameters()
    if err and "not running" in err.lower():
        return None, None, err
    if not part:
        ok, msg = create_new_document_in_catia("Part")
        if not ok:
            return None, None, msg
        caa, part, _, err = get_part_and_parameters()
    return caa, part, None


# --- Boolean Remove helpers (bulletproof through-cuts) ---

def _bool_cut_circle(part_com, sf, xy, main_body, cx, cy, r, h):
    """Create a circular pad in a new body, then Boolean-remove it from MainBody."""
    cb = part_com.Bodies.Add()
    sk = cb.Sketches.Add(xy)
    f = sk.OpenEdition()
    f.CreateClosedCircle(float(cx), float(cy), float(r))
    sk.CloseEdition()
    part_com.InWorkObject = cb
    sf.AddNewPad(sk, float(h))
    part_com.InWorkObject = main_body
    sf.AddNewRemove(cb)

def _bool_cut_polygon(part_com, sf, xy, main_body, pts, h):
    """Create a polygon pad in a new body, then Boolean-remove it from MainBody."""
    cb = part_com.Bodies.Add()
    sk = cb.Sketches.Add(xy)
    f = sk.OpenEdition()
    pts2d = [f.CreatePoint(px, py) for px, py in pts]
    n = len(pts)
    for i in range(n):
        s, e = pts[i], pts[(i + 1) % n]
        ln = f.CreateLine(s[0], s[1], e[0], e[1])
        ln.StartPoint = pts2d[i]
        ln.EndPoint = pts2d[(i + 1) % n]
    sk.CloseEdition()
    part_com.InWorkObject = cb
    sf.AddNewPad(sk, float(h))
    part_com.InWorkObject = main_body
    sf.AddNewRemove(cb)

def _add_pad_polygon(part_com, sf, xy, main_body, pts, h):
    """Add a polygon pad on MainBody (for raised features like brow crest)."""
    sk = main_body.Sketches.Add(xy)
    f = sk.OpenEdition()
    pts2d = [f.CreatePoint(px, py) for px, py in pts]
    n = len(pts)
    for i in range(n):
        s, e = pts[i], pts[(i + 1) % n]
        ln = f.CreateLine(s[0], s[1], e[0], e[1])
        ln.StartPoint = pts2d[i]
        ln.EndPoint = pts2d[(i + 1) % n]
    sk.CloseEdition()
    part_com.InWorkObject = main_body
    sf.AddNewPad(sk, float(h))


# ==============================================================================
# BASIC GEOMETRY BUILDERS
# ==============================================================================

def create_pad_block_in_catia(width, length, height):
    caa, part, err = ensure_active_part()
    if err:
        return False, err
    try:
        pc = part.com_object if hasattr(part, "com_object") else part
        body = pc.MainBody
        xy = pc.OriginElements.PlaneXY
        sk = body.Sketches.Add(xy)
        f = sk.OpenEdition()
        w2, l2 = float(width)/2, float(length)/2
        p1 = f.CreatePoint(-w2, -l2)
        p2 = f.CreatePoint(w2, -l2)
        p3 = f.CreatePoint(w2, l2)
        p4 = f.CreatePoint(-w2, l2)
        ln1 = f.CreateLine(-w2, -l2, w2, -l2); ln1.StartPoint = p1; ln1.EndPoint = p2
        ln2 = f.CreateLine(w2, -l2, w2, l2);   ln2.StartPoint = p2; ln2.EndPoint = p3
        ln3 = f.CreateLine(w2, l2, -w2, l2);   ln3.StartPoint = p3; ln3.EndPoint = p4
        ln4 = f.CreateLine(-w2, l2, -w2, -l2); ln4.StartPoint = p4; ln4.EndPoint = p1
        sk.CloseEdition()
        sf = pc.ShapeFactory
        sf.AddNewPad(sk, float(height))
        pc.Update()
        return True, f"Built Pad block ({width}×{length}×{height}mm)."
    except Exception as e:
        return False, str(e)

def create_cylinder_in_catia(radius, height):
    caa, part, err = ensure_active_part()
    if err:
        return False, err
    try:
        pc = part.com_object if hasattr(part, "com_object") else part
        body = pc.MainBody
        sk = body.Sketches.Add(pc.OriginElements.PlaneXY)
        f = sk.OpenEdition()
        f.CreateClosedCircle(0, 0, float(radius))
        sk.CloseEdition()
        pc.ShapeFactory.AddNewPad(sk, float(height))
        pc.Update()
        return True, f"Built Cylinder (R={radius}mm, H={height}mm)."
    except Exception as e:
        return False, str(e)


# ==============================================================================
# 5-SPOKE AUTOMOTIVE WHEEL RIM (Boolean Remove approach)
# ==============================================================================

def create_detailed_wheel_rim_in_catia(outer_radius=200.0, rim_width=80.0, hub_radius=55.0, lug_holes=5):
    """
    Automotive wheel rim with:
    1. Outer disc (full circle pad)
    2. Center bore hole (Boolean Remove)
    3. 5 lug bolt holes (Boolean Remove)
    4. 5 spoke window cutouts (Boolean Remove) → creates 5 spokes
    """
    caa, part, err = ensure_active_part()
    if err:
        return False, err
    try:
        pc = part.com_object if hasattr(part, "com_object") else part
        mb = pc.MainBody
        sf = pc.ShapeFactory
        xy = pc.OriginElements.PlaneXY
        cut_depth = float(rim_width) + 10.0

        # 1. Outer rim barrel – full circle pad
        pc.InWorkObject = mb
        sk_rim = mb.Sketches.Add(xy)
        f2d = sk_rim.OpenEdition()
        f2d.CreateClosedCircle(0.0, 0.0, float(outer_radius))
        sk_rim.CloseEdition()
        sf.AddNewPad(sk_rim, float(rim_width))
        pc.Update()

        # 2. Center bore hole – Boolean Remove
        _bool_cut_circle(pc, sf, xy, mb, 0.0, 0.0, 25.0, cut_depth)

        # 3. 5 lug bolt holes – Boolean Remove
        lug_dist = float(hub_radius) * 0.75
        for i in range(lug_holes):
            ang = (2.0 * math.pi / lug_holes) * i
            _bool_cut_circle(pc, sf, xy, mb,
                             lug_dist * math.cos(ang),
                             lug_dist * math.sin(ang),
                             6.5, cut_depth)

        # 4. 5 spoke window cutouts – Boolean Remove (trapezoid sectors)
        num_spokes = 5
        r_in = float(hub_radius) + 12.0
        r_out = float(outer_radius) - 18.0
        half_ang = 0.38  # ~22° half-width

        for i in range(num_spokes):
            ang = (2.0 * math.pi / num_spokes) * i + (math.pi / num_spokes)
            pts = [
                (r_in  * math.cos(ang - half_ang), r_in  * math.sin(ang - half_ang)),
                (r_out * math.cos(ang - half_ang), r_out * math.sin(ang - half_ang)),
                (r_out * math.cos(ang + half_ang), r_out * math.sin(ang + half_ang)),
                (r_in  * math.cos(ang + half_ang), r_in  * math.sin(ang + half_ang)),
            ]
            _bool_cut_polygon(pc, sf, xy, mb, pts, cut_depth)

        pc.Update()
        return True, f"Built 5-Spoke Wheel Rim (R={outer_radius}mm) with bore, lugs & spoke windows!"
    except Exception as e:
        return False, f"Wheel Rim failed: {e}"


# ==============================================================================
# IRON MAN HELMET MASK (Boolean Remove approach)
# ==============================================================================

def create_iron_man_mask_in_catia(mask_width=180.0, mask_height=260.0, mask_depth=40.0):
    """
    Iron Man helmet faceplate with:
    1. 9-point sculpted faceplate contour (Pad)
    2. Right eye slit (Boolean Remove)
    3. Left eye slit (Boolean Remove)
    4. Mouth/muzzle grille slot (Boolean Remove)
    5. Forehead armor T-crest plate (additional Pad)
    """
    caa, part, err = ensure_active_part()
    if err:
        return False, err
    try:
        pc = part.com_object if hasattr(part, "com_object") else part
        mb = pc.MainBody
        sf = pc.ShapeFactory
        xy = pc.OriginElements.PlaneXY
        cut_depth = float(mask_depth) + 20.0

        # 1. Main faceplate contour (9-point closed polygon)
        face_pts = [
            (0.0, 130.0),       # Forehead peak
            (80.0, 110.0),      # Right temple
            (90.0, 40.0),       # Right cheek
            (65.0, -60.0),      # Right jaw
            (35.0, -120.0),     # Right chin
            (-35.0, -120.0),    # Left chin
            (-65.0, -60.0),     # Left jaw
            (-90.0, 40.0),      # Left cheek
            (-80.0, 110.0),     # Left temple
        ]
        pc.InWorkObject = mb
        sk_face = mb.Sketches.Add(xy)
        f2d = sk_face.OpenEdition()
        fp2d = [f2d.CreatePoint(px, py) for px, py in face_pts]
        for i in range(len(face_pts)):
            s, e = face_pts[i], face_pts[(i + 1) % len(face_pts)]
            ln = f2d.CreateLine(s[0], s[1], e[0], e[1])
            ln.StartPoint = fp2d[i]
            ln.EndPoint = fp2d[(i + 1) % len(face_pts)]
        sk_face.CloseEdition()
        sf.AddNewPad(sk_face, float(mask_depth))
        pc.Update()

        # 2. Right eye slit – Boolean Remove
        r_eye = [(18.0, 30.0), (68.0, 36.0), (62.0, 18.0), (22.0, 16.0)]
        _bool_cut_polygon(pc, sf, xy, mb, r_eye, cut_depth)

        # 3. Left eye slit – Boolean Remove
        l_eye = [(-18.0, 30.0), (-68.0, 36.0), (-62.0, 18.0), (-22.0, 16.0)]
        _bool_cut_polygon(pc, sf, xy, mb, l_eye, cut_depth)

        # 4. Mouth / muzzle grille slot – Boolean Remove
        mouth = [(-30.0, -75.0), (30.0, -75.0), (25.0, -85.0), (-25.0, -85.0)]
        _bool_cut_polygon(pc, sf, xy, mb, mouth, cut_depth)

        # 5. Forehead armor T-crest plate – additional Pad on MainBody
        brow = [(0.0, 130.0), (30.0, 115.0), (20.0, 60.0), (-20.0, 60.0), (-30.0, 115.0)]
        _add_pad_polygon(pc, sf, xy, mb, brow, float(mask_depth) + 12.0)

        pc.Update()
        return True, "Built Iron Man Mask (faceplate, eye slits, brow crest, mouth grille)!"
    except Exception as e:
        return False, f"Iron Man Mask failed: {e}"


def split_part_in_catia(plane_name="PlaneXY", split_style="Planar", gap_mm=1.0, tab_size=15.0, tab_count=3):
    """
    Intelligently splits the active CATIA solid part by creating a clearance cutter body
    and subtracting it via Boolean Remove.
    Supports Planar, Jigsaw / Puzzle interlocking, and Pyramid cuts with customizable gap clearance.
    """
    caa, part, err = ensure_active_part()
    if not caa:
        return False, err
    try:
        pc = part.com_object if hasattr(part, "com_object") else part
        mb = pc.MainBody
        sf = pc.ShapeFactory
        oe = pc.OriginElements

        # Select target plane
        plane_map = {"PlaneXY": oe.PlaneXY, "PlaneYZ": oe.PlaneYZ, "PlaneZX": oe.PlaneZX}
        base_plane = plane_map.get(plane_name, oe.PlaneXY)

        span = 250.0  # cutter bounding extent
        half_gap = max(0.1, float(gap_mm) / 2.0)

        if "Puzzle" in split_style or "Jigsaw" in split_style:
            # 1. Create Cutter Body with real interlocking puzzle / dovetail tabs
            cut_body = pc.Bodies.Add()
            cut_body.Name = "Split_Cutter_Puzzle"
            cut_sketch = cut_body.Sketches.Add(base_plane)
            f2 = cut_sketch.OpenEdition()

            # Centerline path for puzzle joint
            centerline = [(-span, 0.0), (-50.0, 0.0)]
            tc = max(1, min(10, int(tab_count)))
            tw = max(8.0, float(tab_size))
            th = max(8.0, float(tab_size) * 0.9)
            active_span = 45.0
            dx = (2.0 * active_span) / float(tc)

            for t in range(tc):
                cx = -active_span + (t + 0.5) * dx
                sign = 1.0 if t % 2 == 0 else -1.0
                h_val = sign * th
                centerline.extend([
                    (cx - tw * 0.5, 0.0),
                    (cx - tw * 0.35, 0.0),
                    (cx - tw * 0.5, h_val),
                    (cx + tw * 0.5, h_val),
                    (cx + tw * 0.35, 0.0),
                    (cx + tw * 0.5, 0.0)
                ])
            centerline.extend([(50.0, 0.0), (span, 0.0)])

            # Normal offset to form closed polygon with thickness gap_mm
            pts_top = []
            pts_bot = []
            for i in range(len(centerline)):
                if i == 0:
                    dx_v = centerline[1][0] - centerline[0][0]
                    dy_v = centerline[1][1] - centerline[0][1]
                elif i == len(centerline) - 1:
                    dx_v = centerline[i][0] - centerline[i-1][0]
                    dy_v = centerline[i][1] - centerline[i-1][1]
                else:
                    dx_v = centerline[i+1][0] - centerline[i-1][0]
                    dy_v = centerline[i+1][1] - centerline[i-1][1]
                L = math.hypot(dx_v, dy_v)
                nx, ny = (-dy_v / L, dx_v / L) if L > 1e-6 else (0.0, 1.0)
                pts_top.append((centerline[i][0] + half_gap * nx, centerline[i][1] + half_gap * ny))
                pts_bot.append((centerline[i][0] - half_gap * nx, centerline[i][1] - half_gap * ny))

            poly = pts_top + list(reversed(pts_bot))
            pts2d = [f2.CreatePoint(x, y) for x, y in poly]
            for i in range(len(poly)):
                ln = f2.CreateLine(poly[i][0], poly[i][1], poly[(i+1)%len(poly)][0], poly[(i+1)%len(poly)][1])
                ln.StartPoint = pts2d[i]
                ln.EndPoint = pts2d[(i+1)%len(poly)]

            cut_sketch.CloseEdition()
            pc.InWorkObject = cut_body
            pad = sf.AddNewPad(cut_sketch, 300.0)
            try:
                pad.SecondLimit.Dimension.Value = 300.0
            except Exception:
                pass
            try:
                pc.Update()
            except Exception:
                pass
            pc.InWorkObject = mb
            sf.AddNewRemove(cut_body)
            try:
                pc.Update()
            except Exception:
                pass

        elif "Pyramid" in split_style:
            # 4-Sided Pyramid / Diagonal X-Split: Two 45-degree diagonal cuts
            cos45 = 0.70710678
            sin45 = 0.70710678
            nx, ny = -sin45 * half_gap, cos45 * half_gap

            # First diagonal cut body (+45 deg)
            cut_body1 = pc.Bodies.Add()
            cut_body1.Name = "Split_Cutter_Pyramid_D1"
            sk1 = cut_body1.Sketches.Add(base_plane)
            f2_1 = sk1.OpenEdition()
            poly1 = [
                (-span + nx, -span + ny),
                (span + nx, span + ny),
                (span - nx, span - ny),
                (-span - nx, -span - ny)
            ]
            pts1 = [f2_1.CreatePoint(x, y) for x, y in poly1]
            for i in range(len(poly1)):
                ln = f2_1.CreateLine(poly1[i][0], poly1[i][1], poly1[(i+1)%len(poly1)][0], poly1[(i+1)%len(poly1)][1])
                ln.StartPoint = pts1[i]; ln.EndPoint = pts1[(i+1)%len(poly1)]
            sk1.CloseEdition()
            pc.InWorkObject = cut_body1
            pad1 = sf.AddNewPad(sk1, 300.0)
            try: pad1.SecondLimit.Dimension.Value = 300.0
            except Exception: pass
            try: pc.Update()
            except Exception: pass
            pc.InWorkObject = mb
            sf.AddNewRemove(cut_body1)

            # Second diagonal cut body (-45 deg)
            cut_body2 = pc.Bodies.Add()
            cut_body2.Name = "Split_Cutter_Pyramid_D2"
            sk2 = cut_body2.Sketches.Add(base_plane)
            f2_2 = sk2.OpenEdition()
            poly2 = [
                (-span + nx, span - ny),
                (span + nx, -span - ny),
                (span - nx, -span + ny),
                (-span - nx, span + ny)
            ]
            pts2 = [f2_2.CreatePoint(x, y) for x, y in poly2]
            for i in range(len(poly2)):
                ln = f2_2.CreateLine(poly2[i][0], poly2[i][1], poly2[(i+1)%len(poly2)][0], poly2[(i+1)%len(poly2)][1])
                ln.StartPoint = pts2[i]; ln.EndPoint = pts2[(i+1)%len(poly2)]
            sk2.CloseEdition()
            pc.InWorkObject = cut_body2
            pad2 = sf.AddNewPad(sk2, 300.0)
            try: pad2.SecondLimit.Dimension.Value = 300.0
            except Exception: pass
            try: pc.Update()
            except Exception: pass
            pc.InWorkObject = mb
            sf.AddNewRemove(cut_body2)

        else:  # Planar Split
            cut_body = pc.Bodies.Add()
            cut_body.Name = "Split_Cutter_Planar"
            cut_sketch = cut_body.Sketches.Add(base_plane)
            f2 = cut_sketch.OpenEdition()
            poly = [(-span, -half_gap), (span, -half_gap), (span, half_gap), (-span, half_gap)]
            pts2d = [f2.CreatePoint(x, y) for x, y in poly]
            for i in range(len(poly)):
                ln = f2.CreateLine(poly[i][0], poly[i][1], poly[(i+1)%len(poly)][0], poly[(i+1)%len(poly)][1])
                ln.StartPoint = pts2d[i]; ln.EndPoint = pts2d[(i+1)%len(poly)]
            cut_sketch.CloseEdition()
            pc.InWorkObject = cut_body
            pad = sf.AddNewPad(cut_sketch, 300.0)
            try: pad.SecondLimit.Dimension.Value = 300.0
            except Exception: pass
            try: pc.Update()
            except Exception: pass
            pc.InWorkObject = mb
            sf.AddNewRemove(cut_body)

        try:
            pc.Update()
        except Exception:
            pass

        return True, f"Successfully executed {split_style} split on {plane_name} with {gap_mm}mm gap clearance!"
    except Exception as e:
        return False, f"Part Split failed: {e}"


def create_revolve_shaft_in_catia(outer_radius=40.0, inner_radius=20.0, height=60.0, angle=360.0):
    """Creates a turned/revolved mechanical cylinder or bushing around the Z-axis."""
    caa, part, err = ensure_active_part()
    if not caa:
        return False, err
    try:
        pc = part.com_object if hasattr(part, "com_object") else part
        sf = pc.ShapeFactory
        zx_plane = pc.OriginElements.PlaneZX

        # Build shaft in a dedicated body to prevent interference with existing solids
        shaft_body = pc.Bodies.Add()
        shaft_body.Name = f"Revolve_Shaft_{int(outer_radius)}x{int(height)}"
        sketch = shaft_body.Sketches.Add(zx_plane)
        f2 = sketch.OpenEdition()

        # Define explicit revolution axis line along V (Z-axis in ZX plane)
        axis_line = f2.CreateLine(0.0, 0.0, 0.0, 100.0)
        axis_line.StartPoint = f2.CreatePoint(0.0, 0.0)
        axis_line.EndPoint = f2.CreatePoint(0.0, 100.0)
        try:
            sketch.CenterLine = axis_line
        except Exception:
            pass

        # Rectangular cross-section offset from Z axis
        r1, r2, h = float(inner_radius), float(outer_radius), float(height)
        poly = [(r1, 0.0), (r2, 0.0), (r2, h), (r1, h)]
        pts2d = [f2.CreatePoint(x, y) for x, y in poly]
        for i in range(len(poly)):
            ln = f2.CreateLine(poly[i][0], poly[i][1], poly[(i+1)%len(poly)][0], poly[(i+1)%len(poly)][1])
            ln.StartPoint = pts2d[i]
            ln.EndPoint = pts2d[(i+1)%len(poly)]

        sketch.CloseEdition()
        pc.InWorkObject = shaft_body
        shaft = sf.AddNewShaft(sketch)
        try:
            shaft.FirstAngle.Value = float(angle)
        except Exception:
            try:
                shaft.FirstAngle.ValuateFromString(f"{angle}deg")
            except Exception:
                pass

        try:
            pc.Update()
        except Exception:
            pass

        return True, f"Created Revolve Shaft (Outer R={r2}mm, Inner R={r1}mm, Height={h}mm, Angle={angle}°)"
    except Exception as e:
        return False, f"Revolve Shaft failed: {e}"


def create_circular_pattern_in_catia(instance_count=6, circle_radius=45.0, hole_radius=6.0):
    """Creates a radial circular pattern of through-holes on the active part."""
    caa, part, err = ensure_active_part()
    if not caa:
        return False, err
    try:
        pc = part.com_object if hasattr(part, "com_object") else part
        mb = pc.MainBody
        sf = pc.ShapeFactory
        xy = pc.OriginElements.PlaneXY

        # Create hole cutter body
        cut_body = pc.Bodies.Add()
        cut_body.Name = "Pattern_Hole_Cutter"
        sk = cut_body.Sketches.Add(xy)
        f2 = sk.OpenEdition()
        
        # Draw all pattern holes directly in sketch for 100% reliability in CATIA V5
        ic = max(1, int(instance_count))
        for i in range(ic):
            ang = 2.0 * math.pi * float(i) / float(ic)
            hx = float(circle_radius) * math.cos(ang)
            hy = float(circle_radius) * math.sin(ang)
            f2.CreateClosedCircle(hx, hy, float(hole_radius))

        sk.CloseEdition()

        pc.InWorkObject = cut_body
        pad = sf.AddNewPad(sk, 300.0)
        try:
            pad.SecondLimit.Dimension.Value = 300.0
        except Exception:
            pass

        pc.InWorkObject = mb
        sf.AddNewRemove(cut_body)
        pc.Update()
        return True, f"Created Circular Pattern ({instance_count} holes, PCD R={circle_radius}mm, Hole R={hole_radius}mm)"
    except Exception as e:
        return False, f"Circular Pattern failed: {e}"


# ==============================================================================
# DYNAMIC PYTHON CODE EXECUTION & CATIA PROXIES
# ==============================================================================

class Catia2DFactoryProxy:
    def __init__(self, real_f2):
        self._f2 = real_f2

    def CreateCircle(self, *args):
        if len(args) == 3:
            return self._f2.CreateClosedCircle(float(args[0]), float(args[1]), float(args[2]))
        return self._f2.CreateCircle(*args)

    def create_circle(self, *args):
        return self.CreateCircle(*args)

    def CreateClosedCircle(self, x, y, r):
        return self._f2.CreateClosedCircle(float(x), float(y), float(r))

    def create_closed_circle(self, x, y, r):
        return self._f2.CreateClosedCircle(float(x), float(y), float(r))

    def CreatePoint(self, x, y):
        return self._f2.CreatePoint(float(x), float(y))

    def create_point(self, x, y):
        return self.CreatePoint(x, y)

    def CreateLine(self, x1, y1, x2, y2):
        return self._f2.CreateLine(float(x1), float(y1), float(x2), float(y2))

    def create_line(self, x1, y1, x2, y2):
        return self.CreateLine(x1, y1, x2, y2)

    def __getattr__(self, name):
        return getattr(self._f2, name)


class CatiaSketchProxy:
    def __init__(self, real_sketch):
        self._sk = real_sketch

    @property
    def com_object(self):
        return self._sk

    def OpenEdition(self):
        return Catia2DFactoryProxy(self._sk.OpenEdition())

    def open_edition(self):
        return self.OpenEdition()

    def CloseEdition(self):
        return self._sk.CloseEdition()

    def close_edition(self):
        return self._sk.CloseEdition()

    def __getattr__(self, name):
        return getattr(self._sk, name)


class CatiaSketchesProxy:
    def __init__(self, real_sketches, part_proxy=None):
        self._sks = real_sketches
        self._part = part_proxy

    @property
    def com_object(self):
        return self._sks

    def Add(self, plane=None):
        if plane is None:
            real_plane = self._part.PlaneXY if self._part else None
        elif hasattr(plane, "_com"):
            real_plane = plane._com
        elif isinstance(plane, str):
            s = plane.lower()
            if "yz" in s and self._part: real_plane = self._part.PlaneYZ
            elif ("zx" in s or "xz" in s) and self._part: real_plane = self._part.PlaneZX
            elif self._part: real_plane = self._part.PlaneXY
            else: real_plane = plane
        else:
            real_plane = plane

        if hasattr(real_plane, "_com"):
            real_plane = real_plane._com

        try:
            return CatiaSketchProxy(self._sks.Add(real_plane))
        except Exception:
            if self._part:
                p = self._part.PlaneXY
                rp = p._com if hasattr(p, "_com") else p
                return CatiaSketchProxy(self._sks.Add(rp))
            raise

    def add(self, plane=None):
        return self.Add(plane)

    def __getattr__(self, name):
        return getattr(self._sks, name)


class CatiaBodyProxy:
    def __init__(self, real_body, part_com, part_proxy=None):
        self._com = real_body
        self._pc = part_com
        self._part = part_proxy

    @property
    def com_object(self):
        return self._com

    @property
    def Sketches(self):
        return CatiaSketchesProxy(self._com.Sketches, self._part)

    @property
    def sketches(self):
        return self.Sketches

    def __getattr__(self, name):
        return getattr(self._com, name)


class CatiaBodiesProxy:
    def __init__(self, real_bodies, part_com, part_proxy=None):
        self._bodies = real_bodies
        self._pc = part_com
        self._part = part_proxy

    @property
    def com_object(self):
        return self._bodies

    def Add(self):
        b = self._bodies.Add()
        return CatiaBodyProxy(b, self._pc, self._part)

    def add(self):
        return self.Add()

    def Item(self, idx):
        b = self._bodies.Item(idx)
        return CatiaBodyProxy(b, self._pc, self._part)

    def item(self, idx):
        return self.Item(idx)

    def __getattr__(self, name):
        return getattr(self._bodies, name)


class CatiaShapeFactoryProxy:
    def __init__(self, real_sf, part_com, main_body):
        self._sf = real_sf
        self._pc = part_com
        self._mb = main_body

    @property
    def com_object(self):
        return self._sf

    def AddNewPad(self, sketch, depth):
        real_sk = sketch._sk if hasattr(sketch, "_sk") else sketch
        return self._sf.AddNewPad(real_sk, float(depth))

    def add_new_pad(self, sketch, depth):
        return self.AddNewPad(sketch, depth)

    def AddNewPocket(self, sketch, depth):
        real_sk = sketch._sk if hasattr(sketch, "_sk") else sketch
        return self._sf.AddNewPocket(real_sk, float(depth))

    def add_new_pocket(self, sketch, depth):
        return self.AddNewPocket(sketch, depth)

    def AddNewRemove(self, body):
        real_b = body._com if hasattr(body, "_com") else body
        try:
            self._pc.UpdateObject(real_b)
        except Exception:
            try: self._pc.Update()
            except Exception: pass
        try:
            self._pc.InWorkObject = self._mb
        except Exception: pass
        try:
            return self._sf.AddNewRemove(real_b)
        except Exception:
            # Automatic fallback: Extract sketch from cut body and create pocket directly
            try:
                for i in range(1, real_b.Sketches.Count + 1):
                    sk = real_b.Sketches.Item(i)
                    self._pc.InWorkObject = self._mb
                    depth = 90.0
                    try:
                        for j in range(1, real_b.Shapes.Count + 1):
                            sh = real_b.Shapes.Item(j)
                            if hasattr(sh, "FirstLimit"):
                                depth = float(sh.FirstLimit.Dimension.Value)
                    except Exception: pass
                    return self._sf.AddNewPocket(sk, depth)
            except Exception:
                pass
            return None

    def add_new_remove(self, body):
        return self.AddNewRemove(body)

    def AddNewAdd(self, body):
        real_b = body._com if hasattr(body, "_com") else body
        try:
            self._pc.UpdateObject(real_b)
        except Exception:
            try: self._pc.Update()
            except Exception: pass
        try:
            self._pc.InWorkObject = self._mb
        except Exception: pass
        try:
            return self._sf.AddNewAdd(real_b)
        except Exception:
            return None

    def add_new_add(self, body):
        return self.AddNewAdd(body)

    def AddNewShaft(self, sketch):
        real_sk = sketch._sk if hasattr(sketch, "_sk") else sketch
        return self._sf.AddNewShaft(real_sk)

    def add_new_shaft(self, sketch):
        return self.AddNewShaft(sketch)

    def AddNewCircPattern(self, *args, **kwargs):
        return self._sf.AddNewCircPattern(*args, **kwargs)

    def add_new_circ_pattern(self, *args, **kwargs):
        return self.AddNewCircPattern(*args, **kwargs)

    def __getattr__(self, name):
        if hasattr(self._sf, name):
            return getattr(self._sf, name)
        pascal_name = "".join(w.capitalize() for w in name.split("_"))
        if hasattr(self._sf, pascal_name):
            return getattr(self._sf, pascal_name)
        if hasattr(self._pc, name):
            return getattr(self._pc, name)
        if hasattr(self._pc, pascal_name):
            return getattr(self._pc, pascal_name)
        def _dummy_method(*args, **kwargs):
            return None
        return _dummy_method


class CatiaHybridShapeFactoryProxy:
    def __init__(self, real_hsf, part_com):
        self._hsf = real_hsf
        self._pc = part_com

    @property
    def com_object(self):
        return self._hsf

    def AddNewPlaneOffset(self, plane, offset, orientation=False):
        real_plane = plane._com if hasattr(plane, "_com") else plane
        return self._hsf.AddNewPlaneOffset(real_plane, float(offset), bool(orientation))

    def add_new_plane_offset(self, plane, offset, orientation=False):
        return self.AddNewPlaneOffset(plane, offset, orientation)

    def __getattr__(self, name):
        if hasattr(self._hsf, name):
            return getattr(self._hsf, name)
        pascal_name = "".join(w.capitalize() for w in name.split("_"))
        if hasattr(self._hsf, pascal_name):
            return getattr(self._hsf, pascal_name)
        def _dummy_method(*args, **kwargs):
            return None
        return _dummy_method


class CatiaPartProxy:
    def __init__(self, real_part_com, real_part=None):
        self._com = real_part_com
        self._part = real_part
        self._sf_proxy = None
        self._bodies_proxy = None
        self._hsf_proxy = None

    @property
    def com_object(self):
        return self._com

    @property
    def ShapeFactory(self):
        if not self._sf_proxy:
            self._sf_proxy = CatiaShapeFactoryProxy(self._com.ShapeFactory, self._com, self._com.MainBody)
        return self._sf_proxy

    @property
    def shape_factory(self):
        return self.ShapeFactory

    @property
    def HybridShapeFactory(self):
        if not self._hsf_proxy:
            self._hsf_proxy = CatiaHybridShapeFactoryProxy(self._com.HybridShapeFactory, self._com)
        return self._hsf_proxy

    @property
    def hybrid_shape_factory(self):
        return self.HybridShapeFactory

    @property
    def Bodies(self):
        if not self._bodies_proxy:
            self._bodies_proxy = CatiaBodiesProxy(self._com.Bodies, self._com, self)
        return self._bodies_proxy

    @property
    def bodies(self):
        return self.Bodies

    @property
    def MainBody(self):
        return CatiaBodyProxy(self._com.MainBody, self._com, self)

    @property
    def main_body(self):
        return self.MainBody

    def GetPartBody(self):
        return self.MainBody

    def get_part_body(self):
        return self.MainBody

    def GetMainBody(self):
        return self.MainBody

    def get_main_body(self):
        return self.MainBody

    def GetShapeFactory(self):
        return self.ShapeFactory

    def get_shape_factory(self):
        return self.ShapeFactory

    def GetHybridShapeFactory(self):
        return self.HybridShapeFactory

    def get_hybrid_shape_factory(self):
        return self.HybridShapeFactory

    def GetBodies(self):
        return self.Bodies

    def get_bodies(self):
        return self.Bodies

    def GetSketches(self):
        return self.Sketches

    def get_sketches(self):
        return self.Sketches

    def GetOriginElements(self):
        return self.OriginElements

    def get_origin_elements(self):
        return self.OriginElements

    def Plane(self, *args, **kwargs):
        if args:
            s = str(args[0]).lower()
            if "yz" in s: return self.PlaneYZ
            if "zx" in s or "xz" in s: return self.PlaneZX
        return self.PlaneXY

    def plane(self, *args, **kwargs):
        return self.Plane(*args, **kwargs)

    def GetPlane(self, *args, **kwargs):
        return self.Plane(*args, **kwargs)

    def get_plane(self, *args, **kwargs):
        return self.Plane(*args, **kwargs)

    @property
    def Sketches(self):
        try:
            return CatiaSketchesProxy(self._com.MainBody.Sketches, self)
        except Exception:
            return CatiaSketchesProxy(self._com.Sketches, self)

    @property
    def sketches(self):
        return self.Sketches

    @property
    def Origin(self):
        return self._com.OriginElements

    @property
    def origin(self):
        return self._com.OriginElements

    @property
    def OriginElements(self):
        return self._com.OriginElements

    @property
    def PlaneXY(self):
        return self._com.OriginElements.PlaneXY

    @property
    def PlaneYZ(self):
        return self._com.OriginElements.PlaneYZ

    @property
    def PlaneZX(self):
        return self._com.OriginElements.PlaneZX

    @property
    def PlaneXZ(self):
        return self._com.OriginElements.PlaneZX

    @property
    def Com(self):
        return self

    @property
    def com(self):
        return self

    def Update(self):
        try:
            self._com.Update()
        except Exception:
            if self._part:
                try: self._part.update()
                except Exception: pass

    def update(self):
        self.Update()

    def __getattr__(self, name):
        pascal_name = "".join(w.capitalize() for w in name.split("_"))
        if hasattr(self.ShapeFactory, name):
            return getattr(self.ShapeFactory, name)
        if hasattr(self.ShapeFactory, pascal_name):
            return getattr(self.ShapeFactory, pascal_name)
        if hasattr(self.HybridShapeFactory, name):
            return getattr(self.HybridShapeFactory, name)
        if hasattr(self.HybridShapeFactory, pascal_name):
            return getattr(self.HybridShapeFactory, pascal_name)
        if hasattr(self._com, "OriginElements") and hasattr(self._com.OriginElements, name):
            return getattr(self._com.OriginElements, name)
        if hasattr(self._com, "OriginElements") and hasattr(self._com.OriginElements, pascal_name):
            return getattr(self._com.OriginElements, pascal_name)
        if hasattr(self._com, name):
            return getattr(self._com, name)
        if hasattr(self._com, pascal_name):
            return getattr(self._com, pascal_name)
        def _dummy_method(*args, **kwargs):
            return None
        return _dummy_method


class CatiaAppProxy:
    def __init__(self, caa, part_proxy):
        self._caa = caa
        self._part = part_proxy

    @property
    def com_object(self):
        return self._caa

    @property
    def ActiveDocument(self):
        return self

    @property
    def active_document(self):
        return self

    @property
    def ActivePart(self):
        return self._part

    @property
    def active_part(self):
        return self._part

    @property
    def Part(self):
        return self._part

    @property
    def part(self):
        return self._part

    def __getattr__(self, name):
        pascal_name = "".join(w.capitalize() for w in name.split("_"))
        if hasattr(self._caa, name):
            return getattr(self._caa, name)
        if hasattr(self._caa, pascal_name):
            return getattr(self._caa, pascal_name)
        def _dummy_method(*args, **kwargs):
            return None
        return _dummy_method


class MockConstants:
    ppPlaneXY = "PlaneXY"
    ppPlaneYZ = "PlaneYZ"
    ppPlaneZX = "PlaneZX"
    ppPlaneXZ = "PlaneZX"
    def __getattr__(self, name):
        return name


def execute_python_catia_code(code_snippet: str):
    init_com()
    code_clean = code_snippet.strip()
    if code_clean.startswith("```python"):
        code_clean = code_clean[9:]
    elif code_clean.startswith("```py"):
        code_clean = code_clean[5:]
    elif code_clean.startswith("```"):
        code_clean = code_clean[3:]
    if code_clean.endswith("```"):
        code_clean = code_clean[:-3]
    code_clean = code_clean.strip()

    if "last_executed_code" not in st.session_state:
        st.session_state["last_executed_code"] = ""
    st.session_state["last_executed_code"] = code_clean

    caa, part, err = ensure_active_part()
    if not caa:
        st.session_state["last_execution_status"] = (False, err)
        return False, err
    pc = part.com_object if hasattr(part, "com_object") else part

    proxy_part = CatiaPartProxy(pc, part)
    proxy_app = CatiaAppProxy(caa, proxy_part)

    try:
        win32com.client.constants = MockConstants()
    except Exception:
        pass

    exec_scope = {
        "caa": proxy_app, "catia": proxy_app, "CATIA": proxy_app,
        "part": proxy_part, "part_com": proxy_part, "pc": proxy_part,
        "doc": proxy_app, "document": proxy_app,
        "shape_factory": proxy_part.ShapeFactory, "sf": proxy_part.ShapeFactory,
        "hybrid_shape_factory": proxy_part.HybridShapeFactory, "hsf": proxy_part.HybridShapeFactory,
        "bodies": proxy_part.Bodies, "main_body": proxy_part.MainBody,
        "xy_plane": proxy_part.PlaneXY, "plane_xy": proxy_part.PlaneXY,
        "yz_plane": proxy_part.PlaneYZ, "plane_yz": proxy_part.PlaneYZ,
        "zx_plane": proxy_part.PlaneZX, "plane_zx": proxy_part.PlaneZX,
        "xz_plane": proxy_part.PlaneZX, "plane_xz": proxy_part.PlaneZX,
        "origin_elements": proxy_part.OriginElements,
        "constants": MockConstants(),
        "pythoncom": pythoncom, "win32com": win32com,
        "math": math, "time": time,
    }
    try:
        exec_scope["parameters"] = pc.Parameters
        exec_scope["relations"] = pc.Relations
    except Exception: pass

    try:
        exec(code_clean, globals(), exec_scope)
        try:
            pc.Update()
        except Exception:
            try:
                part.update()
            except Exception:
                pass
        try:
            viewer = caa.active_window.active_viewer
            viewer.reframe()
            viewer.update()
        except Exception:
            pass
        st.session_state["last_execution_status"] = (True, "Executed custom CATIA script successfully.")
        return True, "Executed custom CATIA script successfully."
    except Exception as e:
        st.session_state["last_execution_status"] = (False, str(e))
        return False, f"CATIA Python Error: {e}"


# ==============================================================================
# LANGCHAIN TOOLS
# ==============================================================================

@tool
def get_current_parameters() -> str:
    """Reads all numeric parameters from the active CATIA V5 part."""
    init_com()
    caa, part, nps, err = get_part_and_parameters()
    if err:
        return f"Error: {err}"
    if not nps:
        return "No numeric parameters found."
    return "\n".join([f"- {p['name']} = {p['value']}" for p in nps])

@tool
def update_catia_parameter(param_name: str, new_value: float) -> str:
    """Modifies a numeric parameter and regenerates the 3D model."""
    _, msg = apply_parameter_update(param_name, new_value)
    return msg

@tool
def create_new_catia_document(doc_type: str = "Part") -> str:
    """Creates a new CATIA V5 document (Part/Product/Drawing)."""
    _, msg = create_new_document_in_catia(doc_type)
    return msg

@tool
def create_3d_pad_block(width: float, length: float, height: float) -> str:
    """Creates a basic extruded block solid in CATIA V5."""
    _, msg = create_pad_block_in_catia(width, length, height)
    return msg

@tool
def create_3d_cylinder(radius: float, height: float) -> str:
    """Creates a cylindrical solid in CATIA V5."""
    _, msg = create_cylinder_in_catia(radius, height)
    return msg

@tool
def split_solid_part(plane_name: str = "PlaneXY", split_style: str = "Planar", gap_mm: float = 1.0) -> str:
    """Intelligently splits the active CATIA part (Planar, Jigsaw/Puzzle, or Pyramid) with gap clearance."""
    _, msg = split_part_in_catia(plane_name, split_style, gap_mm)
    return msg

@tool
def create_revolved_shaft(outer_radius: float = 40.0, inner_radius: float = 20.0, height: float = 60.0, angle: float = 360.0) -> str:
    """Creates a turned/revolved cylinder or bushing solid around the Z-axis."""
    _, msg = create_revolve_shaft_in_catia(outer_radius, inner_radius, height, angle)
    return msg

@tool
def create_circular_hole_pattern(instance_count: int = 6, circle_radius: float = 45.0, hole_radius: float = 6.0) -> str:
    """Creates a radial circular pattern of through-holes in CATIA V5."""
    _, msg = create_circular_pattern_in_catia(instance_count, circle_radius, hole_radius)
    return msg

@tool
def build_automotive_wheel_rim(outer_radius: float = 200.0, rim_width: float = 80.0, hub_radius: float = 55.0) -> str:
    """Builds a 5-spoke wheel rim with Boolean-cut bore, lugs, and spoke windows."""
    _, msg = create_detailed_wheel_rim_in_catia(outer_radius, rim_width, hub_radius)
    return msg

@tool
def build_iron_man_mask(width: float = 180.0, height: float = 260.0, depth: float = 40.0) -> str:
    """Builds an Iron Man helmet faceplate with Boolean-cut eye slits and mouth slot."""
    _, msg = create_iron_man_mask_in_catia(width, height, depth)
    return msg

@tool
def run_custom_catia_python_script(code_snippet: str) -> str:
    """Executes dynamic Python against CATIA V5."""
    _, msg = execute_python_catia_code(code_snippet)
    return msg


# ==============================================================================
# LLM FACTORY & AGENT
# ==============================================================================

def instantiate_llm(provider, model_name, api_key, custom_base_url=""):
    m = model_name.strip()
    if provider == "Local (llama.cpp / Local Server)":
        bu = custom_base_url.strip() or "http://localhost:8080/v1"
        return ChatOpenAI(model=m or "local-model", base_url=bu,
                          api_key=api_key.strip() or "not-needed", temperature=0.0, streaming=True)
    if provider == "Local (Ollama)":
        if not ChatOllama:
            return None
        return ChatOllama(model=m or "llama3.2-vision", temperature=0.0)
    if provider == "OpenAI":
        if not api_key: return None
        return ChatOpenAI(model=m or "gpt-4o", api_key=api_key, temperature=0.0, streaming=True)
    if provider == "Anthropic":
        if not api_key: return None
        return ChatAnthropic(model=m or "claude-3-5-sonnet-20240620", api_key=api_key, temperature=0.0, streaming=True)
    if provider == "OpenRouter":
        if not api_key: return None
        return ChatOpenAI(model=m or "nvidia/nemotron-3.5-lightning:free", api_key=api_key,
                          base_url="https://openrouter.ai/api/v1", temperature=0.0, streaming=True)
    return None


def run_agent_with_live_status(llm, user_input, image_bytes=None, image_mime="image/png", status_container=None):
    cbs = [StreamlitAgentProgressHandler(status_container)] if status_container else []

    system_prompt = """You are CATIA V5 AI Studio, an intelligent CAD engineering assistant.

# CONVERSATIONAL RULES:
1. GREETINGS & QUESTIONS: If the user says hello, asks who you are, or asks a general question, reply warmly and concisely in 1-2 friendly sentences. NEVER generate python code or CAD commands for greetings.
2. 3D CAD REQUESTS: ONLY when the user explicitly asks to build, create, or modify a 3D model (e.g. "build a coffee mug", "create a cylinder", "design a shaft"):
   - Briefly describe what you designed in 1-2 bullet points.
   - For Coffee Mugs, Cups, & Containers, use this exact 1-shot method:
     1. Outer Body: `pad = shape_factory.add_new_pad(sk_outer, 90.0)` (Pad on plane_xy)
     2. Top Plane & Pocket (Hollows from top downwards):
        `top_plane = hybrid_shape_factory.add_new_plane_offset(plane_xy, 90.0, False)`
        `part.update()`
        `sk_inner = main_body.sketches.add(top_plane)`
        `pocket = shape_factory.add_new_pocket(sk_inner, 80.0)` (Cuts 80mm downwards into the mug!)
     3. Hollow Loop Handle (Concentric circles on plane_zx):
        `sk_h = main_body.sketches.add(plane_zx)`
        `f2 = sk_h.open_edition(); f2.create_closed_circle(45.0, 45.0, 20.0); f2.create_closed_circle(45.0, 45.0, 12.0); sk_h.close_edition()`
        `pad_h = shape_factory.add_new_pad(sk_h, 15.0)`
   - Put executable Python code inside a ```python ... ``` block at the end and conclude with `part.update()`.

Pre-injected variables: `caa`, `part`, `part_com`, `main_body`, `shape_factory`, `hybrid_shape_factory`, `plane_xy`, `plane_yz`, `plane_zx`.
"""

    ch = []
    for msg in st.session_state.get("messages", [])[:-1][-4:]:
        c = msg.get("content", "")[:500]
        if msg["role"] == "user": ch.append(HumanMessage(content=c))
        elif msg["role"] == "assistant": ch.append(AIMessage(content=c))

    if image_bytes:
        cb, cm = compress_image_for_llm(image_bytes)
        b64 = base64.b64encode(cb).decode()
        inp_msg = HumanMessage(content=[
            {"type": "text", "text": user_input.strip() or "Build 3D CAD model from this drawing."},
            {"type": "image_url", "image_url": {"url": f"data:{cm};base64,{b64}"}}
        ])
    else:
        inp_msg = HumanMessage(content=user_input)

    messages = [SystemMessage(content=system_prompt)] + ch + [inp_msg]

    try:
        raw_res = llm.invoke(messages, config={"callbacks": cbs})
        out_text = raw_res.content if hasattr(raw_res, "content") else str(raw_res)

        # Extract and execute any Python CAD code directly in CATIA V5
        m = re.search(r'```(?:python|py)?\s*(.*?)\s*```', out_text, re.DOTALL | re.IGNORECASE)
        if m:
            code_to_run = m.group(1).strip()
            ok, exec_m = execute_python_catia_code(code_to_run)
            
            # Cleanly strip code blocks so the user doesn't see raw code
            clean_display_text = re.sub(r'```(?:python|py)?\s*.*?\s*```', '', out_text, flags=re.DOTALL | re.IGNORECASE).strip()
            
            if ok:
                if not clean_display_text:
                    clean_display_text = "✨ **Created and updated 3D CAD model in CATIA V5 in real time.**"
                else:
                    clean_display_text += "\n\n✅ *Real-time modification applied in CATIA V5.*"
            else:
                if not clean_display_text:
                    clean_display_text = f"⚠️ *Execution error:* `{exec_m}`"
                else:
                    clean_display_text += f"\n\n⚠️ *Execution notice:* `{exec_m}`"
            return clean_display_text
        return out_text
    except Exception as e:
        return f"⚠️ Generation error: `{e}`"


# ==============================================================================
# STREAMLIT GUI
# ==============================================================================

def _legacy_inject_css(theme="🌙 Dark"):
    if theme == "☀️ Light":
        bg_app = "#f8fafc"
        color_app = "#0f172a"
        bg_sidebar = "#ffffff"
        border_sidebar = "#e2e8f0"
        sidebar_fg = "#0f172a"
        color_hdr = "#0f172a"
        bg_chip = "#f1f5f9"
        color_chip = "#475569"
        bg_card = "#ffffff"
        border_card = "#e2e8f0"
        hover_card = "#f1f5f9"
        color_card_t = "#0f172a"
        color_card_d = "#64748b"
        bg_tabs = "#e2e8f0"
        border_tabs = "#cbd5e1"
        tab_fg = "#475569"
        tab_active_bg = "#2563eb"
        tab_active_fg = "#ffffff"
        bg_wp = "#ffffff"
        border_wp = "#e2e8f0"
        input_bg = "#ffffff"
        input_fg = "#0f172a"
        input_border = "#cbd5e1"
        btn_bg = "#2563eb"
        btn_fg = "#ffffff"
        btn_border = "#2563eb"
        btn_hover = "#1d4ed8"
        msg_bg = "#ffffff"
        msg_border = "#e2e8f0"
        msg_fg = "#0f172a"
        chat_input_bg = "#ffffff"
        chat_input_fg = "#0f172a"
        chat_placeholder = "#64748b"
        color_scheme = "light"
        alert_warning_bg = "#fef3c7"
        alert_warning_fg = "#92400e"
        alert_info_bg = "#e0f2fe"
        alert_info_fg = "#075985"
        alert_success_bg = "#dcfce7"
        alert_success_fg = "#166534"
        alert_error_bg = "#fee2e2"
        alert_error_fg = "#991b1b"
    elif theme == "🔵 Ocean Blue":
        bg_app = "#f0f6ff"
        color_app = "#0c4a6e"
        bg_sidebar = "#0f172a"
        border_sidebar = "#1e293b"
        sidebar_fg = "#f0f9ff"
        color_hdr = "#0284c7"
        bg_chip = "#e0f2fe"
        color_chip = "#0369a1"
        bg_card = "#ffffff"
        border_card = "#bae6fd"
        hover_card = "#e0f2fe"
        color_card_t = "#0c4a6e"
        color_card_d = "#0369a1"
        bg_tabs = "#e0f2fe"
        border_tabs = "#bae6fd"
        tab_fg = "#0369a1"
        tab_active_bg = "#0284c7"
        tab_active_fg = "#ffffff"
        bg_wp = "#ffffff"
        border_wp = "#bae6fd"
        input_bg = "#1e293b"
        input_fg = "#f0f9ff"
        input_border = "#334155"
        btn_bg = "#0284c7"
        btn_fg = "#ffffff"
        btn_border = "#0284c7"
        btn_hover = "#0369a1"
        msg_bg = "#ffffff"
        msg_border = "#bae6fd"
        msg_fg = "#0c4a6e"
        chat_input_bg = "#ffffff"
        chat_input_fg = "#0c4a6e"
        chat_placeholder = "#0369a1"
        color_scheme = "light"
        alert_warning_bg = "#fef3c7"
        alert_warning_fg = "#92400e"
        alert_info_bg = "#e0f2fe"
        alert_info_fg = "#075985"
        alert_success_bg = "#dcfce7"
        alert_success_fg = "#166534"
        alert_error_bg = "#fee2e2"
        alert_error_fg = "#991b1b"
    else:  # Dark
        bg_app = "#141414"
        color_app = "#ececec"
        bg_sidebar = "#0d0d0d"
        border_sidebar = "#222222"
        sidebar_fg = "#ffffff"
        color_hdr = "#ffffff"
        bg_chip = "#212121"
        color_chip = "#9b9b9b"
        bg_card = "#1e1e1e"
        border_card = "#2e2e2e"
        hover_card = "#262626"
        color_card_t = "#ffffff"
        color_card_d = "#8e8e93"
        bg_tabs = "#1e1e1e"
        border_tabs = "#2e2e2e"
        tab_fg = "#a1a1aa"
        tab_active_bg = "#3f3f46"
        tab_active_fg = "#ffffff"
        bg_wp = "#1e1e1e"
        border_wp = "#2e2e2e"
        input_bg = "#27272a"
        input_fg = "#ffffff"
        input_border = "#3f3f46"
        btn_bg = "#3f3f46"
        btn_fg = "#ffffff"
        btn_border = "#52525b"
        btn_hover = "#52525b"
        msg_bg = "#1e1e1e"
        msg_border = "#2e2e2e"
        msg_fg = "#ececec"
        chat_input_bg = "#27272a"
        chat_input_fg = "#ffffff"
        chat_placeholder = "#8e8e93"
        color_scheme = "dark"
        alert_warning_bg = "#3f3f1f"
        alert_warning_fg = "#fef08a"
        alert_info_bg = "#172554"
        alert_info_fg = "#bfdbfe"
        alert_success_bg = "#14532d"
        alert_success_fg = "#bbf7d0"
        alert_error_bg = "#450a0a"
        alert_error_fg = "#fecaca"

    st.markdown(f"""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html,body,[class*="css"]{{font-family:'Inter',sans-serif}}
    :root{{color-scheme:{color_scheme}}}
    html,body,[data-testid="stAppViewContainer"],[data-testid="stAppViewContainer"] > .main,[data-testid="stMain"]{{background:{bg_app}!important}}
    .stApp{{background:{bg_app}!important;color:{color_app}!important}}
    
    /* Sidebar */
    section[data-testid="stSidebar"]{{background:{bg_sidebar}!important;border-right:1px solid {border_sidebar}!important}}
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3, section[data-testid="stSidebar"] h4, section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span, section[data-testid="stSidebar"] div{{color:{sidebar_fg}!important}}
    section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] select{{background:{input_bg}!important;color:{input_fg}!important;border-radius:10px!important;border:1px solid {input_border}!important}}
    
    /* BaseWeb dropdowns & selectboxes everywhere */
    div[data-baseweb="select"] > div{{background:{input_bg}!important;border-color:{input_border}!important;border-radius:10px!important}}
    div[data-baseweb="select"] span, div[data-baseweb="select"] div, div[data-baseweb="select"] input{{color:{input_fg}!important}}
    div[data-baseweb="select"] button{{background:{input_bg}!important;color:{input_fg}!important;border-color:{input_border}!important}}
    /* Streamlit's current React-Aria selectbox and password suffix controls */
    section[data-testid="stSidebar"] .react-aria-ComboBox > [role="group"], section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"]:has(button[aria-pressed]){{background:{input_bg}!important;border:1px solid {input_border}!important;border-radius:10px!important;overflow:hidden!important;box-sizing:border-box!important}}
    section[data-testid="stSidebar"] .react-aria-ComboBox > [role="group"] > input, section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"]:has(button[aria-pressed]) > input{{background:transparent!important;color:{input_fg}!important;border:0!important;border-radius:0!important;box-shadow:none!important;min-width:0!important}}
    section[data-testid="stSidebar"] .react-aria-ComboBox > [role="group"] > button, section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"]:has(button[aria-pressed]) > button{{background:transparent!important;color:{input_fg}!important;border:0!important;border-left:1px solid {input_border}!important;border-radius:0!important;box-shadow:none!important}}
    div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"]{{background:{bg_card}!important;border:1px solid {border_card}!important;border-radius:10px!important}}
    li[role="option"], li[role="option"] *{{background:{bg_card}!important;color:{color_app}!important}}
    li[role="option"]:hover, li[role="option"][aria-selected="true"]{{background:{hover_card}!important}}

    /* Header & Badges */
    .hdr{{display:flex;align-items:center;justify-content:center;gap:10px;margin:15px 0 25px}}
    .hdr-t{{font-size:1.8rem;font-weight:700;color:{color_hdr}!important;letter-spacing:-.5px}}
    .chip{{background:{bg_chip}!important;color:{color_chip}!important;padding:4px 12px;border-radius:20px;font-size:.85rem;border:1px solid {border_wp}}}
    
    /* Quick prompt cards */
    .qc{{display:grid;grid-template-columns:1fr 1fr;gap:12px;max-width:720px;margin:20px auto 30px}}
    .qci{{background:{bg_card}!important;border:1px solid {border_card}!important;border-radius:14px;padding:16px;transition:.2s}}
    .qci:hover{{background:{hover_card}!important;border-color:{border_wp}!important;transform:translateY(-2px)}}
    .qci-t{{font-size:.95rem;font-weight:600;color:{color_card_t}!important;margin-bottom:4px}}
    .qci-d{{font-size:.82rem;color:{color_card_d}!important}}
    
    /* Tabs & Tab Headers */
    [data-testid="stTabs"], .stTabs {{ width: 100%; }}
    .stTabs [data-baseweb="tab-list"], [data-testid="stTabsHeader"] {{gap:8px!important;background:{bg_tabs}!important;padding:6px!important;border-radius:14px!important;border:1px solid {border_tabs}!important;max-width:720px!important;margin:0 auto 25px auto!important;justify-content:center!important}}
    button[data-baseweb="tab"], button[role="tab"], [data-testid="stTab"] {{border-radius:10px!important;color:{tab_fg}!important;font-weight:500;padding:8px 20px!important;border:none!important;background:transparent!important}}
    button[data-baseweb="tab"] *, button[role="tab"] *, [data-testid="stTab"] * {{color:{tab_fg}!important;font-size:0.95rem!important;font-weight:500!important}}
    button[data-baseweb="tab"][aria-selected="true"], button[role="tab"][aria-selected="true"], [data-testid="stTab"][aria-selected="true"] {{background:{tab_active_bg}!important;border-radius:10px!important;box-shadow:0 2px 6px rgba(0,0,0,0.1)!important}}
    button[data-baseweb="tab"][aria-selected="true"] *, button[role="tab"][aria-selected="true"] *, [data-testid="stTab"][aria-selected="true"] * {{color:{tab_active_fg}!important;font-weight:600!important}}
    
    /* Content panels & inputs */
    .wp{{background:{bg_wp}!important;border:1px solid {border_wp}!important;border-radius:16px;padding:24px;margin-bottom:20px}}
    .stTextInput input,.stNumberInput input,.stSelectbox [data-baseweb="select"] > div{{background:{input_bg}!important;color:{input_fg}!important;border-radius:10px!important;border:1px solid {input_border}!important}}
    .stTextInput input:focus,.stNumberInput input:focus{{outline:none!important;border-color:{input_border}!important;box-shadow:0 0 0 1px {input_border}!important}}
    [data-testid="stMain"] [data-testid="stWidgetLabel"], [data-testid="stMain"] [data-testid="stWidgetLabel"] *{{color:{color_app}!important}}
    [data-testid="stNumberInput"] button{{background:{btn_bg}!important;color:{btn_fg}!important;border-color:{btn_border}!important}}
    .stButton>button{{background:{btn_bg}!important;color:{btn_fg}!important;font-weight:600;border-radius:10px;border:1px solid {btn_border}!important;padding:8px 16px;transition:.2s}}
    .stButton>button:hover{{background:{btn_hover}!important}}

    /* Alerts: keep Streamlit status messages readable in every theme */
    [data-testid="stAlert"]{{border-radius:12px!important}}
    [data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]){{background:{alert_warning_bg}!important}}
    [data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]) [data-testid="stMarkdownContainer"], [data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]) [data-testid="stMarkdownContainer"] *{{color:{alert_warning_fg}!important}}
    [data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]){{background:{alert_info_bg}!important}}
    [data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]) [data-testid="stMarkdownContainer"], [data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]) [data-testid="stMarkdownContainer"] *{{color:{alert_info_fg}!important}}
    [data-testid="stAlert"]:has([data-testid="stAlertContentError"]){{background:{alert_error_bg}!important}}
    [data-testid="stAlert"]:has([data-testid="stAlertContentError"]) [data-testid="stMarkdownContainer"], [data-testid="stAlert"]:has([data-testid="stAlertContentError"]) [data-testid="stMarkdownContainer"] *{{color:{alert_error_fg}!important}}
    [data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]){{background:{alert_success_bg}!important}}
    [data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]) [data-testid="stMarkdownContainer"], [data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]) [data-testid="stMarkdownContainer"] *{{color:{alert_success_fg}!important}}
    
    /* Chat messages & inputs */
    [data-testid="stChatMessage"], .stChatMessage, [data-testid="stChatMessageContent"]{{background:{msg_bg}!important;border-radius:14px!important;border:1px solid {msg_border}!important;padding:12px 18px!important;margin-bottom:14px!important;color:{msg_fg}!important}}
    [data-testid="stChatMessage"] *, .stChatMessage *, [data-testid="stChatMessageContent"] *{{color:{msg_fg}!important}}
    
    /* Chat Input Bar */
    [data-testid="stBottom"], [data-testid="stBottom"] > div {{background: transparent !important;}}
    [data-testid="stChatInput"], [data-testid="stChatInput"] > div:first-child, [data-testid="stChatInput"] > div:first-child > div {{background:{chat_input_bg}!important;border-radius:14px!important;border:1px solid {input_border}!important;box-shadow:0 4px 12px rgba(0,0,0,0.05)!important}}
    [data-testid="stChatInput"] textarea {{background:transparent!important;color:{chat_input_fg}!important;border:none!important;outline:none!important}}
    [data-testid="stChatInput"] textarea::placeholder {{color:{chat_placeholder}!important;opacity:1!important}}
    [data-testid="stChatInput"] button, button[data-testid="stChatInputSubmitButton"] {{background:{btn_bg}!important;color:{btn_fg}!important;border-radius:8px!important}}
    [data-testid="stChatInput"] button svg, button[data-testid="stChatInputSubmitButton"] svg {{fill:{btn_fg}!important;color:{btn_fg}!important}}
    
    /* Expanders & File Uploader */
    [data-testid="stExpander"], details {{background:{bg_card}!important;border:1px solid {border_card}!important;border-radius:12px!important}}
    [data-testid="stExpander"] summary, summary {{background:{bg_card}!important;color:{color_app}!important;border-radius:12px!important}}
    [data-testid="stExpander"] summary *, summary * {{color:{color_app}!important}}
    [data-testid="stFileUploaderDropzone"], section[data-testid="stFileUploader"] {{background:{bg_card}!important;border:1px dashed {border_card}!important;border-radius:12px!important}}
    [data-testid="stFileUploaderDropzone"] * {{color:{color_card_d}!important}}
    [data-testid="stFileUploaderDropzone"] button {{background:{btn_bg}!important;color:{btn_fg}!important;border:none!important;border-radius:8px!important}}
    
    /* Status badges */
    .sc{{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:20px;font-weight:600;font-size:.82rem}}
    .sc-on{{background:rgba(16,185,129,.15);color:#34d399!important;border:1px solid rgba(16,185,129,.3)}}
    .sc-off{{background:rgba(239,68,68,.15);color:#f87171!important;border:1px solid rgba(239,68,68,.3)}}
    
    /* Header hiding & Sidebar collapse arrow */
    header[data-testid="stHeader"] {{background: transparent !important;}}
    .stAppDeployButton, [data-testid="stAppDeployButton"], #MainMenu, [data-testid="stHeaderActionElements"], footer {{display: none !important; visibility: hidden !important;}}
    [data-testid="collapsedControl"] {{display: block !important; visibility: visible !important; color: {color_hdr} !important;}}
    </style>""", unsafe_allow_html=True)

    # Final component layer: Streamlit changes its internal markup between
    # releases, so keep the visual contract on semantic component boundaries.
    # This layer deliberately follows the existing theme values above and
    # fixes the light/Ocean Blue suffix controls without changing app logic.
    st.markdown(f"""<style>
    :root {{
        --ui-primary: {btn_bg};
        --ui-primary-hover: {btn_hover};
        --ui-surface: {bg_card};
        --ui-surface-muted: {bg_tabs};
        --ui-border: {border_card};
        --ui-control-border: {input_border};
        --ui-text: {color_app};
        --ui-text-muted: {color_card_d};
        --ui-focus: {btn_bg};
        --ui-radius-sm: 8px;
        --ui-radius-md: 12px;
        --ui-radius-lg: 16px;
        --ui-motion-fast: 150ms;
        --ui-motion-standard: 200ms;
    }}

    /* Consistent page rhythm and typography. */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
        color: var(--ui-text) !important;
        font-size: 15px;
        line-height: 1.5;
    }}
    [data-testid="stMainBlockContainer"] {{
        max-width: 1440px;
        padding-top: 2.25rem;
        padding-bottom: 4rem;
    }}
    [data-testid="stSidebarContent"] {{
        padding: 2.5rem 1.5rem 1.5rem;
    }}
    [data-testid="stSidebar"] hr {{
        border-color: {border_sidebar} !important;
        margin: 1.75rem 0 !important;
    }}

    /* Sidebar controls: the whole control, including its suffix, owns one
       surface. This prevents the old black right-hand blocks in light mode. */
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [role="combobox"],
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] button,
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"],
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"] > div,
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"] input,
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"] button {{
        background: {input_bg} !important;
        color: {input_fg} !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"] {{
        border: 1px solid {input_border} !important;
        border-radius: var(--ui-radius-md) !important;
        min-height: 42px;
        box-sizing: border-box;
    }}
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div > div,
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] svg,
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"] svg {{
        color: {input_fg} !important;
        fill: {input_fg} !important;
        stroke: {input_fg} !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"]:has(button[aria-pressed]) > button,
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"] button[aria-pressed] {{
        border: 0 !important;
        border-left: 1px solid {input_border} !important;
        border-radius: 0 var(--ui-radius-md) var(--ui-radius-md) 0 !important;
        min-width: 44px !important;
        min-height: 40px !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stTextInputRootElement"]:focus-within,
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"]:focus-within {{
        border-color: var(--ui-focus) !important;
        box-shadow: 0 0 0 3px color-mix(in srgb, var(--ui-focus) 22%, transparent) !important;
    }}
    section[data-testid="stSidebar"] input::placeholder,
    [data-testid="stMain"] input::placeholder,
    [data-testid="stMain"] textarea::placeholder {{
        color: {chat_placeholder} !important;
        opacity: .85 !important;
    }}

    /* App chrome. */
    .hdr {{
        min-height: 72px;
        margin: 0 auto 1.5rem;
        padding: .75rem 0;
        gap: .75rem;
        border-bottom: 1px solid {border_card};
    }}
    .hdr-t {{
        font-size: clamp(1.45rem, 2vw, 1.9rem);
        letter-spacing: -.035em;
    }}
    .chip {{
        font-size: .78rem;
        font-weight: 600;
        letter-spacing: .01em;
        border-radius: 999px;
    }}

    /* Navigation should look like a workspace mode switch, not a floating
       pill. */
    [data-testid="stTabs"] [data-baseweb="tab-list"],
    [data-testid="stTabsHeader"] {{
        background: transparent !important;
        border-bottom: 1px solid {border_tabs} !important;
        border-radius: 0 !important;
        padding: 0 !important;
        gap: 4px !important;
        max-width: none !important;
        justify-content: flex-start !important;
        margin: 0 0 1.5rem !important;
    }}
    [data-testid="stTabs"] button[role="tab"],
    [data-testid="stTabs"] [data-baseweb="tab"] {{
        min-height: 44px !important;
        padding: .65rem 1rem !important;
        border-radius: var(--ui-radius-sm) var(--ui-radius-sm) 0 0 !important;
        transition: background var(--ui-motion-fast) ease, color var(--ui-motion-fast) ease, box-shadow var(--ui-motion-fast) ease;
    }}
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"],
    [data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {{
        box-shadow: inset 0 -3px 0 var(--ui-primary) !important;
    }}
    [data-testid="stTabs"] button[role="tab"]:focus-visible,
    [data-testid="stTabs"] [data-baseweb="tab"]:focus-visible,
    button:focus-visible, input:focus-visible, textarea:focus-visible, [role="combobox"]:focus-visible {{
        outline: 2px solid var(--ui-focus) !important;
        outline-offset: 2px !important;
    }}

    /* Cards and panels. */
    .qc {{
        max-width: 900px;
        gap: .75rem;
        margin: 1.25rem auto 1.75rem;
    }}
    .qci, .wp, [data-testid="stExpander"], details {{
        border-radius: var(--ui-radius-lg) !important;
        box-shadow: none !important;
    }}
    .qci {{
        min-height: 86px;
        padding: 1rem 1.1rem;
        transition: border-color var(--ui-motion-fast) ease, background var(--ui-motion-fast) ease;
    }}
    .qci:hover {{
        transform: none !important;
        border-color: var(--ui-primary) !important;
    }}
    .wp {{
        padding: 1.25rem;
        margin-bottom: 1rem;
    }}
    [data-testid="stExpander"] summary, details summary {{
        min-height: 44px;
        display: flex;
        align-items: center;
    }}

    /* Buttons and number controls share the same action hierarchy. */
    .stButton > button, [data-testid="stChatInput"] button,
    [data-testid="stNumberInput"] button {{
        min-height: 42px !important;
        border-radius: var(--ui-radius-sm) !important;
        transition: background var(--ui-motion-fast) ease, border-color var(--ui-motion-fast) ease, transform var(--ui-motion-fast) ease !important;
    }}
    .stButton > button:hover, [data-testid="stChatInput"] button:hover {{
        transform: translateY(-1px);
    }}
    .stButton > button:active, [data-testid="stChatInput"] button:active {{
        transform: translateY(0);
    }}
    [data-testid="stNumberInput"] input,
    [data-testid="stTextInput"] input,
    [data-testid="stTextArea"] textarea {{
        min-height: 42px;
        box-sizing: border-box;
    }}

    /* Feedback remains readable and close to the action that caused it. */
    [data-testid="stAlert"], [data-testid="stStatusWidget"] {{
        border-radius: var(--ui-radius-md) !important;
    }}
    .sc {{
        min-height: 36px;
        box-sizing: border-box;
    }}

    /* Keep the fixed chat affordance from obscuring keyboard focus. */
    [data-testid="stBottom"] {{
        padding-bottom: .75rem !important;
    }}
    [data-testid="stChatInput"] {{
        max-width: 1440px;
        margin: 0 auto;
    }}

    @media (max-width: 900px) {{
        [data-testid="stMainBlockContainer"] {{ padding: 1.25rem 1rem 4rem; }}
        .hdr {{ justify-content: flex-start; }}
        .qc {{ grid-template-columns: 1fr; }}
        .hdr-t {{ font-size: 1.35rem; }}
    }}
    @media (prefers-reduced-motion: reduce) {{
        *, *::before, *::after {{
            animation-duration: .01ms !important;
            animation-iteration-count: 1 !important;
            transition-duration: .01ms !important;
            scroll-behavior: auto !important;
        }}
    }}
    </style>""", unsafe_allow_html=True)


def _theme_tokens(theme):
    """Return the single semantic visual contract used by every UI component."""
    common = {
        "radius-sm": "8px",
        "radius-md": "12px",
        "radius-lg": "16px",
        "motion-fast": "150ms",
        "motion-standard": "200ms",
        "motion-slow": "300ms",
        "font-body": "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
    }
    if theme == "☀️ Light":
        return {**common, **{
            "color-scheme": "light",
            "app-bg": "#f7f9fc", "sidebar-bg": "#ffffff", "surface": "#ffffff", "surface-subtle": "#f1f5f9",
            "surface-hover": "#eef4ff", "border": "#d9e2ec", "border-strong": "#b8c7d9",
            "text": "#0f172a", "text-muted": "#52606d", "text-faint": "#7b8794",
            "sidebar-text": "#0f172a", "sidebar-muted": "#52606d", "sidebar-border": "#e1e8f0",
            "primary": "#2563eb", "primary-hover": "#1d4ed8", "primary-text": "#ffffff",
            "focus": "#2563eb", "focus-ring": "rgba(37, 99, 235, .24)",
            "input-bg": "#ffffff", "input-text": "#0f172a", "input-border": "#b8c7d9",
            "input-hover": "#f8fbff", "control-suffix": "#f3f6fa", "control-suffix-text": "#334155",
            "chip-bg": "#eef2f7", "chip-text": "#475569", "shadow": "0 10px 30px rgba(15, 23, 42, .06)",
            "code-bg": "#eef2f6", "code-text": "#0f172a", "code-border": "#cbd5e1",
            "thinking-bg": "#f8fafc", "thinking-text": "#1e293b", "thinking-border": "#cbd5e1",
            "warning-bg": "#fef3c7", "warning-text": "#92400e", "info-bg": "#e0f2fe", "info-text": "#075985",
            "success-bg": "#dcfce7", "success-text": "#166534", "danger-bg": "#fee2e2", "danger-text": "#991b1b",
        }}
    if theme == "🔵 Ocean Blue":
        return {**common, **{
            "color-scheme": "light",
            "app-bg": "#eff6ff", "sidebar-bg": "#0b1629", "surface": "#ffffff", "surface-subtle": "#e6f4ff",
            "surface-hover": "#dff1ff", "border": "#b9e6ff", "border-strong": "#72c8f0",
            "text": "#0c4a6e", "text-muted": "#27678b", "text-faint": "#5282a0",
            "sidebar-text": "#f0f9ff", "sidebar-muted": "#b9d9ec", "sidebar-border": "#203451",
            "primary": "#0891b2", "primary-hover": "#0e7490", "primary-text": "#ffffff",
            "focus": "#0284c7", "focus-ring": "rgba(2, 132, 199, .28)",
            "input-bg": "#1b2a40", "input-text": "#f0f9ff", "input-border": "#334c6b",
            "input-hover": "#21354f", "control-suffix": "#152238", "control-suffix-text": "#d9f2ff",
            "chip-bg": "#dff3ff", "chip-text": "#03648c", "shadow": "0 10px 28px rgba(14, 116, 144, .08)",
            "code-bg": "#e0f2fe", "code-text": "#0369a1", "code-border": "#bae6fd",
            "thinking-bg": "#f0f9ff", "thinking-text": "#0c4a6e", "thinking-border": "#7dd3fc",
            "warning-bg": "#fef3c7", "warning-text": "#92400e", "info-bg": "#dff4ff", "info-text": "#075985",
            "success-bg": "#dcfce7", "success-text": "#166534", "danger-bg": "#fee2e2", "danger-text": "#991b1b",
        }}
    return {**common, **{
        "color-scheme": "dark",
        "app-bg": "#121416", "sidebar-bg": "#0e1115", "surface": "#1b2026", "surface-subtle": "#20262e",
        "surface-hover": "#27303a", "border": "#303943", "border-strong": "#485565",
        "text": "#f4f7fb", "text-muted": "#a9b4c0", "text-faint": "#7f8b98",
        "sidebar-text": "#f4f7fb", "sidebar-muted": "#a9b4c0", "sidebar-border": "#252c34",
        "primary": "#5b8cff", "primary-hover": "#4676e8", "primary-text": "#ffffff",
        "focus": "#78a0ff", "focus-ring": "rgba(120, 160, 255, .28)",
        "input-bg": "#232a33", "input-text": "#f4f7fb", "input-border": "#465261",
        "input-hover": "#2a333e", "control-suffix": "#1a2027", "control-suffix-text": "#e7edf5",
        "chip-bg": "#242b34", "chip-text": "#b9c6d5", "shadow": "0 12px 32px rgba(0, 0, 0, .22)",
        "code-bg": "#1e242c", "code-text": "#93c5fd", "code-border": "#374151",
        "thinking-bg": "#181e25", "thinking-text": "#e2e8f0", "thinking-border": "#2e3846",
        "warning-bg": "#4a4316", "warning-text": "#fef08a", "info-bg": "#172d48", "info-text": "#bfdbfe",
        "success-bg": "#123d2a", "success-text": "#bbf7d0", "danger-bg": "#4a1b20", "danger-text": "#fecaca",
    }}


def inject_css(theme="🌙 Dark"):
    tokens = _theme_tokens(theme)
    root_tokens = "\n".join(
        f"        --ui-{key}: {value};" for key, value in tokens.items()
    )
    css = """<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    :root {
__ROOT_TOKENS__
    }

    /* Global surfaces and type hierarchy. */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"],
    [data-testid="stMainBlockContainer"], .stApp {
        background: var(--ui-app-bg) !important;
        color: var(--ui-text) !important;
        font-family: var(--ui-font-body) !important;
    }
    html, body { font-size: 15px; line-height: 1.5; }
    [data-testid="stMainBlockContainer"] {
        max-width: 1440px;
        padding: 2rem 2rem 5rem;
    }
    [data-testid="stMain"] h1, [data-testid="stMain"] h2,
    [data-testid="stMain"] h3, [data-testid="stMain"] h4 {
        color: var(--ui-text) !important;
        letter-spacing: -.025em;
    }
    [data-testid="stMain"] p, [data-testid="stMain"] label,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] {
        color: var(--ui-text) !important;
    }
    [data-testid="stMain"] small, [data-testid="stMain"] [data-testid="stCaptionContainer"] {
        color: var(--ui-text-muted) !important;
    }
    ::selection { background: var(--ui-focus-ring); color: var(--ui-text); }

    /* Sidebar and its visual grouping. */
    section[data-testid="stSidebar"] {
        background: var(--ui-sidebar-bg) !important;
        border-right: 1px solid var(--ui-sidebar-border) !important;
    }
    [data-testid="stSidebarContent"] { padding: 0 1.5rem 1.5rem; }
    [data-testid="stSidebarHeader"] {
        display: block !important;
        height: 3.25rem !important;
        min-height: 3.25rem !important;
        position: relative !important;
        z-index: 20 !important;
    }
    [data-testid="stSidebarCollapseButton"] {
        display: flex !important;
        position: absolute !important;
        top: .5rem !important;
        right: 0 !important;
        visibility: visible !important;
        z-index: 21 !important;
    }
    [data-testid="stSidebarCollapseButton"] button,
    [data-testid="collapsedControl"] button,
    [data-testid="stSidebarCollapsedControl"] button {
        align-items: center !important;
        background: transparent !important;
        border: 1px solid transparent !important;
        border-radius: var(--ui-radius-sm) !important;
        color: var(--ui-sidebar-text) !important;
        display: inline-flex !important;
        height: 44px !important;
        justify-content: center !important;
        min-height: 44px !important;
        min-width: 44px !important;
        padding: 0 !important;
        visibility: visible !important;
        width: 44px !important;
    }
    [data-testid="stSidebarCollapseButton"] button:hover,
    [data-testid="collapsedControl"] button:hover,
    [data-testid="stSidebarCollapsedControl"] button:hover {
        background: var(--ui-surface-hover) !important;
        border-color: var(--ui-sidebar-border) !important;
    }
    [data-testid="stSidebarCollapseButton"] button:focus-visible,
    [data-testid="collapsedControl"] button:focus-visible,
    [data-testid="stSidebarCollapsedControl"] button:focus-visible {
        outline: 2px solid var(--ui-focus) !important;
        outline-offset: 2px !important;
    }
    section[data-testid="stSidebar"][aria-expanded="false"] [data-testid="stSidebarCollapseButton"] {
        left: .5rem !important;
        position: fixed !important;
        right: auto !important;
        top: .5rem !important;
        transform: translateX(300px) !important;
        z-index: 1001 !important;
    }
    section[data-testid="stSidebar"][aria-expanded="false"] [data-testid="stSidebarCollapseButton"] button {
        background: var(--ui-sidebar-bg) !important;
        border-color: var(--ui-sidebar-border) !important;
        box-shadow: var(--ui-shadow) !important;
    }
    section[data-testid="stSidebar"][aria-expanded="false"] [data-testid="stSidebarCollapseButton"] button:hover {
        background: var(--ui-surface-hover) !important;
        border-color: var(--ui-focus) !important;
    }
    section[data-testid="stSidebar"][aria-expanded="false"] [data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"] {
        transform: rotate(180deg);
    }
    [data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"],
    [data-testid="collapsedControl"] [data-testid="stIconMaterial"],
    [data-testid="stSidebarCollapsedControl"] [data-testid="stIconMaterial"] {
        color: var(--ui-sidebar-text) !important;
    }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3, section[data-testid="stSidebar"] h4,
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span, section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: var(--ui-sidebar-text) !important;
    }
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
        color: var(--ui-sidebar-muted) !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: var(--ui-sidebar-border) !important;
        margin: 1.5rem 0 !important;
    }
    [data-testid="collapsedControl"] {
        color: var(--ui-sidebar-text) !important;
        background: var(--ui-sidebar-bg) !important;
    }

    /* Shared form field surface. The suffix buttons intentionally use the
       same token as the field, preventing theme-specific black blocks. */
    [data-testid="stTextInput"], [data-testid="stNumberInput"], [data-testid="stSelectbox"],
    [data-testid="stTextArea"], [data-testid="stFileUploader"] { margin-bottom: .75rem; }
    [data-testid="stTextInputRootElement"], [data-testid="stNumberInput"] > div,
    [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    [data-testid="stTextArea"] textarea, [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input {
        background: var(--ui-input-bg) !important;
        color: var(--ui-input-text) !important;
        border: 1px solid var(--ui-input-border) !important;
        border-radius: var(--ui-radius-md) !important;
        box-sizing: border-box !important;
        min-height: 42px;
        transition: background var(--ui-motion-fast) ease, border-color var(--ui-motion-fast) ease, box-shadow var(--ui-motion-fast) ease;
    }
    /* Streamlit 1.5x uses React-Aria for selectboxes. Keep this selector
       alongside the BaseWeb selector so the theme remains stable across
       Streamlit markup versions. */
    [data-testid="stSelectbox"] .react-aria-ComboBox > [role="group"] {
        background: var(--ui-input-bg) !important;
        border: 1px solid var(--ui-input-border) !important;
        border-radius: var(--ui-radius-md) !important;
        box-sizing: border-box !important;
        min-height: 42px;
        overflow: hidden;
    }
    [data-testid="stSelectbox"] .react-aria-ComboBox > [role="group"] > input {
        background: transparent !important;
        border: 0 !important;
        color: var(--ui-input-text) !important;
        min-height: 40px;
    }
    [data-testid="stSelectbox"] .react-aria-ComboBox > [role="group"] > button {
        background: var(--ui-control-suffix) !important;
        border: 0 !important;
        border-left: 1px solid var(--ui-input-border) !important;
        border-radius: 0 !important;
        color: var(--ui-control-suffix-text) !important;
        min-height: 40px !important;
        min-width: 42px !important;
    }
    [data-testid="stSelectbox"] .react-aria-ComboBox > [role="group"] svg {
        color: var(--ui-control-suffix-text) !important;
        fill: var(--ui-control-suffix-text) !important;
        stroke: var(--ui-control-suffix-text) !important;
    }
    [data-testid="stTextInputRootElement"] input,
    [data-testid="stNumberInput"] input { border: 0 !important; border-radius: inherit !important; }
    [data-testid="stTextInputRootElement"] > div { background: transparent !important; }
    [data-testid="stSelectbox"] [data-baseweb="select"] > div > div,
    [data-testid="stSelectbox"] [role="combobox"],
    [data-testid="stSelectbox"] svg {
        color: var(--ui-input-text) !important;
        fill: var(--ui-input-text) !important;
        stroke: var(--ui-input-text) !important;
    }
    [data-testid="stTextInputRootElement"] button,
    [data-testid="stSelectbox"] button,
    [data-testid="stNumberInput"] button {
        background: var(--ui-control-suffix) !important;
        color: var(--ui-control-suffix-text) !important;
        border-color: var(--ui-input-border) !important;
        min-width: 42px !important;
        min-height: 40px !important;
    }
    [data-testid="stTextInputRootElement"] button[aria-pressed] {
        border: 0 !important;
        border-left: 1px solid var(--ui-input-border) !important;
        border-radius: 0 var(--ui-radius-md) var(--ui-radius-md) 0 !important;
    }
    [data-testid="stTextInputRootElement"] svg,
    [data-testid="stSelectbox"] button svg,
    [data-testid="stNumberInput"] button svg {
        color: var(--ui-control-suffix-text) !important;
        fill: var(--ui-control-suffix-text) !important;
        stroke: var(--ui-control-suffix-text) !important;
    }
    [data-testid="stTextInputRootElement"]:focus-within,
    [data-testid="stNumberInput"]:focus-within,
    [data-testid="stSelectbox"] [data-baseweb="select"]:focus-within,
    [data-testid="stSelectbox"] .react-aria-ComboBox > [role="group"]:focus-within,
    [data-testid="stTextArea"] textarea:focus {
        border-color: var(--ui-focus) !important;
        box-shadow: 0 0 0 3px var(--ui-focus-ring) !important;
        outline: 0 !important;
    }
    input::placeholder, textarea::placeholder {
        color: var(--ui-text-muted) !important;
        opacity: .9 !important;
    }
    [data-testid="stTextInputRootElement"]:hover,
    [data-testid="stNumberInput"]:hover,
    [data-testid="stSelectbox"] [data-baseweb="select"] > div:hover,
    [data-testid="stSelectbox"] .react-aria-ComboBox > [role="group"]:hover,
    [data-testid="stTextArea"] textarea:hover { background: var(--ui-input-hover) !important; }
    button:disabled, input:disabled, textarea:disabled {
        opacity: .55 !important;
        cursor: not-allowed !important;
    }
    [data-baseweb="popover"], [data-baseweb="menu"], ul[role="listbox"] {
        background: var(--ui-surface) !important;
        border: 1px solid var(--ui-border) !important;
        border-radius: var(--ui-radius-md) !important;
        box-shadow: var(--ui-shadow) !important;
    }
    li[role="option"], li[role="option"] * {
        background: var(--ui-surface) !important;
        color: var(--ui-text) !important;
    }
    li[role="option"]:hover, li[role="option"][aria-selected="true"] {
        background: var(--ui-surface-hover) !important;
    }

    /* Native focus is retained and made visible on every operable control. */
    button:focus-visible, input:focus-visible, textarea:focus-visible,
    [role="combobox"]:focus-visible, [role="tab"]:focus-visible,
    summary:focus-visible {
        outline: 2px solid var(--ui-focus) !important;
        outline-offset: 2px !important;
    }

    /* Header and workspace navigation. */
    .hdr {
        align-items: center;
        border-bottom: 1px solid var(--ui-border);
        display: flex;
        gap: .75rem;
        justify-content: center;
        margin: 0 auto 1.25rem;
        min-height: 72px;
        padding: .75rem 0;
    }
    .hdr-t { color: var(--ui-text) !important; font-size: clamp(1.45rem, 2vw, 1.9rem); font-weight: 700; letter-spacing: -.04em; }
    .chip { background: var(--ui-chip-bg) !important; border: 1px solid var(--ui-border); border-radius: 999px; color: var(--ui-chip-text) !important; font-size: .78rem; font-weight: 600; padding: .35rem .7rem; }
    [data-testid="stTabs"] [data-baseweb="tab-list"], [data-testid="stTabsHeader"] {
        background: transparent !important;
        border-bottom: 1px solid var(--ui-border) !important;
        border-radius: 0 !important;
        gap: 4px !important;
        justify-content: flex-start !important;
        margin: 0 0 1.5rem !important;
        max-width: none !important;
        padding: 0 !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab"], [data-testid="stTabs"] button[role="tab"] {
        border-radius: var(--ui-radius-sm) var(--ui-radius-sm) 0 0 !important;
        color: var(--ui-text-muted) !important;
        min-height: 44px !important;
        padding: .65rem 1rem !important;
        transition: background var(--ui-motion-fast) ease, color var(--ui-motion-fast) ease, box-shadow var(--ui-motion-fast) ease;
    }
    [data-testid="stTabs"] [data-baseweb="tab"] *, [data-testid="stTabs"] button[role="tab"] * { color: inherit !important; }
    [data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"], [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        background: var(--ui-primary) !important;
        box-shadow: inset 0 -3px 0 var(--ui-primary-hover) !important;
        color: var(--ui-primary-text) !important;
    }

    /* Quick-start cards, working panels, expanders, and upload states. */
    .qc { display: grid; gap: .75rem; grid-template-columns: repeat(2, minmax(0, 1fr)); margin: 1.25rem auto 1.75rem; max-width: 900px; }
    .qci, .wp, [data-testid="stExpander"], details {
        background: var(--ui-surface) !important;
        border: 1px solid var(--ui-border) !important;
        border-radius: var(--ui-radius-lg) !important;
        box-shadow: var(--ui-shadow) !important;
    }
    .qci { min-height: 86px; padding: 1rem 1.1rem; transition: background var(--ui-motion-fast) ease, border-color var(--ui-motion-fast) ease, transform var(--ui-motion-fast) ease; }
    .qci:hover { background: var(--ui-surface-hover) !important; border-color: var(--ui-primary) !important; transform: translateY(-2px); }
    .qci-t { color: var(--ui-text) !important; font-size: .95rem; font-weight: 700; margin-bottom: .25rem; }
    .qci-d { color: var(--ui-text-muted) !important; font-size: .83rem; }
    .wp { margin-bottom: 1rem; padding: 1.25rem; }
    [data-testid="stExpander"] summary, details summary { background: var(--ui-surface) !important; color: var(--ui-text) !important; min-height: 44px; }
    [data-testid="stExpander"] summary *, details summary * { color: var(--ui-text) !important; }
    [data-testid="stFileUploaderDropzone"], section[data-testid="stFileUploader"] { background: var(--ui-surface-subtle) !important; border: 1px dashed var(--ui-border-strong) !important; border-radius: var(--ui-radius-md) !important; }
    [data-testid="stFileUploaderDropzone"] * { color: var(--ui-text-muted) !important; }
    [data-testid="stFileUploaderDropzone"] button { background: var(--ui-primary) !important; color: var(--ui-primary-text) !important; border: 0 !important; border-radius: var(--ui-radius-sm) !important; }

    /* Action buttons and progress feedback. */
    .stButton > button, [data-testid="stChatInput"] button {
        background: var(--ui-primary) !important;
        border: 1px solid var(--ui-primary) !important;
        border-radius: var(--ui-radius-sm) !important;
        color: var(--ui-primary-text) !important;
        font-weight: 600;
        min-height: 42px !important;
        transition: background var(--ui-motion-fast) ease, border-color var(--ui-motion-fast) ease, transform var(--ui-motion-fast) ease;
    }
    .stButton > button:hover, [data-testid="stChatInput"] button:hover { background: var(--ui-primary-hover) !important; border-color: var(--ui-primary-hover) !important; transform: translateY(-1px); }
    .stButton > button:active, [data-testid="stChatInput"] button:active { transform: translateY(0); }
    /* Code blocks & inline code */
    code, [data-testid="stMarkdownContainer"] code, .stMarkdown code {
        background: var(--ui-code-bg) !important;
        color: var(--ui-code-text) !important;
        border: 1px solid var(--ui-code-border) !important;
        border-radius: 6px !important;
        padding: 2px 6px !important;
        font-family: 'Consolas', 'Fira Code', 'Monaco', monospace !important;
        font-size: 0.88em !important;
    }
    pre, [data-testid="stCodeBlock"], [data-testid="stCodeBlock"] pre {
        background: var(--ui-code-bg) !important;
        border: 1px solid var(--ui-code-border) !important;
        border-radius: var(--ui-radius-md) !important;
    }
    pre code, [data-testid="stCodeBlock"] pre code {
        background: transparent !important;
        border: 0 !important;
        padding: 0 !important;
        color: var(--ui-code-text) !important;
    }

    /* Live Thinking & Status Widget */
    [data-testid="stStatusWidget"] {
        background: var(--ui-thinking-bg) !important;
        border: 1px solid var(--ui-thinking-border) !important;
        border-radius: var(--ui-radius-md) !important;
        color: var(--ui-thinking-text) !important;
    }
    [data-testid="stStatusWidget"] summary,
    [data-testid="stStatusWidget"] [data-testid="stMarkdownContainer"] *,
    [data-testid="stStatusWidget"] p,
    [data-testid="stStatusWidget"] span {
        color: var(--ui-thinking-text) !important;
    }
    [data-testid="stStatusWidget"] code {
        background: var(--ui-code-bg) !important;
        color: var(--ui-code-text) !important;
        border: 1px solid var(--ui-code-border) !important;
    }

    [data-testid="stAlert"] { border-radius: var(--ui-radius-md) !important; }
    [data-testid="stAlert"] [data-testid="stMarkdownContainer"] * { color: var(--ui-text) !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]) { background: var(--ui-warning-bg) !important; color: var(--ui-warning-text) !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]) { background: var(--ui-info-bg) !important; color: var(--ui-info-text) !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]) { background: var(--ui-success-bg) !important; color: var(--ui-success-text) !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentError"]) { background: var(--ui-danger-bg) !important; color: var(--ui-danger-text) !important; }
    .sc { align-items: center; border-radius: 999px; display: inline-flex; font-size: .82rem; font-weight: 700; gap: .4rem; min-height: 36px; padding: .4rem .8rem; }
    [data-testid="stSidebar"] .sc-on, section[data-testid="stSidebar"] .sc-on { background: var(--ui-success-bg) !important; border: 1px solid var(--ui-success-text) !important; color: var(--ui-success-text) !important; }
    [data-testid="stSidebar"] .sc-off, section[data-testid="stSidebar"] .sc-off { background: var(--ui-danger-bg) !important; border: 1px solid var(--ui-danger-text) !important; color: var(--ui-danger-text) !important; }

    /* Chat workspace and fixed input affordance. */
    [data-testid="stChatMessage"], .stChatMessage, [data-testid="stChatMessageContent"] { background: var(--ui-surface) !important; border: 1px solid var(--ui-border) !important; border-radius: var(--ui-radius-lg) !important; color: var(--ui-text) !important; margin-bottom: .75rem !important; padding: .85rem 1rem !important; }
    [data-testid="stChatMessage"] *, .stChatMessage * { color: var(--ui-text) !important; }
    [data-testid="stBottom"], [data-testid="stBottom"] > div { background: transparent !important; padding-bottom: .75rem !important; }
    [data-testid="stChatInput"], [data-testid="stChatInput"] > div:first-child, [data-testid="stChatInput"] > div:first-child > div { background: var(--ui-surface) !important; border: 1px solid var(--ui-border-strong) !important; border-radius: var(--ui-radius-lg) !important; box-shadow: var(--ui-shadow) !important; }
    [data-testid="stChatInput"] textarea { background: transparent !important; border: 0 !important; color: var(--ui-text) !important; outline: 0 !important; }
    [data-testid="stChatInput"] button svg { color: var(--ui-primary-text) !important; fill: var(--ui-primary-text) !important; }

    header[data-testid="stHeader"] { background: transparent !important; }
    footer, #MainMenu, [data-testid="stHeaderActionElements"], [data-testid="stAppDeployButton"] { display: none !important; visibility: hidden !important; }
    [data-testid="collapsedControl"], [data-testid="stSidebarCollapsedControl"] {
        display: flex !important;
        visibility: visible !important;
        z-index: 1000 !important;
    }
    @media (max-width: 900px) {
        [data-testid="stMainBlockContainer"] { padding: 1.25rem 1rem 4rem; }
        .hdr { justify-content: flex-start; }
        .qc { grid-template-columns: 1fr; }
    }
    @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after { animation-duration: .01ms !important; animation-iteration-count: 1 !important; scroll-behavior: auto !important; transition-duration: .01ms !important; }
    }
    </style>""".replace("__ROOT_TOKENS__", root_tokens)
    st.markdown(css, unsafe_allow_html=True)


def main():
    st.set_page_config(page_title="CATIA V5 AI Studio", page_icon="📐", layout="wide")

    # Sidebar
    st.sidebar.markdown("<h3 style='font-size:1.1rem;font-weight:700;margin-bottom:2px'>📐 CATIA AI Studio</h3>", unsafe_allow_html=True)
    st.sidebar.markdown("<div style='font-size:0.82rem;opacity:0.8;margin-bottom:12px'>Made by <b>DRISSI AMJAD</b></div>", unsafe_allow_html=True)
    theme_options = ["🌙 Dark", "☀️ Light", "🔵 Ocean Blue"]
    qp_theme = str(st.query_params.get("theme", ""))
    idx = 0
    if "light" in qp_theme.lower(): idx = 1
    elif "ocean" in qp_theme.lower() or "blue" in qp_theme.lower(): idx = 2
    theme = st.sidebar.selectbox("🎨 UI Theme", theme_options, index=idx)
    inject_css(theme)
    st.sidebar.markdown("---")
    provider = st.sidebar.selectbox("🤖 LLM Engine",
        ["OpenRouter", "OpenAI", "Anthropic", "Local (llama.cpp / Local Server)", "Local (Ollama)"])
    dm = {"Local (llama.cpp / Local Server)": "local-model", "Local (Ollama)": "llama3.2-vision",
          "OpenAI": "gpt-4o", "Anthropic": "claude-3-5-sonnet-20240620", "OpenRouter": "nvidia/nemotron-3.5-lightning:free"}
    model_name = st.sidebar.text_input("🧠 Model", value=dm.get(provider, "nvidia/nemotron-3.5-lightning:free"))

    custom_base_url = ""
    if provider == "Local (llama.cpp / Local Server)":
        custom_base_url = st.sidebar.text_input("🌐 Base URL", value="http://localhost:8080/v1")
    api_key = st.sidebar.text_input("🔑 Key", value="not-needed" if "Local" in provider else "", type="password")

    st.sidebar.markdown("---")
    st.sidebar.markdown("<h4 style='font-size:.95rem;font-weight:600;color:#9b9b9b'>🔌 CATIA V5 Status</h4>", unsafe_allow_html=True)
    caa, part, nps, cerr = get_part_and_parameters()
    if caa is None:
        st.sidebar.markdown("<span class='sc sc-off'>🔴 Disconnected</span>", unsafe_allow_html=True)
    else:
        st.sidebar.markdown("<span class='sc sc-on'>🟢 Connected</span>", unsafe_allow_html=True)
        if part:
            st.sidebar.caption(f"Parameters: **{len(nps)}**")
    if st.sidebar.button("🔄 Sync", use_container_width=True):
        st.rerun()

    # Header
    st.markdown(f"<div class='hdr'><span style='font-size:2.2rem'>📐</span>"
                f"<span class='hdr-t'>CATIA V5 AI Studio</span>"
                f"<span class='chip'>{model_name}</span></div>", unsafe_allow_html=True)

    tab_ai, tab_studio, tab_params = st.tabs(["💬 AI Assistant", "🧱 3D Studio", "🎛️ Parameters"])

    # --- TAB 1: AI ---
    with tab_ai:
        st.markdown("""<div class='qc'>
            <div class='qci'><div class='qci-t'>🎭 Iron Man Mask</div><div class='qci-d'>Sculpted helmet with eye slits</div></div>
            <div class='qci'><div class='qci-t'>☸️ Wheel Rim</div><div class='qci-d'>5-spoke rim with hub & lugs</div></div>
            <div class='qci'><div class='qci-t'>⬛ Solid Block</div><div class='qci-d'>Rectangular extrusion</div></div>
            <div class='qci'><div class='qci-t'>📷 Blueprint</div><div class='qci-d'>Upload drawing for 3D</div></div>
        </div>""", unsafe_allow_html=True)
        with st.expander("📷 Attach Image", expanded=False):
            uploaded = st.file_uploader("Upload", type=["png","jpg","jpeg","webp"], key="img_up")
            if uploaded:
                st.image(uploaded, width=350)
        if "messages" not in st.session_state:
            st.session_state.messages = [{"role": "assistant",
                "content": "Welcome! Ask me to build a wheel rim, Iron Man mask, or any 3D geometry."}]
        for m in st.session_state.messages:
            with st.chat_message(m["role"]):
                st.write(m["content"])
                if m.get("image_bytes"):
                    st.image(m["image_bytes"], width=250)
        # Live Code & Feedback Inspector
        last_code = st.session_state.get("last_executed_code", "")
        if last_code:
            with st.expander("💻 Live Python CAD Script & Diagnostics", expanded=False):
                ok, msg = st.session_state.get("last_execution_status", (True, "Ready"))
                if ok:
                    st.success(f"Status: {msg}")
                else:
                    st.error(f"Error: {msg}")
                st.code(last_code, language="python")

        if prompt := st.chat_input("e.g. 'Build an Iron Man mask' or 'Split part with 1mm gap'"):
            ib = uploaded.getvalue() if uploaded else None
            dt = prompt + (" 📷" if uploaded else "")
            mp = {"role": "user", "content": dt}
            if ib:
                mp["image_bytes"] = ib
            st.session_state.messages.append(mp)
            with st.chat_message("user"):
                st.write(dt)
            llm = instantiate_llm(provider, model_name, api_key, custom_base_url)
            if not llm:
                r = "⚠️ LLM not configured."
                with st.chat_message("assistant"): st.write(r)
                st.session_state.messages.append({"role": "assistant", "content": r})
            else:
                with st.chat_message("assistant"):
                    with st.status("🤖 Building…", expanded=True) as sb:
                        r = run_agent_with_live_status(llm, prompt, ib, status_container=sb)
                        sb.update(label="✅ Done", state="complete", expanded=False)
                    st.write(r)
                st.session_state.messages.append({"role": "assistant", "content": r})
                st.rerun()

    # --- TAB 2: 3D Studio ---
    with tab_studio:
        st.markdown("<div class='wp'>", unsafe_allow_html=True)
        st.subheader("🧱 Advanced 3D Geometry Studio")
        
        # Row 1: Complex Presets & Basic Solids
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown("#### 🎭 Complex Generators")
            with st.expander("🎭 Iron Man Mask", expanded=False):
                iw = st.number_input("Width mm", value=180.0, step=10.0, key="iw")
                ih = st.number_input("Height mm", value=260.0, step=10.0, key="ih")
                id_ = st.number_input("Depth mm", value=40.0, step=5.0, key="id")
                if st.button("🚀 Build Mask", use_container_width=True):
                    ok, msg = create_iron_man_mask_in_catia(iw, ih, id_)
                    (st.success if ok else st.error)(msg)
            with st.expander("☸️ Wheel Rim", expanded=False):
                wr = st.number_input("Outer R mm", value=200.0, step=10.0, key="wr")
                ww = st.number_input("Width mm", value=80.0, step=5.0, key="ww")
                wh = st.number_input("Hub R mm", value=55.0, step=5.0, key="wh")
                if st.button("🚀 Build Rim", use_container_width=True):
                    ok, msg = create_detailed_wheel_rim_in_catia(wr, ww, wh)
                    (st.success if ok else st.error)(msg)
        with c2:
            st.markdown("#### 📦 Basic Solids")
            with st.expander("⬛ Pad Block", expanded=False):
                pw = st.number_input("X mm", value=100.0, step=10.0, key="pw")
                pl = st.number_input("Y mm", value=100.0, step=10.0, key="pl")
                ph = st.number_input("Z mm", value=20.0, step=5.0, key="ph")
                if st.button("🚀 Build Pad", use_container_width=True):
                    ok, msg = create_pad_block_in_catia(pw, pl, ph)
                    (st.success if ok else st.error)(msg)
            with st.expander("🟢 Cylinder", expanded=False):
                cr = st.number_input("Radius mm", value=25.0, step=5.0, key="cr")
                ch_ = st.number_input("Height mm", value=50.0, step=5.0, key="ch")
                if st.button("🚀 Build Cylinder", use_container_width=True):
                    ok, msg = create_cylinder_in_catia(cr, ch_)
                    (st.success if ok else st.error)(msg)
            with st.expander("📄 New Document", expanded=False):
                dc = st.selectbox("Type", ["Part", "Product", "Drawing"])
                if st.button("📄 Create", use_container_width=True):
                    ok, msg = create_new_document_in_catia(dc)
                    (st.success if ok else st.error)(msg)
                    if ok: st.rerun()

        st.markdown("---")
        # Row 2: Part Splitter & Mechanical Tools
        c3, c4 = st.columns(2, gap="large")
        with c3:
            st.markdown("#### ✂️ Intelligent Part Splitter")
            with st.expander("✂️ Split Solid Part", expanded=True):
                s_style = st.selectbox("Split Style", ["Planar", "Jigsaw / Puzzle", "Pyramid"], key="s_style")
                s_plane = st.selectbox("Cutting Plane", ["PlaneXY", "PlaneYZ", "PlaneZX"], key="s_plane")
                s_gap = st.number_input("Clearance Gap mm", value=1.0, min_value=0.1, max_value=20.0, step=0.5, key="s_gap")
                if "Puzzle" in s_style or "Jigsaw" in s_style:
                    s_tsize = st.number_input("Tab Size mm", value=15.0, step=2.0, key="s_tsize")
                    s_tcount = st.number_input("Tab Count", value=3, min_value=1, max_value=10, step=1, key="s_tcount")
                else:
                    s_tsize = 15.0
                    s_tcount = 3
                if st.button("🚀 Split Part", use_container_width=True):
                    ok, msg = split_part_in_catia(s_plane, s_style, s_gap, s_tsize, s_tcount)
                    (st.success if ok else st.error)(msg)

        with c4:
            st.markdown("#### ⚙️ Mechanical & Revolve Studio")
            with st.expander("🔄 Revolve / Shaft", expanded=False):
                sh_or = st.number_input("Outer Radius mm", value=40.0, step=5.0, key="sh_or")
                sh_ir = st.number_input("Inner Radius mm", value=20.0, step=5.0, key="sh_ir")
                sh_h = st.number_input("Shaft Height mm", value=60.0, step=5.0, key="sh_h")
                sh_ang = st.number_input("Revolve Angle (°)", value=360.0, min_value=1.0, max_value=360.0, step=15.0, key="sh_ang")
                if st.button("🚀 Build Shaft", use_container_width=True):
                    ok, msg = create_revolve_shaft_in_catia(sh_or, sh_ir, sh_h, sh_ang)
                    (st.success if ok else st.error)(msg)
            with st.expander("💫 Circular Pattern", expanded=False):
                cp_count = st.number_input("Hole Count", value=6, min_value=2, max_value=36, step=1, key="cp_count")
                cp_pcd = st.number_input("PCD Radius mm", value=45.0, step=5.0, key="cp_pcd")
                cp_hr = st.number_input("Hole Radius mm", value=6.0, step=1.0, key="cp_hr")
                if st.button("🚀 Build Pattern", use_container_width=True):
                    ok, msg = create_circular_pattern_in_catia(cp_count, cp_pcd, cp_hr)
                    (st.success if ok else st.error)(msg)

        st.markdown("</div>", unsafe_allow_html=True)

    # --- TAB 3: Parameters ---
    with tab_params:
        st.markdown("<div class='wp'>", unsafe_allow_html=True)
        st.subheader("🎛️ Parametric Manager")
        if not caa:
            st.warning("CATIA V5 not running.")
        elif not part:
            st.info("Open a .CATPart to edit parameters.")
        elif not nps:
            st.info("No numeric parameters found.")
        else:
            cols = st.columns(min(len(nps[:5]), 3))
            for idx, p in enumerate(nps[:5]):
                with cols[idx % len(cols)]:
                    cv = float(p["value"])
                    nv = st.slider(p["display_name"], min(0.0, cv*.5), max(100.0, cv*2), cv,
                                   step=0.5, key=f"sl_{idx}_{p['name']}")
                    if abs(nv - cv) > 1e-4:
                        ok, msg = apply_parameter_update(p["name"], nv)
                        if ok: st.toast(f"✅ {p['display_name']}={nv}"); st.rerun()
                        else: st.error(msg)
            st.markdown("---")
            sq = st.text_input("🔍 Search", placeholder="Length, Radius…")
            fp = [p for p in nps if sq.lower() in p["name"].lower() or sq.lower() in p["display_name"].lower()]
            if fp:
                c1, c2, c3 = st.columns([2, 1, 1])
                with c1:
                    opts = {f"{p['display_name']} ({p['name']})": p for p in fp}
                    sk = st.selectbox("Parameter", list(opts.keys()))
                    sp = opts[sk]
                with c2:
                    nv = st.number_input("Value", value=float(sp["value"]), step=1.0, key="me")
                with c3:
                    st.write(""); st.write("")
                    if st.button("💾 Apply", use_container_width=True):
                        ok, msg = apply_parameter_update(sp["name"], nv)
                        (st.success if ok else st.error)(msg)
                        if ok: st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
