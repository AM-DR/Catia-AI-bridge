# 📐 CATIA V5 R21 AI Studio

**Created by DRISSI AMJAD**

An AI-driven CAD automation and geometry generation bridge connecting Local/Cloud LLMs (llama.cpp, Ollama, OpenAI, Anthropic, OpenRouter) to **CATIA V5 R21** via Windows COM automation (`pycatia` / `win32com`).

---

## 🌟 Key Features

1. **Native Desktop Window (`pywebview`)**:
   - Zero browser popups or tabs; runs as a dedicated standalone Windows GUI window.
2. **Boolean Remove CAD Architecture (`AddNewRemove`)**:
   - Direction-independent through-cuts that bypass CATIA's directional Pocket limitations.
3. **Intelligent Part Splitter Engine**:
   - **Planar Split**: Precise cuts along XY, YZ, or ZX planes with configurable clearance gaps (e.g., `1.0mm`).
   - **Jigsaw / Puzzle Interlocking Split**: Slices parts with dovetail interlocking puzzle teeth for multi-piece assembly and 3D printing.
   - **4-Quadrant Pyramid / Diagonal X-Split**: Dual diagonal cuts separating blocks into 4 interlocking pyramid wedges.
4. **Mechanical & Revolve Studio**:
   - **Revolve / Shaft**: Turned cylinders, bushings, and shafts with explicit 2D CenterLine axis control.
   - **Circular Pattern**: Radial circular hole arrays generated directly in 2D sketches.
5. **AI Assistant & Live Code Inspector**:
   - Natural language CAD coding agent with full `pycatia` cheat sheet and live syntax-highlighted execution inspector.
6. **Multi-Theme Engine**:
   - Customizable UI themes: 🌙 Dark Mode, ☀️ Light Mode, and 🔵 Ocean Blue.

---

## 🚀 Quick Start

- **Launch Executable**: Run `launch_app.bat` or open `dist\CATIA_AI_Bridge\CATIA_AI_Bridge.exe`.
- **Rebuild Executable**: Run `python build_exe.py`.
- **Run in Development**: Run `python run_app.py`.

---

## 📦 Requirements

- Python 3.10+
- CATIA V5 (R20, R21, V5-6R202X)
- Dependencies in `requirements.txt`
