import subprocess
import time
import os
import sys
import socket
from playwright.sync_api import sync_playwright

def is_port_open(port=8501):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def capture():
    # 1. Start Streamlit if not already running
    proc = None
    if not is_port_open(8501):
        print("[*] Starting Streamlit server...")
        proc = subprocess.Popen(
            [sys.executable, "-m", "streamlit", "run", "app.py", "--server.port", "8501", "--server.headless", "true"],
            cwd=r"C:\Users\AMJAD\Desktop\New folder (2)\New folder",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        for _ in range(30):
            if is_port_open(8501):
                print("[+] Streamlit is up and running!")
                break
            time.sleep(1)

    out_dir = r"C:\Users\AMJAD\Desktop\New folder (2)\New folder\docs\images"
    os.makedirs(out_dir, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1400, "height": 900})
        page = context.new_page()

        print("[*] Navigating to http://localhost:8501...")
        page.goto("http://localhost:8501", wait_until="networkidle")
        page.wait_for_selector("[data-testid='stSidebar']", timeout=30000)
        page.wait_for_selector(".block-container", timeout=30000)
        time.sleep(4)

        # Helper to click a tab by text
        def click_tab(tab_name):
            tabs = page.locator("button[data-testid='stTab']")
            count = tabs.count()
            for i in range(count):
                t = tabs.nth(i)
                if tab_name.lower() in (t.inner_text() or "").lower():
                    t.click()
                    time.sleep(2)
                    return True
            return False

        # Helper to set Theme via Selectbox
        def set_theme(theme_name):
            try:
                # Find the theme selectbox container
                sb = page.locator("[data-testid='stSelectbox']").first
                sb.click()
                time.sleep(1)
                # Click the option from dropdown
                opt = page.locator(f"li:has-text('{theme_name}')").first
                opt.click()
                time.sleep(2.5)
                return True
            except Exception as e:
                print(f"Error setting theme {theme_name}: {e}")
                return False

        # 1. Dark Assistant Screenshot
        print("[*] Capturing Dark Assistant...")
        set_theme("Dark")
        click_tab("AI Assistant")
        page.wait_for_timeout(2000)
        page.screenshot(path=os.path.join(out_dir, "dark_assistant.png"))
        print("[+] Saved dark_assistant.png")

        # 2. Ocean Blue 3D Studio Screenshot
        print("[*] Capturing Ocean Blue 3D Studio...")
        set_theme("Ocean Blue")
        click_tab("3D Geometry Studio")
        page.wait_for_timeout(2000)
        page.screenshot(path=os.path.join(out_dir, "ocean_blue_studio.png"))
        print("[+] Saved ocean_blue_studio.png")

        # 3. Light Assistant Screenshot
        print("[*] Capturing Light Assistant...")
        set_theme("Light")
        click_tab("AI Assistant")
        page.wait_for_timeout(2000)
        page.screenshot(path=os.path.join(out_dir, "light_assistant.png"))
        print("[+] Saved light_assistant.png")

        # 4. Parameters Screenshot
        print("[*] Capturing Parameters Manager...")
        set_theme("Dark")
        click_tab("Live Parameter Manager")
        page.wait_for_timeout(2000)
        page.screenshot(path=os.path.join(out_dir, "parameters.png"))
        print("[+] Saved parameters.png")

        # 5. Desktop Window Overview Screenshot
        print("[*] Capturing Desktop Window Overview...")
        set_theme("Ocean Blue")
        click_tab("3D Geometry Studio")
        page.wait_for_timeout(2000)
        page.screenshot(path=os.path.join(out_dir, "desktop_window.png"))
        print("[+] Saved desktop_window.png")

        browser.close()

    if proc:
        print("[*] Stopping background Streamlit worker...")
        proc.terminate()

    print("[SUCCESS] All 5 live screenshots captured and saved to docs/images/!")

if __name__ == "__main__":
    capture()
