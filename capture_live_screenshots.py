import subprocess, time, sys, os
from playwright.sync_api import sync_playwright

subprocess.run(["taskkill", "/F", "/IM", "streamlit.exe"], capture_output=True)
proc = subprocess.Popen(
    [sys.executable, "-m", "streamlit", "run", "app.py", "--server.port", "8501", "--server.headless", "true"],
    cwd=r"C:\Users\AMJAD\Desktop\New folder (2)\New folder",
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE
)
time.sleep(4)

out_dir = r"C:\Users\AMJAD\Desktop\New folder (2)\New folder\docs\images"
os.makedirs(out_dir, exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={"width": 1366, "height": 850}, device_scale_factor=1.5)
    page = context.new_page()

    # 1. Dark Assistant
    print("[*] 1. Dark Assistant...")
    page.goto("http://localhost:8501/?theme=dark", wait_until="networkidle")
    page.wait_for_selector("[data-testid='stSidebar']", timeout=30000)
    time.sleep(2)
    page.locator("text=AI Assistant").first.click()
    time.sleep(1.5)
    page.screenshot(path=os.path.join(out_dir, "dark_assistant.png"))
    print("[+] Saved dark_assistant.png")

    # 2. Ocean Blue 3D Studio
    print("[*] 2. Ocean Blue 3D Studio...")
    page.goto("http://localhost:8501/?theme=ocean_blue", wait_until="networkidle")
    page.wait_for_selector("[data-testid='stSidebar']", timeout=30000)
    time.sleep(2)
    page.locator("text=3D Studio").first.click()
    time.sleep(2)
    page.screenshot(path=os.path.join(out_dir, "ocean_blue_studio.png"))
    print("[+] Saved ocean_blue_studio.png")

    # 3. Light Assistant
    print("[*] 3. Light Assistant...")
    page.goto("http://localhost:8501/?theme=light", wait_until="networkidle")
    page.wait_for_selector("[data-testid='stSidebar']", timeout=30000)
    time.sleep(2)
    page.locator("text=AI Assistant").first.click()
    time.sleep(1.5)
    page.screenshot(path=os.path.join(out_dir, "light_assistant.png"))
    print("[+] Saved light_assistant.png")

    # 4. Parameters Manager
    print("[*] 4. Parameters...")
    page.goto("http://localhost:8501/?theme=dark", wait_until="networkidle")
    page.wait_for_selector("[data-testid='stSidebar']", timeout=30000)
    time.sleep(2)
    page.locator("text=Parameters").first.click()
    time.sleep(2)
    page.screenshot(path=os.path.join(out_dir, "parameters.png"))
    print("[+] Saved parameters.png")

    # 5. Desktop Window (3D Studio in Dark Theme)
    print("[*] 5. Desktop Window (3D Studio)...")
    page.goto("http://localhost:8501/?theme=dark", wait_until="networkidle")
    page.wait_for_selector("[data-testid='stSidebar']", timeout=30000)
    time.sleep(2)
    page.locator("text=3D Studio").first.click()
    time.sleep(2)
    page.screenshot(path=os.path.join(out_dir, "desktop_window.png"))
    print("[+] Saved desktop_window.png")

    browser.close()

proc.terminate()
print("[ALL DONE] Perfectly captured all 5 distinct images!")
