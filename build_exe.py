import os
import sys
import subprocess
import streamlit

# Locate Streamlit directory
streamlit_dir = os.path.dirname(streamlit.__file__)
static_dir = os.path.join(streamlit_dir, "static")

print(f"[*] Streamlit module path: {streamlit_dir}")
print(f"[*] Streamlit static assets: {static_dir}")

# Build PyInstaller command with exclusions for lightweight distribution
cmd = [
    "pyinstaller",
    "--noconfirm",
    "--onedir",
    "--name=CATIA_AI_Bridge",
    f"--add-data={static_dir};streamlit/static",
    "--add-data=app.py;.",
    # Hidden imports
    "--hidden-import=streamlit",
    "--hidden-import=pycatia",
    "--hidden-import=win32com",
    "--hidden-import=pythoncom",
    "--hidden-import=webview",
    "--hidden-import=pythonnet",
    "--hidden-import=bottle",
    "--hidden-import=langchain",
    "--hidden-import=langchain_classic",
    "--hidden-import=langchain_core",
    "--hidden-import=langchain_community",
    "--hidden-import=langchain_openai",
    "--hidden-import=langchain_anthropic",
    "--hidden-import=langchain_ollama",
    "--hidden-import=PIL",
    # Exclude unused heavy packages to make build lightweight and fast
    "--exclude-module=matplotlib",
    "--exclude-module=scipy",
    "--exclude-module=pyarrow",
    "--exclude-module=numba",
    "--exclude-module=torch",
    "--exclude-module=tkinter",
    "--collect-all=streamlit",
    "--collect-all=pycatia",
    "--collect-all=webview",
    "--collect-all=langchain",
    "--collect-all=langchain_classic",
    "--collect-all=langchain_core",
    "--collect-all=langchain_community",
    "--collect-all=langchain_openai",
    "--collect-all=langchain_anthropic",
    "--collect-all=langchain_ollama",
    "run_app.py"
]

print("[*] Starting PyInstaller lightweight build...")
result = subprocess.run(cmd)

if result.returncode == 0:
    print("\n========================================================")
    print("[SUCCESS] LIGHTWEIGHT STANDALONE BUILD COMPLETE!")
    print("Executable directory: dist/CATIA_AI_Bridge/")
    print("Main Executable: dist/CATIA_AI_Bridge/CATIA_AI_Bridge.exe")
    print("========================================================\n")
else:
    print(f"\n[ERROR] Build failed with return code {result.returncode}")
