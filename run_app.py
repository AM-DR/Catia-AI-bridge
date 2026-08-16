import sys
import os
import time
import socket
import subprocess
import webbrowser

# ==============================================================================
# CRITICAL: 100% BLOCK PYTHON & STREAMLIT FROM EVER OPENING CHROME / BROWSER
# ==============================================================================
webbrowser.open = lambda *args, **kwargs: True
webbrowser.open_new = lambda *args, **kwargs: True
webbrowser.open_new_tab = lambda *args, **kwargs: True


def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def run_streamlit_server_main_thread():
    """ 
    Runs Streamlit on the MAIN thread of a sub-process so Python signal handlers 
    (signal.SIGTERM) succeed without raising ValueError!
    """
    from streamlit.web.bootstrap import run
    from streamlit import config

    script_path = get_resource_path("app.py")

    config.set_option("server.headless", True)
    config.set_option("global.developmentMode", False)
    config.set_option("browser.gatherUsageStats", False)
    config.set_option("client.toolbarMode", "off")
    config.set_option("server.port", 8501)
    config.set_option("server.address", "127.0.0.1")

    try:
        run(
            main_script_path=script_path,
            is_hello=False,
            args=[],
            flag_options={
                "global.developmentMode": False,
                "server.headless": True,
                "server.port": 8501,
                "server.address": "127.0.0.1",
                "browser.gatherUsageStats": False,
                "client.toolbarMode": "off"
            }
        )
    except Exception as e:
        print(f"Streamlit server exit note: {e}")


def wait_for_port(port=8501, host="127.0.0.1", max_wait_sec=20):
    """ Polling TCP socket until port 8501 is live """
    start_time = time.time()
    while time.time() - start_time < max_wait_sec:
        try:
            with socket.create_connection((host, port), timeout=0.5):
                return True
        except (OSError, ConnectionRefusedError):
            time.sleep(0.2)
    return False


if __name__ == "__main__":
    # 1. SERVER SUBPROCESS WORKER MODE
    if "--server" in sys.argv:
        run_streamlit_server_main_thread()

    # 2. MAIN NATIVE DESKTOP GUI LAUNCHER MODE
    else:
        # Spawn hidden server worker process on its own main thread
        server_process = subprocess.Popen(
            [sys.executable, "--server"],
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        )

        # Wait until local TCP port 8501 is ready
        server_ready = wait_for_port(8501, host="127.0.0.1", max_wait_sec=20)

        if server_ready:
            # Import pywebview and launch native desktop window exclusively
            import webview
            webview.create_window(
                title="CATIA V5 R21 AI Studio",
                url="http://127.0.0.1:8501",
                width=1450,
                height=920,
                resizable=True,
                min_size=(1024, 700)
            )
            webview.start()

        # Cleanly terminate background server process when window closes
        try:
            server_process.terminate()
        except Exception:
            pass
