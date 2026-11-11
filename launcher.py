import os
import sys
import time
import threading
import webbrowser
from pathlib import Path

def resource_path(filename):
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / filename
    return Path(__file__).parent / filename

def open_browser(url):
    time.sleep(3)
    webbrowser.open(url)

def main():
    # Force Streamlit out of development mode before importing its CLI.
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENTMODE"] = "false"

    import streamlit.web.cli as stcli

    app_file = str(resource_path("app.py"))
    port = 8501
    url = f"http://127.0.0.1:{port}"

    threading.Thread(
        target=open_browser,
        args=(url,),
        daemon=True
    ).start()

    sys.argv = [
        "streamlit",
        "run",
        app_file,
        "--server.address=127.0.0.1",
        f"--server.port={port}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false"
    ]

    stcli.main()

if __name__ == "__main__":
    main()
