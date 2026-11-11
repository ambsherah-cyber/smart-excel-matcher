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
    # Disable development mode before Streamlit loads its configuration.
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENTMODE"] = "false"

    import streamlit.config as stconfig
    stconfig.set_option("global.developmentMode", False)

    import streamlit.web.cli as stcli

    app_file = str(resource_path("app.py"))
    # Use Streamlit's default port. Do not pass --server.port:
    # Streamlit rejects an explicitly configured port in development mode.
    url = "http://127.0.0.1:8501"

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
        "--server.headless=true",
        "--browser.gatherUsageStats=false"
    ]

    stcli.main()

if __name__ == "__main__":
    main()
