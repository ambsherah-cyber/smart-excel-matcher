
import os
import sys
import time
import socket
import threading
import webbrowser
from pathlib import Path

def resource_path(filename):
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / filename
    return Path(__file__).parent / filename

def find_port():
    sock = socket.socket()
    sock.bind(("127.0.0.1", 0))
    port = sock.getsockname()[1]
    sock.close()
    return port

def open_browser(url):
    time.sleep(3)
    webbrowser.open(url)

def main():
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENTMODE"] = "false"

    import streamlit.web.cli as stcli

    app_file = str(resource_path("app.py"))
    port = find_port()
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
        "--browser.gatherUsageStats=false"
    ]

    stcli.main()

if __name__ == "__main__":
    main()
