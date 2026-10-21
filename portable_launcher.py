"""Run the attendance app from a single-file Windows bundle."""

import os
import socket
import sys
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path


def open_browser_when_ready(url):
    health_url = f"{url}/_stcore/health"
    for _ in range(120):
        try:
            with urllib.request.urlopen(health_url, timeout=1) as response:
                if response.status == 200:
                    webbrowser.open(url)
                    return
        except (OSError, urllib.error.URLError):
            time.sleep(0.5)
    print(f"Could not open the app automatically. Open {url} in your browser.")


def main():
    resource_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    app_path = resource_dir / "app.py"
    if not app_path.is_file():
        raise FileNotFoundError(f"Bundled app entry point not found: {app_path}")

    sys.path.insert(0, str(resource_dir))
    with socket.socket() as server_socket:
        server_socket.bind(("127.0.0.1", 0))
        port = server_socket.getsockname()[1]

    app_url = f"http://localhost:{port}"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ENABLE_STATIC_SERVING"] = "true"

    threading.Thread(
        target=open_browser_when_ready,
        args=(app_url,),
        daemon=True,
    ).start()

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.headless=true",
        "--server.enableStaticServing=true",
        f"--server.port={port}",
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
    ]

    from streamlit.web.cli import main as streamlit_main

    streamlit_main()


if __name__ == "__main__":
    main()
