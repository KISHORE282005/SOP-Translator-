"""
Single command launcher for the PO Translation System.

Starts the FastAPI backend and the Streamlit frontend together:

    python run.py

Press Ctrl+C to stop both.
"""

import os
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

API_PORT = os.getenv("API_PORT", "8000")
UI_PORT = os.getenv("UI_PORT", "8501")
os.environ.setdefault("API_BASE_URL", f"http://localhost:{API_PORT}")

HEALTH_URL = f"http://127.0.0.1:{API_PORT}/health"
STARTUP_TIMEOUT = 30  # seconds to wait for the API to come up


def port_in_use(port: str) -> bool:
    """Return True if something is already listening on the local port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(1)
        return sock.connect_ex(("127.0.0.1", int(port))) == 0


def wait_for_api(api: subprocess.Popen) -> bool:
    """Wait until the API answers its health check, or the process exits."""
    deadline = time.time() + STARTUP_TIMEOUT
    while time.time() < deadline:
        if api.poll() is not None:
            return False
        try:
            with urllib.request.urlopen(HEALTH_URL, timeout=2):
                return True
        except OSError:
            time.sleep(0.5)
    return False


def stop(process: subprocess.Popen) -> None:
    """Terminate a child process if it is still running."""
    if process.poll() is None:
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()


def main() -> int:
    busy = [port for port in (API_PORT, UI_PORT) if port_in_use(port)]
    if busy:
        print(
            f"Port {' and '.join(busy)} already in use. The application may already be running:\n"
            "close the other terminal(s) running it, or set API_PORT / UI_PORT to free ports.",
            file=sys.stderr,
        )
        return 1

    api = subprocess.Popen(
        [sys.executable, "-m", "uvicorn", "app.main:app", "--port", API_PORT],
        cwd=BASE_DIR,
    )
    ui = None
    try:
        if not wait_for_api(api):
            print("The API did not start. See the log above.", file=sys.stderr)
            return 1

        ui = subprocess.Popen(
            [sys.executable, "-m", "streamlit", "run", "frontend/streamlit_app.py",
             "--server.port", UI_PORT],
            cwd=BASE_DIR,
        )
        print(f"\nAPI: http://localhost:{API_PORT}/docs")
        print(f"App: http://localhost:{UI_PORT}  (Ctrl+C to stop)\n")

        # Exit as soon as either process stops
        while api.poll() is None and ui.poll() is None:
            time.sleep(1)
        return 0
    except KeyboardInterrupt:
        return 0
    finally:
        if ui is not None:
            stop(ui)
        stop(api)


if __name__ == "__main__":
    sys.exit(main())
