import ctypes
import inspect
import logging
import os
import socket
import threading
import time
from pathlib import Path

import requests
import uvicorn
import webview

from app.runtime import create_app, load_runtime_config


logger = logging.getLogger("ms_mail_fetcher.desktop")
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


DEFAULT_DESKTOP_HOST = "127.0.0.1"
DEFAULT_WAIT_SECONDS = 20
WINDOW_TITLE = "MS Mail Fetcher"
SINGLE_INSTANCE_MUTEX_NAME = "Local\\MS_MAIL_FETCHER_DESKTOP_SINGLE_INSTANCE"


class _SingleInstanceGuard:
    def __init__(self, name: str):
        self._name = name
        self._handle = None

    def acquire(self) -> bool:
        kernel32 = ctypes.windll.kernel32
        self._handle = kernel32.CreateMutexW(None, False, self._name)
        if not self._handle:
            raise RuntimeError("Failed to create application mutex.")

        ERROR_ALREADY_EXISTS = 183
        last_error = kernel32.GetLastError()
        return last_error != ERROR_ALREADY_EXISTS

    def release(self) -> None:
        if self._handle:
            ctypes.windll.kernel32.CloseHandle(self._handle)
            self._handle = None


def _is_port_available(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind((host, port))
        except OSError:
            return False
    return True


def _resolve_desktop_bind() -> tuple[str, int]:
    config, _ = load_runtime_config()
    preferred_port = int(config.get("port", 18765))
    retry_count = max(0, int(config.get("port_retry_count", 20)))

    for candidate in range(preferred_port, preferred_port + retry_count + 1):
        if _is_port_available(DEFAULT_DESKTOP_HOST, candidate):
            return DEFAULT_DESKTOP_HOST, candidate

    raise RuntimeError(
        f"No available desktop server port from {preferred_port} to {preferred_port + retry_count}."
    )


def _wait_until_ready(url: str, timeout_seconds: int = DEFAULT_WAIT_SECONDS) -> None:
    deadline = time.time() + timeout_seconds
    health_url = f"{url}/api/health"

    while time.time() < deadline:
        try:
            response = requests.get(health_url, timeout=1.5)
            if response.ok:
                return
        except requests.RequestException:
            pass
        time.sleep(0.3)

    raise RuntimeError("Backend did not become ready in time.")


def _run_server(server: uvicorn.Server) -> None:
    server.run()


def _prepare_webview_storage() -> Path:
    """Ensure WebView2 uses a stable user data directory so localStorage persists."""
    base = Path(os.getenv("LOCALAPPDATA") or Path.home()) / "ms-mail-fetcher" / "webview2"
    base.mkdir(parents=True, exist_ok=True)
    os.environ.setdefault("WEBVIEW2_USER_DATA_FOLDER", str(base))
    return base


def main() -> None:
    guard = _SingleInstanceGuard(SINGLE_INSTANCE_MUTEX_NAME)
    if not guard.acquire():
        logger.info("Another app instance is already running. Exit.")
        return
    server = None
    server_thread = None
    try:
        host, port = _resolve_desktop_bind()
        api_url = f"http://{host}:{port}"
        storage_path = _prepare_webview_storage()

        app = create_app()
        app.state.server_host = host
        app.state.server_port = port

        config = uvicorn.Config(
            app,
            host=host,
            port=port,
            reload=False,
            access_log=False,
        )
        server = uvicorn.Server(config)

        server_thread = threading.Thread(
            target=_run_server,
            args=(server,),
            daemon=True,
            name="ms-mail-fetcher-api",
        )
        server_thread.start()

        _wait_until_ready(api_url)
        logger.info("Desktop UI opening: %s", api_url)

        create_window_params = inspect.signature(webview.create_window).parameters
        window_kwargs = {
            "title": WINDOW_TITLE,
            "url": api_url,
            "min_size": (1100, 760),
            "width": 1280,
            "height": 860,
        }

        if "private_mode" in create_window_params:
            window_kwargs["private_mode"] = False
        if "storage_path" in create_window_params:
            window_kwargs["storage_path"] = str(storage_path)

        webview.create_window(**window_kwargs)
        webview.start()
    finally:
        if server is not None:
            logger.info("Desktop window closed, stopping backend...")
            server.should_exit = True
        if server_thread is not None:
            server_thread.join(timeout=8)
        guard.release()


if __name__ == "__main__":
    main()
