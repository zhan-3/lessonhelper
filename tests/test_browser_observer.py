import json
import socket
import subprocess
import tempfile
import time
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from playwright.sync_api import sync_playwright

from course_selection.browser_observer import BorrowedBrowserObserver


class FixtureHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        body = b"<title>fixture</title><p>still alive</p>"
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *_args):
        return


def free_port():
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def test_borrowed_observer_inspects_and_detaches_without_closing_browser():
    cdp_port, site_port = free_port(), free_port()
    server = ThreadingHTTPServer(("127.0.0.1", site_port), FixtureHandler)
    server_thread = __import__("threading").Thread(target=server.serve_forever, daemon=True)
    server_thread.start()
    with sync_playwright() as playwright:
        executable = Path(playwright.chromium.executable_path)
    # Launch outside Playwright: the observer must attach to an independently
    # owned browser, not to a browser object from the same sync API instance.
    profile_directory = tempfile.TemporaryDirectory()
    profile = Path(profile_directory.name)
    process = subprocess.Popen([
        str(executable), "--headless=new", "--no-sandbox",
        f"--remote-debugging-port={cdp_port}", "--user-data-dir=" + str(profile),
        f"http://127.0.0.1:{site_port}/fixture",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    observer = BorrowedBrowserObserver(session_id="test")
    try:
        endpoint = f"http://127.0.0.1:{cdp_port}/json/version"
        for _ in range(50):
            try:
                urllib.request.urlopen(endpoint, timeout=0.2).read()
                break
            except OSError:
                time.sleep(0.1)
        assert observer.connect(endpoint).status == "connected"
        first = observer.inspect().to_dict()
        assert first["data"]["connection"] == "borrowed"
        assert first["data"]["target_count"] >= 1
        assert "fixture" not in json.dumps(first)
        assert observer.inspect().to_dict() == first
        assert observer.disconnect().status == "disconnected"
        assert process.poll() is None
        targets = json.loads(urllib.request.urlopen(f"http://127.0.0.1:{cdp_port}/json/list").read())
        assert targets
    finally:
        observer.disconnect()
        process.terminate()
        process.wait(timeout=5)
        server.shutdown()
        server.server_close()
        profile_directory.cleanup()
