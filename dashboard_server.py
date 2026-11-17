#!/usr/bin/env python3
"""
Live Dashboard server  -  serves the Bloomberg-style UI (dashboard/index.html) and streams the scanner state.

    python dashboard_server.py                      # http://127.0.0.1:8080  (reads live/live_status.json,
                                                    #  falls back to MASTER DASHBOARD.xlsx when no scan is running)
    python dashboard_server.py --source github      # follow the `live-status` branch of the repo checkout in "."
                                                    #  (used by the 'Live Dashboard' GitHub workflow)

Endpoints:  /  (UI)   /api/state (JSON)   /events (Server-Sent Events: "state" on change, "tick" every 2 s)
            /healthz
Optional access key (DASH_KEY or --key): every request except /healthz must carry ?key=<key>.

Standard library only (openpyxl is used just for the workbook fallback).
"""
import argparse
import hmac
import json
import os
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse
from zoneinfo import ZoneInfo

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
TZ = ZoneInfo(os.getenv("SCANNER_TIMEZONE", "Asia/Kolkata"))


class Provider:
    """Chooses the freshest available state: live file / GitHub branch  ->  workbook fallback."""

    def __init__(self, args):
        self.a = args
        self.remote_text = None
        self.remote_err = None
        self._wb_cache = (None, None)
        self._lock = threading.Lock()
        if args.source == "github":
            threading.Thread(target=self._poll_git, name="git-poll", daemon=True).start()

    # -- GitHub branch (fetch is a plain git operation, so it is not subject to REST rate limits) --
    def _poll_git(self):
        ref = f"refs/remotes/origin/{self.a.branch}"
        while True:
            try:
                subprocess.run(["git", "-C", self.a.repo_dir, "fetch", "--quiet", "--depth", "1", "origin",
                                f"+refs/heads/{self.a.branch}:{ref}"], check=True, capture_output=True, timeout=60)
                out = subprocess.run(["git", "-C", self.a.repo_dir, "show", f"{ref}:live_status.json"],
                                     check=True, capture_output=True, timeout=30).stdout.decode("utf-8")
                json.loads(out)
                with self._lock:
                    self.remote_text, self.remote_err = out, None
            except Exception as exc:
                detail = getattr(exc, "stderr", b"") or b""
                with self._lock:
                    self.remote_err = (detail.decode("utf-8", "replace") or str(exc))[:200]
            time.sleep(self.a.git_interval)

    def _live_file(self):
        path = os.path.join(self.a.live_dir, "live_status.json")
        try:
            with open(path, encoding="utf-8") as f:
                text = f.read()
            json.loads(text)
            return text
        except (OSError, ValueError):
            return None

    def _from_workbook(self):
        import openpyxl
        import live_status
        path = self.a.workbook
        if not os.path.exists(path) and os.path.exists(path.replace(" ", "_")):
            path = path.replace(" ", "_")
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            return json.dumps({"error": f"workbook not found: {self.a.workbook}", "jobs": [], "summary": {}})
        if self._wb_cache[0] == mtime:
            return self._wb_cache[1]
        try:
            snap = live_status.snapshot_from_workbook(openpyxl.load_workbook(path), TZ)
            snap["updated"] = datetime.fromtimestamp(mtime, TZ).isoformat(timespec="seconds")
            snap["from_workbook"] = True
            text = json.dumps(snap, ensure_ascii=False, default=str)
        except Exception as exc:                  # e.g. the file is being rewritten right now
            if self._wb_cache[1]:
                return self._wb_cache[1]
            return json.dumps({"error": f"{type(exc).__name__}: {exc}", "jobs": [], "summary": {}})
        self._wb_cache = (mtime, text)
        return text

    def snapshot(self):
        """-> (raw_text, payload_dict). raw_text is stable between changes (used for change detection)."""
        with self._lock:
            text = self.remote_text if self.a.source == "github" else None
            err = self.remote_err
        if self.a.source != "github":
            text = self._live_file()
        if text is None:
            text = self._from_workbook()
        data = json.loads(text)
        data["server_time"] = datetime.now(TZ).isoformat(timespec="seconds")
        data["source"] = "github-branch" if (self.a.source == "github" and self.remote_text) else (
            "live-file" if not data.get("from_workbook") else "workbook")
        if err and self.a.source == "github" and not self.remote_text:
            data["source_note"] = "live-status branch not readable yet: " + err
        return text, data


class Handler(BaseHTTPRequestHandler):
    server_version = "ScannerDashboard/1.0"

    def log_message(self, *args):
        pass

    def _send(self, code, body, ctype, extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        url = urlparse(self.path)
        query = parse_qs(url.query)
        if url.path == "/healthz":
            return self._send(200, "ok", "text/plain")
        key = self.server.key
        if key and not hmac.compare_digest(query.get("key", [""])[0], key):
            return self._send(401, "Access key required: open the URL that ends with ?key=...", "text/plain")
        if url.path in ("/", "/index.html"):
            with open(os.path.join(HERE, "dashboard", "index.html"), "rb") as f:
                return self._send(200, f.read(), "text/html; charset=utf-8")
        if url.path == "/api/state":
            return self._send(200, json.dumps(self.server.provider.snapshot()[1], ensure_ascii=False),
                              "application/json; charset=utf-8")
        if url.path == "/events":
            return self._stream()
        self._send(404, "not found", "text/plain")

    def _stream(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache, no-transform")
        self.send_header("X-Accel-Buffering", "no")
        self.send_header("Connection", "keep-alive")
        self.end_headers()
        last_raw, last_tick = None, 0.0
        try:
            while True:
                raw, data = self.server.provider.snapshot()
                if raw != last_raw:
                    self.wfile.write(f"event: state\ndata: {json.dumps(data, ensure_ascii=False)}\n\n".encode())
                    last_raw, last_tick = raw, time.time()
                elif time.time() - last_tick >= 2:
                    self.wfile.write(f"event: tick\ndata: {json.dumps({'server_time': data['server_time']})}\n\n"
                                     .encode())
                    last_tick = time.time()
                self.wfile.flush()
                time.sleep(0.7)
        except (BrokenPipeError, ConnectionResetError, OSError):
            return


def main():
    env = os.getenv
    ap = argparse.ArgumentParser(description="Live Dashboard server for the Stock Scanner")
    ap.add_argument("--host", default=env("DASH_HOST", "127.0.0.1"))
    ap.add_argument("--port", type=int, default=int(env("DASH_PORT", "8080")))
    ap.add_argument("--source", choices=["local", "github"], default=env("DASH_SOURCE", "local"))
    ap.add_argument("--workbook", default=env("SCANNER_TEMPLATE", "MASTER DASHBOARD.xlsx"))
    ap.add_argument("--live-dir", default=env("LIVE_DIR", "live"))
    ap.add_argument("--repo-dir", default=env("DASH_REPO_DIR", "."))
    ap.add_argument("--branch", default=env("LIVE_BRANCH", "live-status"))
    ap.add_argument("--git-interval", type=float, default=float(env("DASH_GIT_INTERVAL", "4")))
    ap.add_argument("--key", default=env("DASH_KEY", ""))
    ap.add_argument("--open", action="store_true", help="open the dashboard in your browser")
    args = ap.parse_args()

    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    srv.daemon_threads = True
    srv.key, srv.provider = args.key, Provider(args)
    url = f"http://{'127.0.0.1' if args.host in ('0.0.0.0', '') else args.host}:{args.port}/"
    url += f"?key={args.key}" if args.key else ""
    print(f"Live Dashboard  ->  {url}   (source: {args.source}; Ctrl+C to stop)", flush=True)
    if args.open:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nstopped")


if __name__ == "__main__":
    main()
