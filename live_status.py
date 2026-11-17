"""
Live status layer for the Stock Scanner  (powers the Bloomberg-style Live Dashboard)

What it does
------------
Scanner.py calls small hooks (job_update / progress / match / event ...). This module keeps an in-memory
copy of everything the MASTER DASHBOARD.xlsx shows (jobs, status, progress, next run, matrix status,
log, config) and updates it on EVERY stock instead of only when the workbook is saved. The state is

  1. written atomically to  <LIVE_DIR>/live_status.json   (read by dashboard_server.py on the same machine), and
  2. optionally published to the `live-status` branch of the GitHub repo (file live_status.py -> live_status.json)
     so the dashboard can be viewed while the scan runs on a GitHub Actions runner.

Everything is OPT-IN and can never break a scan: with LIVE_DASHBOARD unset/0 the scanner uses NullLive
(all hooks are no-ops), and every network/file error inside this module is caught and only printed.

Environment variables
---------------------
LIVE_DASHBOARD        1 to enable (default 0)
LIVE_DIR              folder for the local JSON file (default "live")
LIVE_PUBLISH          1/0 publish to GitHub branch (default 1 on GitHub Actions, else 0)
LIVE_BRANCH           branch that receives the JSON (default "live-status"; created automatically)
LIVE_PUBLISH_INTERVAL minimum seconds between publishes while things change (default 15)
GITHUB_TOKEN, GITHUB_REPOSITORY   provided by the workflow (token needs `contents: write`)

Secrets are never put in the state: config keys containing token / password / secret / matrix / email / smtp
are dropped.
"""
import base64
import copy
import json
import os
import threading
import time
from datetime import datetime
from datetime import time as dtime

import requests

SCHEMA = 1
SECRET_WORDS = ("token", "password", "secret", "matrix", "email", "smtp")
STATE_FILE = "live_status.json"
MAX_EVENTS = 80
MAX_MATCHES = 30
MAX_LOG = 40


# ----------------------------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------------------------
def _yes(value):
    return str(value).strip().upper() in {"YES", "Y", "TRUE", "1", "ON", "ENABLED"}


def _num(value):
    try:
        return float(value) if value is not None and str(value).strip() != "" else None
    except (TypeError, ValueError):
        return None


def _iso(value, tz):
    """cell value / datetime -> ISO string with the scanner's timezone offset (or None)."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        dt = value
    else:
        try:
            dt = datetime.fromisoformat(str(value).strip())
        except ValueError:
            return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=tz)
    return dt.astimezone(tz).isoformat(timespec="seconds")


def _hhmm(value):
    if isinstance(value, (dtime, datetime)):
        return f"{value.hour:02d}:{value.minute:02d}"
    return str(value).strip() if value not in (None, "") else ""


def _text(value, limit=300):
    return "" if value is None else str(value)[:limit]


def _summary(jobs):
    def count(word):
        return sum(1 for j in jobs if word in str(j.get("status") or "").lower())
    return {
        "enabled": sum(1 for j in jobs if j.get("enabled")),
        "in_progress": count("in progress"),
        "completed": count("completed"),
        "scheduled": count("scheduled"),
        "failed": count("failed") + count("interrupted"),
        "jobs": len(jobs),
    }


# ----------------------------------------------------------------------------------------------
# Baseline snapshot = what MASTER DASHBOARD.xlsx currently says
# ----------------------------------------------------------------------------------------------
def snapshot_from_workbook(wb, tz):
    """Build the live state from an openpyxl workbook (Schedule / Dashboard / Log / Config / Watchlist)."""
    sched = wb["Schedule"]
    dash = wb["Dashboard"] if "Dashboard" in wb.sheetnames else None
    jobs = []
    for r in range(2, 22):                       # rows 2-21 are the job table (below is the legend)
        name = sched.cell(r, 2).value
        if not name:
            continue

        def g(c, r=r):
            return sched.cell(r, c).value

        matrix = None
        if dash is not None and 7 <= r + 5 <= 18:        # Dashboard row for Schedule row r is r+5, column G
            v = dash.cell(r + 5, 7).value
            if isinstance(v, str) and v.strip() and not v.startswith("="):
                matrix = v.strip()
        progress = _num(g(18))
        jobs.append({
            "row": r, "name": str(name).strip(), "enabled": _yes(g(1)),
            "timeframe": _text(g(3)), "trigger": _text(g(4)), "run_time": _hhmm(g(5)),
            "weekday": _text(g(6)), "day": _text(g(7)),
            "flags": {"rsi": _yes(g(8)), "ema": _yes(g(9)), "patterns": _yes(g(10)),
                      "divergence": _yes(g(11)), "multi_tf": _yes(g(12))},
            "output_rows": _text(g(13)),
            "status": _text(g(17)), "progress": progress if progress is not None else None,
            "next_run": _iso(g(14), tz), "last_started": _iso(g(15), tz),
            "last_completed": _iso(g(16), tz),
            "matrix_status": matrix, "details": _text(g(19), 500),
            "done": None, "total": None, "matched": None, "failed": None, "excluded": None,
            "phase": None, "symbol": None, "rate": None, "eta_sec": None,
        })

    log = []
    if "Log" in wb.sheetnames:
        ws = wb["Log"]
        for r in range(ws.max_row, 1, -1):
            t = ws.cell(r, 1).value
            if t in (None, ""):
                continue
            log.append({"time": _iso(t, tz), "job": _text(ws.cell(r, 2).value),
                        "timeframe": _text(ws.cell(r, 3).value),
                        "scanned": ws.cell(r, 4).value, "matched": ws.cell(r, 5).value,
                        "result": _text(ws.cell(r, 6).value)})
            if len(log) >= MAX_LOG:
                break

    config, rsi_ranges, report_name = {}, [], ""
    if "Config" in wb.sheetnames:
        cws = wb["Config"]
        for r in range(2, cws.max_row + 1):
            k = cws.cell(r, 1).value
            if not k or any(w in str(k).lower() for w in SECRET_WORDS):
                continue
            config[str(k).strip()] = _text(cws.cell(r, 2).value, 80)
        report_name = config.get("Report Name", "")
        for r in range(2, 8):
            tf = cws.cell(r, 5).value
            if tf:
                rsi_ranges.append({"timeframe": str(tf), "active": _yes(cws.cell(r, 6).value),
                                   "min": _num(cws.cell(r, 7).value), "max": _num(cws.cell(r, 8).value)})

    watchlist = 0
    if "Watchlist" in wb.sheetnames:
        watchlist = sum(1 for (v,) in wb["Watchlist"].iter_rows(min_row=2, max_col=1, values_only=True) if v)

    return {
        "schema": SCHEMA, "report_name": report_name, "run_active": False, "run": {},
        "summary": _summary(jobs), "jobs": jobs, "matches": [], "events": [],
        "log": log, "config": config, "rsi_ranges": rsi_ranges, "watchlist_size": watchlist,
    }


# ----------------------------------------------------------------------------------------------
# GitHub publisher (single JSON file on an auto-created branch, via the Contents API)
# ----------------------------------------------------------------------------------------------
class _GitHubPublisher:
    API = "https://api.github.com"

    def __init__(self, repo, token, branch):
        self.repo, self.token, self.branch = repo, token, branch
        self.sha = None
        self.ready = False
        self.fail_streak = 0
        self.last_error = None

    def _h(self):
        return {"Authorization": f"Bearer {self.token}", "Accept": "application/vnd.github+json",
                "X-GitHub-Api-Version": "2022-11-28"}

    def _req(self, method, path, **kw):
        return requests.request(method, f"{self.API}/repos/{self.repo}/{path}", headers=self._h(),
                                timeout=15, **kw)

    def _ensure_branch(self, text):
        r = self._req("GET", f"git/ref/heads/{self.branch}")
        if r.status_code == 200:
            self._learn_sha()
            self.ready = True
            return
        if r.status_code != 404:
            raise RuntimeError(f"branch lookup HTTP {r.status_code}: {r.text[:160]}")
        # create an orphan branch holding just the JSON file
        blob = self._req("POST", "git/blobs", json={"content": text, "encoding": "utf-8"})
        blob.raise_for_status()
        tree = self._req("POST", "git/trees", json={"tree": [
            {"path": STATE_FILE, "mode": "100644", "type": "blob", "sha": blob.json()["sha"]}]})
        tree.raise_for_status()
        commit = self._req("POST", "git/commits", json={"message": "live status: init", "tree": tree.json()["sha"],
                                                        "parents": []})
        commit.raise_for_status()
        ref = self._req("POST", "git/refs", json={"ref": f"refs/heads/{self.branch}", "sha": commit.json()["sha"]})
        ref.raise_for_status()
        self.sha = blob.json()["sha"]
        self.ready = True

    def _learn_sha(self):
        r = self._req("GET", f"contents/{STATE_FILE}", params={"ref": self.branch})
        self.sha = r.json().get("sha") if r.status_code == 200 else None

    def publish(self, text):
        try:
            if not self.ready:
                self._ensure_branch(text)
            for attempt in range(2):
                body = {"message": "live status", "branch": self.branch,
                        "content": base64.b64encode(text.encode("utf-8")).decode("ascii")}
                if self.sha:
                    body["sha"] = self.sha
                r = self._req("PUT", f"contents/{STATE_FILE}", json=body)
                if r.status_code in (200, 201):
                    self.sha = r.json()["content"]["sha"]
                    self.fail_streak, self.last_error = 0, None
                    return True
                if r.status_code in (409, 422) and attempt == 0:
                    self._learn_sha()                 # stale sha -> refresh and retry once
                    continue
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:160]}")
        except Exception as exc:
            self.fail_streak += 1
            msg = f"{type(exc).__name__}: {str(exc)[:200]}"
            if msg != self.last_error:
                print(f"[live] publish to branch '{self.branch}' failed ({msg})", flush=True)
            self.last_error = msg
        return False


# ----------------------------------------------------------------------------------------------
# Live state
# ----------------------------------------------------------------------------------------------
class NullLive:
    """Used when the live dashboard is switched off: every hook is a no-op."""
    enabled = False
    idle_run = False

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        return lambda *a, **k: None


class LiveStatus:
    enabled = True
    idle_run = False          # set by Scanner.py when nothing was due (skip the final GitHub publish)

    def __init__(self, tz, local_dir, publisher=None, interval=15.0):
        self.tz = tz
        self.dir = local_dir
        self.pub = publisher
        self.interval = max(5.0, float(interval))
        self.lock = threading.RLock()
        self.state = None
        self.jobs = {}
        self.events, self.matches = [], []
        self.active = False
        self.run = {}
        self.rev = 0
        self._t0 = {}
        self._stop = threading.Event()
        self._wake = threading.Event()
        self._thread = None

    # ---- lifecycle ---------------------------------------------------------------------------
    def start(self, base_state, trigger="manual"):
        with self.lock:
            self.state = base_state
            self.jobs = {j["row"]: j for j in base_state["jobs"]}
            self.active = True
            self.run = {"id": os.getenv("GITHUB_RUN_ID") or "local", "trigger": trigger,
                        "started": datetime.now(self.tz).isoformat(timespec="seconds"),
                        "host": "github-actions" if os.getenv("GITHUB_ACTIONS") else "local"}
            self._bump()
        self.event(f"Scanner run started ({trigger})")
        self._write_local()
        if self.pub:
            self.pub.publish(self._dump())
        self._thread = threading.Thread(target=self._loop, name="live-status", daemon=True)
        self._thread.start()
        print(f"[live] live dashboard feed on (local file: {os.path.join(self.dir, STATE_FILE)}"
              f"{', GitHub branch: ' + self.pub.branch if self.pub else ''})", flush=True)

    def close(self, final_publish=True):
        if self.state is None:
            return
        self._stop.set()
        self._wake.set()
        if self._thread:
            self._thread.join(timeout=5)
        with self.lock:
            self.active = False
            self._bump()
        self._write_local()
        if self.pub and final_publish:
            self.pub.publish(self._dump())

    # ---- hooks called by Scanner.py -------------------------------------------------------------
    def job_update(self, row, status, progress, details, started=None, completed=None, next_run=None):
        """Mirror of update_job_status(): same data the workbook's Schedule sheet receives."""
        with self.lock:
            j = self.jobs.get(row)
            if j is None:
                return
            low = str(status or "").lower()
            j["status"], j["details"] = _text(status), _text(details, 500)
            j["progress"] = _num(progress)
            if started:
                j["last_started"] = _iso(started, self.tz)
            if completed:
                j["last_completed"] = _iso(completed, self.tz)
            if next_run is not None:
                j["next_run"] = _iso(next_run, self.tz)
            if "in progress" in low:
                if row not in self._t0:               # first call of a run
                    self._t0[row] = time.time()
                    j.update(done=0, total=None, matched=0, failed=0, excluded=0, rate=None,
                             eta_sec=None, phase="Starting", symbol=None)
            else:
                self._t0.pop(row, None)
                j.update(phase=None, symbol=None, eta_sec=None)
            self._bump()

    def progress(self, row, done, total, matched, failed=0, excluded=0, symbol=None):
        with self.lock:
            j = self.jobs.get(row)
            if j is None:
                return
            j.update(done=done, total=total, matched=matched, failed=failed, excluded=excluded,
                     symbol=symbol, progress=(done / total) if total else 0.0, phase="Scanning")
            t0 = self._t0.get(row)
            if t0 and done:
                rate = done / max(time.time() - t0, 1e-6)
                j["rate"] = round(rate, 2)
                j["eta_sec"] = int((total - done) / rate) if rate > 0 and total else None
            self._bump()

    def set_phase(self, row, text):
        with self.lock:
            j = self.jobs.get(row)
            if j is not None:
                j["phase"] = _text(text, 120)
                self._bump()

    def set_matrix(self, row, text):
        with self.lock:
            j = self.jobs.get(row)
            if j is not None:
                j["matrix_status"] = _text(text, 120)
                self._bump()

    def match(self, row, symbol, ltp, rsi):
        with self.lock:
            j = self.jobs.get(row)
            self.matches.append({"t": datetime.now(self.tz).isoformat(timespec="seconds"),
                                 "job": j["name"] if j else "", "symbol": symbol, "ltp": ltp, "rsi": rsi})
            del self.matches[:-MAX_MATCHES]
            self._bump()

    def event(self, text, level="info"):
        with self.lock:
            self.events.append({"t": datetime.now(self.tz).isoformat(timespec="seconds"),
                                "level": level, "msg": _text(text, 300)})
            del self.events[:-MAX_EVENTS]
            self._bump()

    def log_row(self, values):
        """values = the list appended to the workbook's Log sheet."""
        with self.lock:
            t, job, tf, scanned, matched, result = (list(values) + [None] * 6)[:6]
            self.state["log"].insert(0, {"time": _iso(t, self.tz), "job": _text(job), "timeframe": _text(tf),
                                         "scanned": scanned, "matched": matched, "result": _text(result)})
            del self.state["log"][MAX_LOG:]
            self._bump()

    # ---- state / output --------------------------------------------------------------------------
    def _bump(self):
        self.rev += 1
        self._wake.set()

    def snapshot(self):
        with self.lock:
            s = copy.deepcopy(self.state)
            s["events"] = list(self.events)[::-1]          # newest first
            s["matches"] = list(self.matches)[::-1]
            s["run"], s["run_active"], s["seq"] = dict(self.run), self.active, self.rev
        s["updated"] = datetime.now(self.tz).isoformat(timespec="seconds")
        s["summary"] = _summary(s["jobs"])
        return s

    def _dump(self):
        return json.dumps(self.snapshot(), ensure_ascii=False, separators=(",", ":"), default=str)

    def _write_local(self):
        try:
            os.makedirs(self.dir, exist_ok=True)
            path = os.path.join(self.dir, STATE_FILE)
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(self._dump())
            os.replace(tmp, path)                       # atomic: readers never see a half-written file
        except Exception as exc:
            print(f"[live] local write failed: {type(exc).__name__}: {exc}", flush=True)

    def _loop(self):
        last_local = last_remote = time.time()
        seen_local = seen_remote = self.rev
        while not self._stop.is_set():
            self._wake.wait(1.0)
            self._wake.clear()
            if self._stop.is_set():
                break
            now, rev = time.time(), self.rev
            if (rev != seen_local and now - last_local >= 0.5) or now - last_local >= 5:
                self._write_local()
                last_local, seen_local = now, rev
            if self.pub and self.pub.fail_streak < 8:
                if (rev != seen_remote and now - last_remote >= self.interval) or now - last_remote >= 60:
                    self.pub.publish(self._dump())          # heartbeat every 60 s even if nothing changed
                    last_remote, seen_remote = now, rev
            elif self.pub and self.pub.fail_streak >= 8:
                print("[live] too many publish failures - GitHub publishing disabled for this run "
                      "(local file keeps updating)", flush=True)
                self.pub = None


def create(tz):
    """Factory used by Scanner.py: NullLive unless LIVE_DASHBOARD=1."""
    if not _yes(os.getenv("LIVE_DASHBOARD", "0")):
        return NullLive()
    publisher = None
    token, repo = os.getenv("GITHUB_TOKEN", "").strip(), os.getenv("GITHUB_REPOSITORY", "").strip()
    default_pub = "1" if os.getenv("GITHUB_ACTIONS") else "0"
    if _yes(os.getenv("LIVE_PUBLISH", default_pub)):
        if token and repo:
            publisher = _GitHubPublisher(repo, token, os.getenv("LIVE_BRANCH", "live-status").strip() or "live-status")
        else:
            print("[live] LIVE_PUBLISH is on but GITHUB_TOKEN / GITHUB_REPOSITORY are missing - "
                  "writing the local file only", flush=True)
    return LiveStatus(tz, os.getenv("LIVE_DIR", "live"), publisher,
                      float(os.getenv("LIVE_PUBLISH_INTERVAL", "15") or 15))
