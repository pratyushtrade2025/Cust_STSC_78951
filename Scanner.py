"""
Stock Scanner for MASTER DASHBOARD.xlsx

Flow
----
1. Reads Config, Schedule, Watchlist and Patterns from the workbook.
2. Runs every ENABLED job in the Schedule sheet that is due (or all enabled jobs on a manual run).
3. Scans the watchlist. STOCKS ARE SELECTED ONLY BY THE RSI RANGES in the Config sheet
   (RSI Min / RSI Max of the Active timeframes). EMA, patterns and divergence are informational.
4. Builds a clean, standalone report workbook (Scan_Results_*.xlsx):
      Output Rows = ALL           -> Symbol, Name, LTP, RSI, EMA columns, Patterns, RSI Divergence
      Output Rows = SIGNALS ONLY  -> Symbol, Name, LTP, RSI, Patterns, RSI Divergence
5. Sends the report to Matrix and/or e-mail (switches in the Config sheet).
6. Updates Schedule / Dashboard / Log inside MASTER DASHBOARD.xlsx (the workflow commits it back).

Secrets / environment variables
-------------------------------
Matrix  : MATRIX_HOMESERVER, MATRIX_ROOM_ID, MATRIX_ACCESS_TOKEN  (+ optional MATRIX_REFRESH_TOKEN)
          The working token is REMEMBERED between runs in .matrix_state/matrix_token.json (the workflow keeps that
          folder in the GitHub Actions cache). Order each run: this run's token -> secret token -> remembered token
          -> refresh with the refresh token (the new pair is stored again) -> password login if MATRIX_USER +
          MATRIX_PASSWORD exist (only for accounts that have a password; Google sign-in accounts do not).
          Env wins; the Config sheet is a fallback for homeserver, room and 'Bot Matrix ID' (user) only.
          Keep tokens / passwords in GitHub Secrets, NOT in the workbook.
E-mail  : SMTP_HOST, SMTP_PORT (465 SSL or 587 STARTTLS), SMTP_USER, SMTP_PASSWORD,
          optional EMAIL_FROM, EMAIL_TO (recipient comes from Config 'Email ID' if filled)
Run     : RUN_ALL_ENABLED=1 -> run every enabled job now;  JOB_NAME=<name> -> restrict to that job
Live    : LIVE_DASHBOARD=1 -> stream real-time progress (per stock) for the Bloomberg-style Live Dashboard
          (see live_status.py and dashboard_server.py). Off by default; never affects the scan itself.
"""
import os
import re
import json
import warnings
import ssl
import time
import math
import uuid
import random
import smtplib
import threading
import urllib.parse
from datetime import datetime, timedelta, date, time as dtime
from email.message import EmailMessage
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed

import numpy as np
import pandas as pd
import yfinance as yf
import talib
import openpyxl
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import live_status

# yfinance emits thousands of harmless pandas deprecation warnings (Pandas4Warning) that bury the real log.
# Libraries can reset warning filters, so mute at the display level AFTER all imports.
warnings.filterwarnings("ignore")
warnings.showwarning = lambda *args, **kwargs: None

# ----------------------------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------------------------
TEMPLATE_FILE = os.getenv("SCANNER_TEMPLATE", "MASTER DASHBOARD.xlsx")
TIMEZONE = ZoneInfo(os.getenv("SCANNER_TIMEZONE", "Asia/Kolkata"))
MAX_WORKERS = int(os.getenv("MAX_WORKERS", "6"))
REQUEST_INTERVAL = float(os.getenv("REQUEST_INTERVAL", "0.35"))
MAX_RETRIES = int(os.getenv("MAX_RETRIES", "4"))
RETRY_BASE_SLEEP = int(os.getenv("RETRY_BASE_SLEEP", "4"))
RETRY_ROUNDS = int(os.getenv("RETRY_ROUNDS", "1"))
RETRY_COOLDOWN = int(os.getenv("RETRY_COOLDOWN", "45"))
REPORT_DIR = os.getenv("REPORT_DIR", "reports")
MARKET_CLOSE = dtime(15, 45)  # NSE closes 15:30; small buffer for final data

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
RSI_TFS = ["4H", "Daily", "Weekly", "Monthly", "Quarterly", "Half-Yearly"]
HIST_PERIOD = {"4H": "3y", "Daily": "3y", "Weekly": "6y", "Monthly": "10y",
               "Quarterly": "max", "Half-Yearly": "max"}
HIST_RANK = {"3y": 0, "6y": 1, "10y": 2, "max": 3}
AGG = {"Open": "first", "High": "max", "Low": "min", "Close": "last"}
EXCLUDED = "Excluded by RSI filter"
INSUFFICIENT = "Insufficient history"
STALE = "Stale data (no recent trades)"
NO_RETRY = (INSUFFICIENT, STALE)

THROTTLE_LOCK = threading.Lock()
NEXT_SLOT = 0.0

# Real-time feed for the Live Dashboard. NullLive (all no-ops) until main() switches it on.
LIVE = live_status.NullLive()

PATTERN_FUNCS = {
    "Doji": "CDLDOJI",
    "Dragonfly Doji": "CDLDRAGONFLYDOJI",
    "Gravestone Doji": "CDLGRAVESTONEDOJI",
    "Hammer": "CDLHAMMER",
    "Inverted Hammer": "CDLINVERTEDHAMMER",
    "Hanging Man": "CDLHANGINGMAN",
    "Shooting Star": "CDLSHOOTINGSTAR",
    "Bullish Engulfing": "CDLENGULFING",
    "Bearish Engulfing": "CDLENGULFING",
    "Piercing Line": "CDLPIERCING",
    "Dark Cloud Cover": "CDLDARKCLOUDCOVER",
    "Bullish Harami": "CDLHARAMI",
    "Bearish Harami": "CDLHARAMI",
    "Morning Star": "CDLMORNINGSTAR",
    "Evening Star": "CDLEVENINGSTAR",
    "Bullish Marubozu": "CDLMARUBOZU",
    "Bearish Marubozu": "CDLMARUBOZU",
    "Three White Soldiers": "CDL3WHITESOLDIERS",
    "Three Black Crows": "CDL3BLACKCROWS",
    "Spinning Top": "CDLSPINNINGTOP",
}


# ----------------------------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------------------------
def yes(value):
    return str(value).strip().upper() in {"YES", "Y", "TRUE", "1", "ON", "ENABLED"}


def as_float(value, default=None):
    try:
        return float(value) if value is not None and str(value).strip() != "" else default
    except (TypeError, ValueError):
        return default


def norm_tf(value):
    s = str(value or "").strip().lower().replace(" ", "").replace("-", "")
    aliases = {"4h": "4H", "4hr": "4H", "4hour": "4H", "daily": "Daily", "day": "Daily",
               "weekly": "Weekly", "week": "Weekly", "monthly": "Monthly", "month": "Monthly",
               "quarterly": "Quarterly", "quarter": "Quarterly",
               "halfyearly": "Half-Yearly", "semiannual": "Half-Yearly"}
    return aliases.get(s)


def parse_dt(value):
    """Return a timezone-aware (IST) datetime from a cell value, or None."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        dt = value
    else:
        try:
            dt = datetime.fromisoformat(str(value).strip())
        except ValueError:
            return None
    return dt.replace(tzinfo=TIMEZONE) if dt.tzinfo is None else dt.astimezone(TIMEZONE)


def naive(dt):
    return dt.replace(tzinfo=None) if dt is not None else None


# ----------------------------------------------------------------------------------------------
# Workbook readers
# ----------------------------------------------------------------------------------------------
def load_workbook():
    path = TEMPLATE_FILE
    if not os.path.exists(path):
        alt = path.replace(" ", "_")
        if os.path.exists(alt):
            path = alt
        else:
            raise FileNotFoundError(f"Workbook not found: {TEMPLATE_FILE}")
    return openpyxl.load_workbook(path), path


def read_config(ws):
    cfg = {}
    for row in range(2, ws.max_row + 1):
        key, value = ws.cell(row, 1).value, ws.cell(row, 2).value
        if key:
            cfg[str(key).strip()] = value
    return cfg


def load_watchlist(ws, suffix):
    items, seen = [], set()
    for row in range(2, ws.max_row + 1):
        symbol, name = ws.cell(row, 1).value, ws.cell(row, 2).value
        if not symbol:
            continue
        raw = str(symbol).strip()
        if not raw or raw in seen:
            continue
        seen.add(raw)
        full = raw + suffix if ("." not in raw and suffix) else raw
        items.append({"raw": raw, "symbol": full, "name": str(name or raw).strip()})
    return items


def load_patterns(ws):
    return [str(ws.cell(r, 1).value).strip() for r in range(2, ws.max_row + 1)
            if ws.cell(r, 1).value and yes(ws.cell(r, 2).value)]


def load_rsi_ranges(ws):
    """Config E2:H7 -> {timeframe: (min, max)} for Active = YES rows."""
    ranges = {}
    for row in range(2, 8):
        tf = norm_tf(ws.cell(row, 5).value)
        if not tf or not yes(ws.cell(row, 6).value):
            continue
        lo, hi = as_float(ws.cell(row, 7).value), as_float(ws.cell(row, 8).value)
        if lo is not None or hi is not None:
            ranges[tf] = (lo, hi)
    return ranges


def load_ema_periods(cfg):
    periods = []
    for token in str(cfg.get("EMA periods", "20,50")).split(","):
        try:
            n = int(token.strip())
            if n > 0:
                periods.append(n)
        except ValueError:
            pass
    return sorted(set(periods)) or [20, 50]


def read_jobs(ws):
    jobs = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not name or r > 21:          # rows 2-21 are the job table; below is the legend
            continue
        jobs.append({
            "row": r, "enabled": ws.cell(r, 1).value, "name": str(name).strip(),
            "timeframe": ws.cell(r, 3).value, "trigger": ws.cell(r, 4).value,
            "run_time": ws.cell(r, 5).value, "weekday": ws.cell(r, 6).value,
            "day": ws.cell(r, 7).value, "rsi": ws.cell(r, 8).value,
            "ema": ws.cell(r, 9).value, "patterns": ws.cell(r, 10).value,
            "divergence": ws.cell(r, 11).value, "multi_tf": ws.cell(r, 12).value,
            "output_rows": ws.cell(r, 13).value,
            "next_run": ws.cell(r, 14).value, "last_started": ws.cell(r, 15).value,
            "last_completed": ws.cell(r, 16).value, "status": ws.cell(r, 17).value,
        })
    return jobs


# ----------------------------------------------------------------------------------------------
# Scheduling (times in IST). Catch-up logic: GitHub cron can be late or skipped, so a job is due
# whenever its latest scheduled slot has passed, has not been started yet, and is not too old.
# ----------------------------------------------------------------------------------------------
def parse_hhmm(value):
    if isinstance(value, dtime):
        return value.hour, value.minute
    if isinstance(value, datetime):
        return value.hour, value.minute
    try:
        hh, mm = [int(x) for x in str(value or "00:00").strip().split(":")[:2]]
        return hh, mm
    except Exception:
        return 0, 0


def slots_on_day(job, d):
    """Scheduled run datetimes (aware, IST) of the job on calendar date d."""
    hh, mm = parse_hhmm(job.get("run_time"))
    trigger = str(job.get("trigger") or "").strip().lower()
    tf = norm_tf(job.get("timeframe"))

    def at(h):
        return datetime(d.year, d.month, d.day, h, mm, tzinfo=TIMEZONE)

    if "every 4 hours" in trigger:
        return [at(h) for h in range(24) if h % 4 == hh % 4] if d.weekday() < 5 else []
    if "daily" in trigger:
        return [at(hh)] if d.weekday() < 5 else []
    if "weekly" in trigger:
        names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        wd = str(job.get("weekday") or "Saturday").strip().capitalize()
        return [at(hh)] if d.weekday() == (names.index(wd) if wd in names else 5) else []
    if "monthly" in trigger:
        if d.day != int(as_float(job.get("day"), 1)):
            return []
        if tf == "Quarterly" and d.month not in (1, 4, 7, 10):
            return []
        if tf == "Half-Yearly" and d.month not in (1, 7):
            return []
        return [at(hh)]
    return []


def latest_slot(job, now):
    for back in range(0, 200):
        d = now.date() - timedelta(days=back)
        past = [s for s in slots_on_day(job, d) if s <= now]
        if past:
            return max(past)
    return None


def next_run_for(job, now):
    for fwd in range(0, 400):
        d = now.date() + timedelta(days=fwd)
        future = [s for s in slots_on_day(job, d) if s > now]
        if future:
            return min(future)
    return None


def max_slot_age(job):
    trigger = str(job.get("trigger") or "").lower()
    if "every 4 hours" in trigger:
        return timedelta(hours=3)
    if "monthly" in trigger:
        return timedelta(hours=72)
    return timedelta(hours=24)


def job_is_due(job, now):
    if not yes(job["enabled"]):
        return False
    slot = latest_slot(job, now)
    if slot is None or now - slot > max_slot_age(job):
        return False
    started = parse_dt(job.get("last_started"))
    if started is not None and started >= slot:
        # already started for this slot; only retry if that run was interrupted
        return str(job.get("status") or "").startswith("⚠️")
    return True


def fmt_next(dt):
    return "⏳ Scheduled for " + dt.strftime("%a %d %b, %H:%M") if dt else "⏳ Scheduled"


# ----------------------------------------------------------------------------------------------
# Data download
# ----------------------------------------------------------------------------------------------
def throttle():
    global NEXT_SLOT
    with THROTTLE_LOCK:
        now = time.time()
        wait = max(0, NEXT_SLOT - now)
        NEXT_SLOT = max(now, NEXT_SLOT) + REQUEST_INTERVAL
    if wait:
        time.sleep(wait)


def fetch_history(symbol, period, interval):
    throttle()
    df = yf.Ticker(symbol).history(period=period, interval=interval, actions=False, auto_adjust=True)
    if df is None or df.empty:
        return None
    if getattr(df.index, "tz", None) is not None:
        df.index = df.index.tz_localize(None)           # keep exchange wall-clock time
    df = df[["Open", "High", "Low", "Close"]].astype(float).dropna()
    return df if not df.empty else None


def fetch_with_retry(symbol, period, interval):
    """Returns a DataFrame or None (no data). Raises the last exception if every attempt errored."""
    last_exc, empties = None, 0
    for attempt in range(MAX_RETRIES):
        try:
            df = fetch_history(symbol, period, interval)
            if df is not None:
                return df
            last_exc = None
            empties += 1
            if empties >= 2:                    # genuinely empty (delisted / bad symbol) - stop early
                return None
        except Exception as exc:
            last_exc = exc
        if attempt < MAX_RETRIES - 1:
            time.sleep(RETRY_BASE_SLEEP * (attempt + 1) + random.random())
    if last_exc:
        raise last_exc
    return None


# ----------------------------------------------------------------------------------------------
# Timeframe construction (only fully closed candles when 'Ignore forming candle' = YES)
# ----------------------------------------------------------------------------------------------
def session_closed(now):
    return now.weekday() >= 5 or now.time() >= MARKET_CLOSE


def period_forming(end_date, now):
    """True if the candle for a period ending on end_date may still receive trades."""
    today = now.date()
    if today > end_date:
        return False
    if today < end_date and np.busday_count(today + timedelta(days=1), end_date + timedelta(days=1)) > 0:
        return True
    return not session_closed(now)


def aggregate(daily, tf):
    idx = daily.index
    if tf == "Weekly":
        keys = idx.to_period("W-FRI")
    elif tf == "Monthly":
        keys = idx.to_period("M")
    elif tf == "Quarterly":
        keys = idx.to_period("Q")
    else:  # Half-Yearly
        keys = pd.Index([f"{d.year}-H{1 if d.month <= 6 else 2}" for d in idx])
    bars = daily.groupby(keys).agg(AGG).dropna()
    ends = []
    for k in bars.index:
        if isinstance(k, pd.Period):
            ends.append(k.end_time.date())
        else:
            y, h = k.split("-H")
            ends.append(date(int(y), 6, 30) if h == "1" else date(int(y), 12, 31))
    return bars, ends


def make_frame(daily, hourly, tf, now, ignore_forming):
    """OHLC DataFrame for a timeframe, or None when data is missing."""
    now_naive = naive(now)
    if tf == "4H":
        if hourly is None or hourly.empty:
            return None
        bars = hourly.resample(pd.Timedelta(hours=4), offset=pd.Timedelta(hours=9, minutes=15)).agg(AGG).dropna()
        if ignore_forming and len(bars):
            start = bars.index[-1]
            end = min(start + pd.Timedelta(hours=4), start.normalize() + pd.Timedelta(hours=15, minutes=30))
            if now_naive < end:
                bars = bars.iloc[:-1]
        return bars
    if daily is None or daily.empty:
        return None
    if tf == "Daily":
        bars = daily
        if ignore_forming and len(bars) and period_forming(bars.index[-1].date(), now):
            bars = bars.iloc[:-1]
        return bars
    bars, ends = aggregate(daily, tf)
    if ignore_forming and len(bars) and period_forming(ends[-1], now):
        bars = bars.iloc[:-1]
    return bars


# ----------------------------------------------------------------------------------------------
# Indicators
# ----------------------------------------------------------------------------------------------
def rsi_last(close, period):
    arr = np.asarray(close, dtype=float)
    if len(arr) < period + 1:
        return None
    v = talib.RSI(arr, timeperiod=period)[-1]
    return round(float(v), 2) if np.isfinite(v) else None


def ema_position(close, ltp, periods):
    arr = np.asarray(close, dtype=float)
    out = {}
    for n in periods:
        if len(arr) < n:
            out[n] = "N/A"
            continue
        v = talib.EMA(arr, timeperiod=n)[-1]
        out[n] = "N/A" if not np.isfinite(v) else ("Above" if ltp > v else "Below")
    return out


def find_patterns(df, enabled, lookback=1):
    found = []
    if df is None or len(df) < 10:
        return found
    op, hi, lo, cl = (np.asarray(df[c], dtype=float) for c in ("Open", "High", "Low", "Close"))
    n = max(1, min(int(lookback or 1), len(df)))
    for name in enabled:
        func = getattr(talib, PATTERN_FUNCS.get(name, ""), None)
        if func is None:
            continue
        try:
            recent = np.asarray(func(op, hi, lo, cl))[-n:]
            if name.startswith("Bullish"):
                hit = np.any(recent > 0)
            elif name.startswith("Bearish"):
                hit = np.any(recent < 0)
            else:
                hit = np.any(recent != 0)
            if hit:
                found.append(name)
        except Exception:
            continue
    return found


def divergence_signal(close, cfg, rsi_period):
    """Returns 'Bullish', 'Bearish' or 'Neutral' using confirmed swing pivots on price vs RSI."""
    s = max(2, int(as_float(cfg.get("Divergence pivot strength"), 3)))
    lookback = int(as_float(cfg.get("Divergence max lookback (bars)"), 60))
    recency = int(as_float(cfg.get("Divergence recency (bars)"), 5))
    hidden = yes(cfg.get("Include hidden divergence", "YES"))
    price = np.asarray(close, dtype=float)
    n = len(price)
    if n < 2 * s + rsi_period + 6:
        return "Neutral"
    rsi = talib.RSI(price, timeperiod=rsi_period)
    lows, highs = [], []
    for i in range(max(s, n - lookback), n - s):
        window = price[i - s:i + s + 1]
        if price[i] == window.min():
            lows.append(i)
        if price[i] == window.max():
            highs.append(i)
    found = []   # (latest pivot index, label)
    if len(lows) >= 2:
        a, b = lows[-2], lows[-1]
        if np.isfinite(rsi[a]) and np.isfinite(rsi[b]) and (n - 1 - (b + s)) <= recency:
            if price[b] < price[a] and rsi[b] > rsi[a]:
                found.append((b, "Bullish"))                       # regular
            elif hidden and price[b] > price[a] and rsi[b] < rsi[a]:
                found.append((b, "Bullish"))                       # hidden
    if len(highs) >= 2:
        a, b = highs[-2], highs[-1]
        if np.isfinite(rsi[a]) and np.isfinite(rsi[b]) and (n - 1 - (b + s)) <= recency:
            if price[b] > price[a] and rsi[b] < rsi[a]:
                found.append((b, "Bearish"))                       # regular
            elif hidden and price[b] < price[a] and rsi[b] > rsi[a]:
                found.append((b, "Bearish"))                       # hidden
    return max(found)[1] if found else "Neutral"


def in_range(val, lo, hi):
    return val is not None and (lo is None or val >= lo) and (hi is None or val <= hi)


def qualifies(rsi_values, ranges, rule):
    if not ranges:
        return True
    checks = [in_range(rsi_values.get(tf), lo, hi) for tf, (lo, hi) in ranges.items()]
    return any(checks) if str(rule).strip().upper() == "ANY" else all(checks)


def job_ranges(job, active_ranges, cfg):
    """RSI ranges that decide selection for this job, plus the rule to combine them."""
    tf = norm_tf(job["timeframe"])
    if yes(job.get("multi_tf")):
        return dict(active_ranges), str(cfg.get("Multi-TF match rule", "ANY") or "ANY")
    if tf in active_ranges:
        return {tf: active_ranges[tf]}, "ALL"
    return {}, "ALL"


# ----------------------------------------------------------------------------------------------
# Per-stock processing
# ----------------------------------------------------------------------------------------------
def process_stock(stock, job, cfg, ranges, rule, enabled_patterns, ema_periods, rsi_tfs, now):
    tf = norm_tf(job["timeframe"])
    rsi_period = int(as_float(cfg.get("RSI period"), 14))
    ignore = yes(cfg.get("Ignore forming candle", "YES"))
    symbol = stock["symbol"]
    try:
        daily = hourly = None
        if any(t != "4H" for t in rsi_tfs):
            period = max((HIST_PERIOD[t] for t in rsi_tfs if t != "4H"), key=lambda p: HIST_RANK[p])
            daily = fetch_with_retry(symbol, period, "1d")
            if daily is None:
                return None, "No data"
        if "4H" in rsi_tfs:
            hourly = fetch_with_retry(symbol, "720d", "1h")
            if hourly is None:
                return None, "No data"

        ltp_src = daily if daily is not None else hourly
        ltp = round(float(ltp_src["Close"].iloc[-1]), 2)
        # Suspended / illiquid / delisted tickers keep their last old candles on Yahoo; an RSI worked out from them
        # is not today's RSI, so such symbols must never be listed as a match. 0 disables the check.
        max_age = as_float(cfg.get("Max data age (days)"), 10)
        if max_age and (naive(now) - ltp_src.index[-1].to_pydatetime()).days > max_age:
            return None, STALE

        frames = {t: make_frame(daily, hourly, t, now, ignore) for t in rsi_tfs}
        main_frame = frames.get(tf)
        if main_frame is None or len(main_frame) < rsi_period + 2:
            return None, INSUFFICIENT

        rsi_values = {t: (rsi_last(frames[t]["Close"], rsi_period) if frames[t] is not None else None)
                      for t in rsi_tfs}
        if not qualifies(rsi_values, ranges, rule):
            return None, EXCLUDED

        close = main_frame["Close"]
        rec = {
            "symbol": stock["raw"], "name": stock["name"], "ltp": ltp, "rsi": rsi_values,
            "ema": ema_position(close, ltp, ema_periods) if yes(job.get("ema")) else {},
            "patterns": find_patterns(main_frame, enabled_patterns,
                                      as_float(cfg.get("Pattern lookback (candles)"), 1))
            if yes(job.get("patterns")) else [],
            "divergence": divergence_signal(close, cfg, rsi_period) if yes(job.get("divergence")) else "Neutral",
        }
        return rec, None
    except Exception as exc:
        return None, f"{type(exc).__name__}: {str(exc)[:120]}"


# ----------------------------------------------------------------------------------------------
# Report workbook
# ----------------------------------------------------------------------------------------------
def build_report(path, job, results, cfg, ema_periods, rsi_tfs, ranges, rule, counts, now, failed_syms):
    tf = norm_tf(job["timeframe"])
    all_rows = str(job.get("output_rows") or "").strip().upper() == "ALL"
    show_ema = all_rows and yes(job.get("ema"))
    show_pat = yes(job.get("patterns"))
    show_div = yes(job.get("divergence"))

    headers = ["Symbol", "Stock Name", "LTP"] + [f"RSI ({t})" for t in rsi_tfs]
    if show_ema:
        headers += [f"EMA {n} ({tf})" for n in ema_periods]
    if show_pat:
        headers.append("Patterns")
    if show_div:
        headers.append("RSI Divergence")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Scanner Report"
    head_fill = PatternFill("solid", fgColor="1F3864")
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    green, red = Font(name="Arial", color="007A33", bold=True), Font(name="Arial", color="C00000", bold=True)
    base = Font(name="Arial")
    ordered = sorted(results, key=lambda r: (r["rsi"].get(tf) is None, r["rsi"].get(tf) or 0))
    for r, rec in enumerate(ordered, 2):
        row = [rec["symbol"], rec["name"], rec["ltp"]] + [rec["rsi"].get(t) for t in rsi_tfs]
        if show_ema:
            row += [rec["ema"].get(n, "N/A") for n in ema_periods]
        if show_pat:
            row.append(", ".join(rec["patterns"]) if rec["patterns"] else "None")
        if show_div:
            row.append(rec["divergence"])
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.font = base
            h = headers[c - 1]
            if h.startswith("EMA"):
                cell.alignment = Alignment(horizontal="center")
                if v == "Above":
                    cell.font = green
                elif v == "Below":
                    cell.font = red
            elif h == "RSI Divergence":
                cell.alignment = Alignment(horizontal="center")
                if v == "Bullish":
                    cell.font = green
                elif v == "Bearish":
                    cell.font = red
            elif h == "LTP" or h.startswith("RSI ("):
                cell.number_format = "#,##0.00"
    widths = {"Symbol": 14, "Stock Name": 38, "LTP": 12, "Patterns": 34, "RSI Divergence": 16}
    for c, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(h, 14)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, len(ordered) + 1)}"

    info = wb.create_sheet("Scan Info")
    rng_txt = "; ".join(f"{t} {lo if lo is not None else '-'}–{hi if hi is not None else '-'}"
                        for t, (lo, hi) in ranges.items()) or "None (all stocks listed)"
    lines = [
        ("Report", cfg.get("Report Name") or "Scanner Reports"),
        ("Job", job["name"]), ("Timeframe", job["timeframe"]),
        ("Output rows", "ALL" if all_rows else "SIGNALS ONLY"),
        ("Generated (IST)", now.strftime("%d %b %Y %H:%M")),
        ("RSI filter", rng_txt), ("Multi-TF rule", rule if len(ranges) > 1 else "n/a"),
        ("Symbols scanned", counts["total"]), ("Matched RSI filter", len(results)),
        (f"Actual {tf} RSI range in this report",
         f"{counts['seen'][0]} – {counts['seen'][1]}" if counts.get("seen") else "n/a"),
        ("Excluded by RSI", counts["excluded"]), ("Failed / no data", counts["failed"]),
        ("  of which stale data (no trades in last "
         f"{int(as_float(cfg.get('Max data age (days)'), 10) or 0)} days)", counts.get("stale", 0)),
        ("Failed symbols (first 50)", ", ".join(failed_syms[:50])),
    ]
    for r, (k, v) in enumerate(lines, 1):
        info.cell(r, 1, k).font = Font(name="Arial", bold=True)
        info.cell(r, 2, v).font = Font(name="Arial")
    info.column_dimensions["A"].width = 26
    info.column_dimensions["B"].width = 90
    wb.save(path)


# ----------------------------------------------------------------------------------------------
# Delivery: Matrix
# ----------------------------------------------------------------------------------------------
def matrix_conf(cfg):
    hs = (os.getenv("MATRIX_HOMESERVER") or str(cfg.get("Matrix homeserver URL") or "")).strip().rstrip("/")
    token = (os.getenv("MATRIX_ACCESS_TOKEN") or str(cfg.get("Matrix bot access token") or "")).strip().strip("\"'")
    room = (os.getenv("MATRIX_ROOM_ID") or str(cfg.get("Matrix room ID") or cfg.get("Matrix room id") or "")).strip()
    user = (os.getenv("MATRIX_USER") or str(cfg.get("Bot Matrix ID") or "")).strip()
    password = os.getenv("MATRIX_PASSWORD") or ""
    if hs and not hs.startswith(("http://", "https://")):
        hs = "https://" + hs
    return hs, token, room, user, password


def _matrix_request(method, url, **kw):
    for attempt in range(4):
        res = requests.request(method, url, timeout=kw.pop("timeout", 30), **kw)
        if res.status_code == 429:
            try:
                wait = res.json().get("retry_after_ms", 2000) / 1000
            except Exception:
                wait = 2
            time.sleep(min(wait + 0.5, 30))
            continue
        return res
    return res


# Token obtained by a password login during this run (None when the configured token was used).
_MATRIX_LOGIN = {"hs": None, "token": None}


def matrix_login(hs, user, password):
    """Password login -> (access_token, message). Gives a fresh token every time, so nothing ever expires."""
    body = {"type": "m.login.password", "identifier": {"type": "m.id.user", "user": user},
            "password": password, "initial_device_display_name": "Stock Scanner (automatic)"}
    res = _matrix_request("POST", f"{hs}/_matrix/client/v3/login", json=body)
    if res.status_code != 200:
        try:
            err = res.json().get("error", res.text[:120])
        except Exception:
            err = res.text[:120]
        return None, f"Password login rejected (HTTP {res.status_code}: {err}). Check MATRIX_USER / MATRIX_PASSWORD."
    tok = res.json().get("access_token")
    if not tok:
        return None, "Password login returned no access token"
    _MATRIX_LOGIN.update(hs=hs, token=tok)
    return tok, "logged in"


def matrix_logout_session():
    """End the temporary session created by matrix_login (never touches a token the user supplied)."""
    hs, tok = _MATRIX_LOGIN["hs"], _MATRIX_LOGIN["token"]
    _MATRIX_LOGIN.update(hs=None, token=None)
    if hs and tok:
        try:
            requests.post(f"{hs}/_matrix/client/v3/logout", headers={"Authorization": f"Bearer {tok}"},
                          json={}, timeout=15)
        except Exception:
            pass


STATE_DIR = os.getenv("MATRIX_STATE_DIR", ".matrix_state")
STATE_FILE = os.path.join(STATE_DIR, "matrix_token.json")


def matrix_state_load(hs):
    """Token pair remembered by an earlier run -> {'access_token', 'refresh_token'} (empty if none / other server)."""
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            d = json.load(f)
        return d if d.get("hs") == hs else {}
    except (OSError, ValueError):
        return {}


def matrix_state_save(hs, access, refresh=""):
    """Remember the token pair for the next run. Writes (and flags 'changed') only when something is new."""
    old = matrix_state_load(hs)
    if old.get("access_token") == access and (old.get("refresh_token") or "") == (refresh or ""):
        return
    try:
        os.makedirs(STATE_DIR, exist_ok=True)
        fd = os.open(STATE_FILE, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump({"hs": hs, "access_token": access, "refresh_token": refresh or ""}, f)
        with open(os.path.join(STATE_DIR, "changed"), "w", encoding="utf-8") as f:
            f.write("1")                                   # tells the workflow to save the cache
    except OSError as exc:
        print(f"[matrix] could not remember token: {type(exc).__name__}")


def matrix_refresh(hs, refresh_token):
    """Exchange a refresh token for a new (access, refresh) pair -> (access, refresh) or (None, message)."""
    res = _matrix_request("POST", f"{hs}/_matrix/client/v3/refresh", json={"refresh_token": refresh_token})
    if res.status_code != 200:
        return None, f"refresh rejected (HTTP {res.status_code})"
    d = res.json()
    return d.get("access_token"), d.get("refresh_token") or refresh_token


def matrix_authenticate(hs, token, user, password, force_login=False, refresh_token=""):
    """Returns (access_token or None, message). Remembers whatever works for the next run."""
    state = matrix_state_load(hs)
    stored_rt = state.get("refresh_token") or ""
    if not force_login:
        for cand in dict.fromkeys(t for t in (_MATRIX_LOGIN["token"], token, state.get("access_token")) if t):
            who = _matrix_request("GET", f"{hs}/_matrix/client/v3/account/whoami",
                                  headers={"Authorization": f"Bearer {cand}"})
            if who.status_code == 200:
                if cand != _MATRIX_LOGIN["token"]:
                    rt = stored_rt if cand == state.get("access_token") else (refresh_token or stored_rt)
                    matrix_state_save(hs, cand, rt)
                return cand, who.json().get("user_id", "")
    for rt in dict.fromkeys(t for t in (stored_rt, refresh_token) if t):     # tokens expired -> renew by itself
        new_access, new_rt = matrix_refresh(hs, rt)
        if new_access:
            matrix_state_save(hs, new_access, new_rt)
            return new_access, "refreshed"
    if user and password:
        if _MATRIX_LOGIN["token"]:                      # an earlier login of this run went stale
            matrix_logout_session()
        return matrix_login(hs, user, password)
    return None, ("Matrix token is dead and cannot renew itself. Create a new access token ONCE (Element > Settings > "
                  "Help & About > Access Token, in a private window, then just close it - do not sign out) and put it "
                  "in the MATRIX_ACCESS_TOKEN secret; the scanner will remember it from then on.")


def matrix_prepare(hs, token, room, user="", password="", force_login=False, refresh_token=""):
    """Authenticate and make sure the bot is in the room. Returns (token or None, message)."""
    tok, info = matrix_authenticate(hs, token, user, password, force_login, refresh_token)
    if not tok:
        return None, info
    quoted = urllib.parse.quote(room, safe="")
    _matrix_request("POST", f"{hs}/_matrix/client/v3/join/{quoted}",
                    headers={"Authorization": f"Bearer {tok}"}, json={})   # no-op if already joined
    return tok, info


def matrix_send_text(hs, token, room, message):
    url = (f"{hs}/_matrix/client/v3/rooms/{urllib.parse.quote(room, safe='')}"
           f"/send/m.room.message/{uuid.uuid4().hex}")
    res = _matrix_request("PUT", url, json={"msgtype": "m.text", "body": message},
                          headers={"Authorization": f"Bearer {token}"})
    return res.status_code == 200, f"HTTP {res.status_code}: {res.text[:200]}"


def matrix_send_file(hs, token, room, path):
    headers = {"Authorization": f"Bearer {token}"}
    name = os.path.basename(path)
    with open(path, "rb") as f:
        up = requests.post(f"{hs}/_matrix/media/v3/upload", params={"filename": name},
                           headers={**headers, "Content-Type": XLSX_MIME}, data=f, timeout=180)
    if up.status_code != 200:
        return False, f"Upload HTTP {up.status_code}: {up.text[:200]}"
    payload = {"msgtype": "m.file", "body": name, "filename": name, "url": up.json()["content_uri"],
               "info": {"mimetype": XLSX_MIME, "size": os.path.getsize(path)}}
    url = (f"{hs}/_matrix/client/v3/rooms/{urllib.parse.quote(room, safe='')}"
           f"/send/m.room.message/{uuid.uuid4().hex}")
    res = _matrix_request("PUT", url, json=payload, headers=headers)
    return res.status_code == 200, f"HTTP {res.status_code}: {res.text[:200]}"


def deliver_matrix(cfg, message, report_path):
    if not yes(cfg.get("Send to Matrix", "YES")):
        return "Disabled in Config"
    hs, token, room, user, password = matrix_conf(cfg)
    rtoken = (os.getenv("MATRIX_REFRESH_TOKEN") or "").strip()
    if not (hs and room and (token or user or matrix_state_load(hs))):
        return "Failed: Matrix not configured (set MATRIX_HOMESERVER, MATRIX_ROOM_ID and MATRIX_ACCESS_TOKEN secrets)"
    try:
        tok, info = matrix_prepare(hs, token, room, user, password, refresh_token=rtoken)
        if not tok:
            return "Failed: " + info
        for attempt in (1, 2):
            problems = []
            if yes(cfg.get("Send summary message", "YES")):
                ok_t, d = matrix_send_text(hs, tok, room, message)
                if not ok_t:
                    problems.append("text " + d)
            ok_f, d = matrix_send_file(hs, tok, room, report_path)
            if not ok_f:
                problems.append("file " + d)
            if problems and attempt == 1 and any("HTTP 401" in p for p in problems):
                tok, info = matrix_prepare(hs, "", room, user, password, force_login=True, refresh_token=rtoken)
                if tok:
                    continue
                return "Failed: " + info
            break
        return "Sent Successfully" if not problems else "Failed: " + " | ".join(problems)
    except Exception as exc:
        return f"Failed: {type(exc).__name__}: {str(exc)[:150]}"


# ----------------------------------------------------------------------------------------------
# Delivery: e-mail (SMTP)
# ----------------------------------------------------------------------------------------------
def deliver_email(cfg, subject, body, report_path):
    if not yes(cfg.get("Send Email", "NO")):
        return "Disabled in Config"
    to_raw = str(cfg.get("Email ID") or os.getenv("EMAIL_TO") or "").strip()
    recipients = [x.strip() for x in re.split(r"[;,\s]+", to_raw) if x.strip()]
    host = (os.getenv("SMTP_HOST") or "").strip()
    user = (os.getenv("SMTP_USER") or "").strip()
    pwd = os.getenv("SMTP_PASSWORD") or ""
    port = int(os.getenv("SMTP_PORT") or "465")
    if not recipients:
        return "Skipped: no address in Config 'Email ID' (or EMAIL_TO secret)"
    if not (host and user and pwd):
        return "Skipped: SMTP_HOST / SMTP_USER / SMTP_PASSWORD secrets missing"
    try:
        msg = EmailMessage()
        msg["Subject"], msg["From"], msg["To"] = subject, (os.getenv("EMAIL_FROM") or user), ", ".join(recipients)
        msg.set_content(body)
        with open(report_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application",
                               subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               filename=os.path.basename(report_path))
        ctx = ssl.create_default_context()
        if port == 465:
            with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
                s.login(user, pwd)
                s.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=60) as s:
                s.starttls(context=ctx)
                s.login(user, pwd)
                s.send_message(msg)
        return "Sent Successfully"
    except Exception as exc:
        return f"Failed: {type(exc).__name__}: {str(exc)[:150]}"


# ----------------------------------------------------------------------------------------------
# Workbook status updates (Schedule / Dashboard / Log)
# ----------------------------------------------------------------------------------------------
def set_cell(ws, r, c, value):
    """Write only when different; returns True if the sheet changed."""
    if ws.cell(r, c).value == value:
        return False
    ws.cell(r, c, value)
    return True


def refresh_dashboard_status(wb):
    if "Dashboard" not in wb.sheetnames:
        return False
    d = wb["Dashboard"]
    changed = set_cell(d, 3, 4, '=COUNTIF(Schedule!Q2:Q21,"*In progress*")')
    changed |= set_cell(d, 3, 8, '=COUNTIF(Schedule!Q2:Q21,"*Scheduled*")')
    try:
        wb.calculation.fullCalcOnLoad = True
    except AttributeError:
        pass
    return changed


def update_job_status(ws, job, status, progress, details, started=None, completed=None, next_run=None):
    r = job["row"]
    ws.cell(r, 17, status)
    ws.cell(r, 18, progress)
    ws.cell(r, 19, details)
    if started:
        ws.cell(r, 15, naive(started))
    if completed:
        ws.cell(r, 16, naive(completed))
    if next_run is not None:
        ws.cell(r, 14, naive(next_run))
    LIVE.job_update(r, status, progress, details, started, completed, next_run)


def set_dashboard_matrix_status(wb, job, text):
    """Dashboard row for Schedule row r is r+5 (rows 7-18 use column G = Matrix Status)."""
    LIVE.set_matrix(job["row"], text)
    r = job["row"] + 5
    if "Dashboard" in wb.sheetnames and 7 <= r <= 18:
        cell = wb["Dashboard"].cell(r, 7)
        if not (isinstance(cell.value, str) and cell.value.startswith("=")):
            cell.value = text[:60]


def tidy_schedule(wb, jobs, cfg, now):
    """Keep Next Run / Status of idle jobs correct. Returns True if anything changed."""
    ws = wb["Schedule"]
    hold = timedelta(minutes=as_float(cfg.get("Completed status duration (min)"), 60))
    changed = False
    for job in jobs:
        r, status = job["row"], str(job.get("status") or "")
        if not yes(job["enabled"]):
            changed |= set_cell(ws, r, 17, "Disabled")
            changed |= set_cell(ws, r, 14, None)
            continue
        nxt = naive(next_run_for(job, now))
        changed |= set_cell(ws, r, 14, nxt)
        if "In progress" in status:      # nothing is running at the start of a fresh run -> was cut off
            changed |= set_cell(ws, r, 17, "⚠️ Interrupted")
            changed |= set_cell(ws, r, 19, "Previous run was interrupted; will retry if still inside its time window")
            job["status"] = "⚠️ Interrupted"
        elif status.startswith("✅"):
            done = parse_dt(job.get("last_completed"))
            if done is None or now - done >= hold:
                changed |= set_cell(ws, r, 17, fmt_next(next_run_for(job, now)))
        elif status in ("", "Disabled") or status.startswith("⏳"):
            changed |= set_cell(ws, r, 17, fmt_next(next_run_for(job, now)))
    return changed


# ----------------------------------------------------------------------------------------------
# Job execution
# ----------------------------------------------------------------------------------------------
def scan(stocks, job, cfg, ranges, rule, patterns, ema_periods, rsi_tfs, now, on_progress):
    results, errors, excluded = [], {}, 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(process_stock, s, job, cfg, ranges, rule, patterns, ema_periods, rsi_tfs, now): s
                   for s in stocks}
        for i, fut in enumerate(as_completed(futures), 1):
            s = futures[fut]
            try:
                rec, err = fut.result()
            except Exception as exc:
                rec, err = None, str(exc)[:120]
            if rec:
                results.append(rec)
                LIVE.match(job["row"], rec["symbol"], rec["ltp"], rec["rsi"].get(norm_tf(job["timeframe"])))
            elif err == EXCLUDED:
                excluded += 1
            else:
                errors[s["raw"]] = (s, err or "Unknown")
            on_progress(i, len(stocks), len(results), failed=len(errors), excluded=excluded, symbol=s["raw"])
    return results, errors, excluded


def run_job(wb, workbook_path, job, ctx, now):
    cfg, watchlist = ctx["cfg"], ctx["watchlist"]
    schedule = wb["Schedule"]
    tf = norm_tf(job["timeframe"])
    ranges, rule = job_ranges(job, ctx["ranges"], cfg)
    rsi_tfs = [t for t in RSI_TFS if t == tf or t in ranges]

    update_job_status(schedule, job, "🔄 In progress", 0, "Scan started", started=now)
    set_dashboard_matrix_status(wb, job, "Running…")
    refresh_dashboard_status(wb)
    wb.save(workbook_path)

    last_save = {"t": time.time()}

    def on_progress(done, total, matched, label="Processed", failed=0, excluded=0, symbol=None):
        LIVE.progress(job["row"], done, total, matched, failed, excluded, symbol)   # real time, every stock
        if done == total or time.time() - last_save["t"] > 60:
            update_job_status(schedule, job, "🔄 In progress", done / max(total, 1),
                              f"{label} {done}/{total}; matches so far {matched}")
            wb.save(workbook_path)
            last_save["t"] = time.time()

    args = (job, cfg, ranges, rule, ctx["patterns"], ctx["ema_periods"], rsi_tfs, now)
    LIVE.event(f"{job['name']}: scanning {len(watchlist)} symbols ({job['timeframe']})")
    results, errors, excluded = scan(watchlist, *args, on_progress)

    for rnd in range(RETRY_ROUNDS):
        retry = [s for s, e in errors.values() if e not in NO_RETRY]
        if not retry:
            break
        print(f"[{job['name']}] retry round {rnd + 1}: {len(retry)} symbols after {RETRY_COOLDOWN}s cooldown")
        LIVE.event(f"{job['name']}: retrying {len(retry)} failed symbols after {RETRY_COOLDOWN}s cooldown", "warn")
        LIVE.set_phase(job["row"], f"Cooling down {RETRY_COOLDOWN}s before retry ({len(retry)} symbols)")
        time.sleep(RETRY_COOLDOWN)
        r2, e2, x2 = scan(retry, *args, lambda d, t, m, **k: LIVE.set_phase(
            job["row"], f"Retry round {rnd + 1}: {d}/{t} symbols"))
        results += r2
        excluded += x2
        errors = {**{k: v for k, v in errors.items() if v[1] in NO_RETRY}, **e2}

    # FINAL RSI GATE: whatever happened above, only stocks whose RSI is inside the Config range(s) of this job may
    # reach the report, the Matrix message and the e-mail. Also removes duplicates.
    seen, strict = set(), []
    for rec in results:
        if rec["symbol"] not in seen and qualifies(rec["rsi"], ranges, rule):
            seen.add(rec["symbol"])
            strict.append(rec)
    if len(strict) != len(results):
        print(f"[{job['name']}] final RSI gate removed {len(results) - len(strict)} out-of-range/duplicate rows")
    results = strict
    failed_syms = sorted(errors)
    counts = {"total": len(watchlist), "excluded": excluded, "failed": len(errors),
              "stale": sum(1 for _, e in errors.values() if e == STALE)}
    seen_vals = [r["rsi"].get(tf) for r in results if r["rsi"].get(tf) is not None]
    counts["seen"] = (min(seen_vals), max(seen_vals)) if seen_vals else None
    LIVE.progress(job["row"], len(watchlist), len(watchlist), len(results), len(errors), excluded)
    LIVE.set_phase(job["row"], "Building report")
    done_at = datetime.now(TIMEZONE)
    ts = done_at.strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^A-Za-z0-9]+", "_", job["name"]).strip("_")
    os.makedirs(REPORT_DIR, exist_ok=True)
    report_path = os.path.join(REPORT_DIR, f"Scan_Results_{ts}_{safe}.xlsx")
    build_report(report_path, job, results, cfg, ctx["ema_periods"], rsi_tfs, ranges, rule, counts, done_at,
                 failed_syms)

    rng_txt = "; ".join(f"{t} {lo if lo is not None else '-'}-{hi if hi is not None else '-'}"
                        for t, (lo, hi) in ranges.items()) or "no RSI filter"
    actual = f", actual {tf} RSI {counts['seen'][0]}-{counts['seen'][1]}" if counts.get("seen") else ""
    summary = (f"{len(watchlist)} scanned, {len(results)} matched RSI filter ({rng_txt}{actual}), "
               f"{counts['failed']} failed; report: {report_path}")
    if failed_syms:
        summary += "; failed sample: " + ", ".join(failed_syms[:10])
        print(f"[{job['name']}] sample errors:")
        for sym in failed_syms[:8]:
            print(f"   {sym}: {errors[sym][1]}")
    if len(watchlist) and len(errors) / len(watchlist) > 0.5:
        summary = "⚠️ MORE THAN HALF OF DOWNLOADS FAILED (Yahoo rate limit?) - " + summary
    print(f"[{job['name']}] report written: {os.path.abspath(report_path)} ({os.path.getsize(report_path)} bytes)")

    # ---- messages --------------------------------------------------------------------------
    top_n = int(as_float(cfg.get("Max signals listed in message"), 15))
    ordered = sorted(results, key=lambda r: (r["rsi"].get(tf) is None, r["rsi"].get(tf) or 0))
    lines = []
    for i, rec in enumerate(ordered[:top_n], 1):
        bits = [f"{i}. {rec['symbol']}  LTP {rec['ltp']}",
                "RSI " + "/".join(f"{t[:1]}{v}" for t, v in rec["rsi"].items() if v is not None)]
        if rec["patterns"]:
            bits.append(", ".join(rec["patterns"]))
        if rec["divergence"] != "Neutral":
            bits.append(f"Div {rec['divergence']}")
        lines.append(" | ".join(bits))
    msg = (f"📊 Stock Scanner – {job['name']} ({job['timeframe']})\n"
           f"Time: {done_at.strftime('%d %b %Y %H:%M IST')}\n"
           f"Scanned: {len(watchlist)} | Matched: {len(results)} | Failed: {counts['failed']}\n"
           f"RSI filter: {rng_txt}\n")
    if lines:
        msg += f"\nTop {len(lines)} (lowest {tf} RSI first):\n" + "\n".join(lines)
    else:
        msg += "\nNo stock matched the RSI filter."

    LIVE.set_phase(job["row"], "Delivering report (Matrix / e-mail)")
    matrix_status = deliver_matrix(cfg, msg, report_path)
    email_status = deliver_email(cfg, f"Stock Scanner – {job['name']} ({job['timeframe']})", msg, report_path)

    full = f"{summary} | Matrix: {matrix_status} | Email: {email_status}"
    log_values = [naive(done_at), job["name"], job["timeframe"], len(watchlist), len(results), full]
    wb["Log"].append(log_values)
    LIVE.log_row(log_values)
    update_job_status(schedule, job, "✅ Completed", 1, full, completed=done_at,
                      next_run=next_run_for(job, done_at))
    set_dashboard_matrix_status(wb, job, matrix_status)
    refresh_dashboard_status(wb)
    wb.save(workbook_path)
    LIVE.event(f"{job['name']}: completed - {len(results)} matched, {counts['failed']} failed", "ok")
    print(f"[{job['name']}] {full}")


# ----------------------------------------------------------------------------------------------
def run():
    global LIVE
    now = datetime.now(TIMEZONE)
    print(f"Stock Scanner started: {now.isoformat()}")
    wb, path = load_workbook()
    cfg = read_config(wb["Config"])
    suffix = str(cfg.get("Ticker suffix", ".NS") or "").strip()
    if suffix and not suffix.startswith(".") and suffix.upper() in ("NS", "BO"):
        suffix = "." + suffix.upper()
    ctx = {
        "cfg": cfg,
        "watchlist": load_watchlist(wb["Watchlist"], suffix),
        "patterns": load_patterns(wb["Patterns"]),
        "ranges": load_rsi_ranges(wb["Config"]),
        "ema_periods": load_ema_periods(cfg),
    }
    jobs = read_jobs(wb["Schedule"])

    changed = tidy_schedule(wb, jobs, cfg, now)
    changed |= refresh_dashboard_status(wb)

    LIVE = live_status.create(TIMEZONE)
    try:
        LIVE.start(live_status.snapshot_from_workbook(wb, TIMEZONE), os.getenv("GITHUB_EVENT_NAME") or "local")
    except Exception as exc:                      # the live feed must never stop a scan
        print(f"[live] disabled: {type(exc).__name__}: {exc}")
        LIVE = live_status.NullLive()

    manual = os.getenv("GITHUB_EVENT_NAME", "") == "workflow_dispatch"
    run_all = manual or os.getenv("RUN_ALL_ENABLED", "").strip().lower() in ("1", "yes", "true")
    print(f"Trigger: {os.getenv('GITHUB_EVENT_NAME') or 'local'} | run all enabled jobs now: {run_all}")
    for j in jobs:
        if yes(j["enabled"]):
            slot, nxt = latest_slot(j, now), next_run_for(j, now)
            print(f"  job '{j['name']}': latest slot {slot:%a %d %b %H:%M} | next slot "
                  f"{nxt:%a %d %b %H:%M} | due now: {job_is_due(j, now)}" if slot and nxt else
                  f"  job '{j['name']}': no valid schedule - check Trigger/Run Time/Weekday/Day columns")
    only = os.getenv("JOB_NAME", "").strip().lower()
    due = [j for j in jobs if (yes(j["enabled"]) if run_all else job_is_due(j, now))]
    if only:
        due = [j for j in due if j["name"].lower() == only]

    if not due:
        print("No enabled schedule job is due in this run. (A manual 'Run workflow' always scans enabled jobs; "
              "scheduled runs wait for the job's slot in the Schedule sheet, IST.)")
        LIVE.event("Nothing due in this run")
        LIVE.idle_run = True
        if changed:
            wb.save(path)
            print("Schedule statuses refreshed.")
        return

    print(f"Watchlist: {len(ctx['watchlist'])} | RSI ranges: {ctx['ranges']} | Due jobs: {[j['name'] for j in due]}")
    for job in due:
        try:
            run_job(wb, path, job, ctx, now)
        except Exception as exc:
            failed_at = datetime.now(TIMEZONE)
            update_job_status(wb["Schedule"], job, "❌ Failed", 0, f"{type(exc).__name__}: {str(exc)[:400]}",
                              completed=failed_at)
            set_dashboard_matrix_status(wb, job, "Report not prepared")
            refresh_dashboard_status(wb)
            wb.save(path)
            LIVE.event(f"{job['name']}: FAILED - {type(exc).__name__}: {str(exc)[:200]}", "error")
            print(f"[{job['name']}] FAILED: {exc}")
    print("Scanner run finished.")


def main():
    try:
        run()
    finally:
        # an idle poll (nothing due) only refreshes the local file; real runs publish their final state
        LIVE.close(final_publish=not getattr(LIVE, "idle_run", False))
        matrix_logout_session()          # drop the temporary Matrix session created by the password login


if __name__ == "__main__":
    main()
