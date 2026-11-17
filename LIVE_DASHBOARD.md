# Live Dashboard (Bloomberg-style, real-time)

Shows the contents of `MASTER DASHBOARD.xlsx` live, with a per-job progress bar that moves on **every stock**
(the workbook itself is only saved about once a minute and at the end of a job).

## Files
| File | Purpose |
|---|---|
| `live_status.py` | Real-time state + publisher (called by `Scanner.py`; no-op unless `LIVE_DASHBOARD=1`) |
| `dashboard_server.py` | Serves the UI + `/api/state` + `/events` (SSE). Stdlib only |
| `dashboard/index.html` | The terminal-style UI |
| `.github/workflows/scanner.yml` | Updated: `live_dashboard` checkbox, publishes progress to branch `live-status` |
| `.github/workflows/live-dashboard.yml` | NEW: starts the dashboard on a runner and prints a public link |

## A) Watch a scan running on GitHub Actions
1. **Actions -> Scheduled Stock Scanner -> Run workflow**, keep **live dashboard** ticked.
   (For scheduled runs, add a repository *variable* `LIVE_DASHBOARD` = `1`.)
2. **Actions -> Live Dashboard -> Run workflow** (choose minutes). Open the link in the run summary.
   The scanner creates the `live-status` branch by itself on its first publish.

Alternative without the second workflow: open `dashboard/index.html` from disk (or any static host),
click the gear, enter `owner/repo` (+ a read-only fine-grained token for private repos).

## B) Watch a scan on your own machine
```
LIVE_DASHBOARD=1 python Scanner.py              # terminal 1 (writes live/live_status.json)
python dashboard_server.py --open               # terminal 2 -> http://127.0.0.1:8080
```
With no scan running the dashboard shows the saved workbook (idle).

## Notes
* Secrets never enter the feed (config keys with token/password/secret/matrix/email/smtp are dropped).
* Public repo => the `live-status` branch and workflow logs are public. Use a private repo if that matters.
* The feed can never stop a scan: every error inside it is caught and printed as `[live] ...`.
