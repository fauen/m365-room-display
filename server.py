#!/usr/bin/env python3
"""
Room free/busy board – Python 3.8+, zero external dependencies.

Usage:
    python3 server.py
Then open http://localhost:8888 in a browser.

Azure App Registration needs:
    Application permission: Calendars.ReadBasic.All  (admin consent required)
"""

import json
import os
import threading
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

# On Windows the timezone database is not bundled with Python.
# If you get a ZoneInfoNotFoundError, run: pip install tzdata
try:
    TZ = ZoneInfo("Europe/Stockholm")
except ZoneInfoNotFoundError:
    raise SystemExit(
        "Timezone data not found. Please run: pip install tzdata"
    )

# ── Configuration ─────────────────────────────────────────────────────────────

# Required environment variables – set before running:
#   export AZURE_TENANT_ID="your-tenant-id"
#   export AZURE_CLIENT_ID="your-client-id"
#   export AZURE_CLIENT_SECRET="your-client-secret"
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

PORT = 8888

_config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
with open(_config_path, encoding="utf-8-sig") as _f:
    _config = json.load(_f)
GROUPS = _config["groups"]
THEME = _config.get("theme", {})

ROOMS = [r for g in GROUPS for r in g["rooms"]]

# ──────────────────────────────────────────────────────────────────────────────

_token = {"value": None, "expires": 0.0}
_token_lock = threading.Lock()


def get_token() -> str:
    with _token_lock:
        if time.time() < _token.get("expires", 0) - 60:
            return _token["value"]

        url  = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
        body = urllib.parse.urlencode({
            "grant_type":    "client_credentials",
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope":         "https://graph.microsoft.com/.default",
        }).encode()

        req = urllib.request.Request(url, data=body, method="POST")
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read())

        _token["value"]   = data["access_token"]
        _token["expires"] = time.time() + data["expires_in"]
        return _token["value"]


def get_schedule(token: str, room_emails: list) -> list:
    """
    Uses getSchedule (works with Calendars.ReadBasic.All).
    Returns free/busy blocks without meeting titles.
    Calls the endpoint on the first room in the list as the calling user.
    """
    now_local = datetime.now(TZ)
    start = now_local.replace(hour=0,  minute=0,  second=0).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    end   = now_local.replace(hour=23, minute=59, second=59).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")

    caller = urllib.parse.quote(room_emails[0])
    url    = f"https://graph.microsoft.com/v1.0/users/{caller}/calendar/getSchedule"

    body = json.dumps({
        "schedules":                room_emails,
        "startTime":                {"dateTime": start, "timeZone": "UTC"},
        "endTime":                  {"dateTime": end,   "timeZone": "UTC"},
        "availabilityViewInterval": 15,
    }).encode()

    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    }, method="POST")

    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read()).get("value", [])


def parse_dt(s: str) -> datetime:
    """Parse an ISO datetime string (UTC) into an aware datetime."""
    s = s.rstrip("Z").split(".")[0]
    return datetime.fromisoformat(s).replace(tzinfo=timezone.utc)


def room_status(schedule_items: list):
    """
    Returns (status, relevant_time) where:
      status        – "busy" | "free"
      relevant_time – end time if busy, next start time if free (or None)
    """
    now = datetime.now(timezone.utc)

    for item in schedule_items:
        start = parse_dt(item["start"]["dateTime"])
        end   = parse_dt(item["end"]["dateTime"])
        if start <= now <= end:
            return "busy", end

    for item in schedule_items:
        start = parse_dt(item["start"]["dateTime"])
        if start > now:
            return "free", start

    return "free", None


def fmt_time(dt) -> str:
    if dt is None:
        return ""
    return dt.astimezone(TZ).strftime("%H:%M")


def build_page() -> str:
    token = get_token()

    # Fetch all rooms in a single API call
    room_emails = [r["email"] for r in ROOMS]
    try:
        schedule_data = {s["scheduleId"]: s for s in get_schedule(token, room_emails)}
    except Exception as exc:
        schedule_data = {r: {"error": str(exc)} for r in room_emails}

    def render_card(room):
        data = schedule_data.get(room["email"], {})
        if "error" in data:
            css, badge, detail = "error", "Error", data["error"]
        else:
            try:
                status, rel_time = room_status(data.get("scheduleItems", []))
            except Exception as exc:
                return f'<div class="card error"><div class="card-top"><span class="room-name">{room["name"]}</span><span class="badge">Error</span></div><p class="detail">{exc}</p></div>'
            if status == "busy":
                css, badge, detail = "busy", "Busy", f"Until {fmt_time(rel_time)}"
            else:
                css, badge = "free", "Free"
                detail = f"Next booking at {fmt_time(rel_time)}" if rel_time else "Free all day"
        return f"""
        <div class="card {css}">
          <div class="card-top">
            <span class="room-name">{room['name']}</span>
            <span class="badge">{badge}</span>
          </div>
          <p class="detail">{detail}</p>
        </div>"""

    sections = ""
    for group in GROUPS:
        cards = "".join(render_card(r) for r in group["rooms"])
        sections += f"""
    <section>
      <h2 class="group-heading">{group['name']}</h2>
      <div class="grid">{cards}
      </div>
    </section>"""

    now      = datetime.now(TZ)
    time_str = now.strftime("%H:%M")
    date_str = now.strftime("%d.%m.%Y")

    t = THEME
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="refresh" content="60">
  <title>Room availability</title>
  <style>
    :root {{
      --font-family:      {t.get("font_family", "system-ui, -apple-system, sans-serif")};
      --color-bg:         {t.get("color_background", "#0b1f1c")};
      --color-surface:    {t.get("color_surface", "#112b26")};
      --color-text:       {t.get("color_text", "#ffffff")};
      --color-text-2:     {t.get("color_text_secondary", "#6b9e94")};
      --color-free-bg:    {t.get("color_free_bg", "#1a4a3a")};
      --color-free-fg:    {t.get("color_free_fg", "#3ddc84")};
      --color-busy-bg:    {t.get("color_busy_bg", "#4a1a1a")};
      --color-busy-fg:    {t.get("color_busy_fg", "#ff6b6b")};
      --color-error-bg:   {t.get("color_error_bg", "#2a2a2a")};
      --color-error-fg:   {t.get("color_error_fg", "#888888")};
    }}

    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    body {{
      font-family: var(--font-family);
      background: var(--color-bg);
      color: var(--color-text);
      min-height: 100vh;
      padding: 2rem 2.5rem;
    }}

    header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 2rem;
    }}

    .branding {{ font-size: 0.85rem; color: var(--color-text-2); letter-spacing: .05em; }}

    h1 {{
      font-size: 2.4rem;
      font-weight: 700;
      margin-top: 0.25rem;
      color: var(--color-text);
    }}

    .clock {{
      text-align: right;
    }}
    .clock .time {{
      font-size: 2.4rem;
      font-weight: 700;
      color: var(--color-text);
    }}
    .clock .date {{
      font-size: 0.85rem;
      color: var(--color-text-2);
      margin-top: 0.15rem;
    }}

    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
      gap: 0.6rem;
    }}

    .card {{
      background: var(--color-surface);
      border-radius: 8px;
      padding: 0.75rem 1rem;
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
    }}

    .card-top {{
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}

    .room-name {{
      font-size: 0.9rem;
      font-weight: 600;
      color: var(--color-text);
    }}

    .badge {{
      font-size: 0.7rem;
      font-weight: 600;
      padding: 0.15rem 0.5rem;
      border-radius: 999px;
      white-space: nowrap;
    }}
    .free  .badge {{ background: var(--color-free-bg);  color: var(--color-free-fg);  }}
    .busy  .badge {{ background: var(--color-busy-bg);  color: var(--color-busy-fg);  }}
    .error .badge {{ background: var(--color-error-bg); color: var(--color-error-fg); }}

    .detail {{
      font-size: 0.8rem;
      color: var(--color-text-2);
    }}

    section {{ margin-bottom: 2.5rem; }}

    .group-heading {{
      font-size: 1rem;
      font-weight: 600;
      color: var(--color-text-2);
      text-transform: uppercase;
      letter-spacing: .1em;
      margin-bottom: 0.75rem;
    }}
  </style>
</head>
<body>
  <header>
    <div>
      <div class="branding">bookbeat</div>
      <h1>Meeting rooms</h1>
    </div>
    <!-- <div class="clock">
      <div class="time">{time_str}</div>
      <div class="date">{date_str}</div>
    </div> -->
  </header>
  {sections}
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):          # silence access log
        pass

    def do_GET(self):
        if self.path not in ("/", "/index.html"):
            self.send_response(404)
            self.end_headers()
            return

        try:
            body = build_page().encode()
            self.send_response(200)
            self.send_header("Content-Type",   "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control",  "no-store")
            self.end_headers()
            self.wfile.write(body)
        except Exception as exc:
            error = f"<pre>{exc}</pre>".encode()
            self.send_response(500)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(error)


if __name__ == "__main__":
    print(f"Starting server on http://localhost:{PORT}")
    print("Press Ctrl+C to stop.")
    HTTPServer(("", PORT), Handler).serve_forever()
