"""
PheedLoop Announcement Scheduler
================================

Bulk-creates *scheduled* announcements in PheedLoop from a spreadsheet, since the
whitelabel app builder has no bulk-upload for announcements.

Spreadsheet (sheet "All Notifications", header row auto-detected):

    required:  notify_time | title | description | status
    date        (or weekday, as a fallback)
    optional:   event_time | location | tag | channel | cta_text | cta_link |
                add_to_calendar | ebadge

("event" is accepted as a legacy alias for `title`; "event_id" is accepted too,
for older sheets, but is no longer required — see Field mapping below.)

Field mapping — every field the PheedLoop announcement form has, and how it's covered
---------------------------------------------------------------------------------------
    title  (legacy alias: event)   -> title            (shown to attendees as the announcement
                                                          headline — exactly what they'll see)
    description (+ event_time +    -> description   (API has no separate time/location fields;
                location)                            location is skipped if blank or "TBD")
    date  (fallback: weekday)      -> Display Date & Automatic Publish Date (same value; the
    notify_time                       API calls these publish_date/release_date and
                                       publish_time/release_time)
    status                         -> only rows with status in UPLOADABLE_STATUSES are sent
    channel   (optional; blank/     -> Email/Push Notifications Enabled (is_notification_mail /
              push/email/both/none)   is_notification_push). Blank = DEFAULT_SETTINGS below.
    cta_text, cta_link (optional)  -> Call-To-Action Text / URL
    add_to_calendar (optional)     -> Include Add-To-Calendar Links (is_add_to_calendar)
    ebadge (optional)              -> Include eBadge (is_ebadge_included)
    tag  (optional)                -> Target Settings tags (prepared only — see caveat below)
    (auto)                         -> dedup key (derived from date+notify_time+title unless an
                                       explicit `event_id` column is present; recorded in a local
                                       ledger, so re-running is safe either way)
    DEFAULT_SETTINGS (script-wide) -> Allow Multiple Deliveries, all 6 Visibility toggles,
                                       Include Login Credentials — same for every row unless
                                       you edit the constant below; not exposed as columns
                                       because they don't vary row to row for this event.

Not automatable at all — the API's create/read schema simply has no fields for these
(confirmed against a live GET), so they aren't in the spreadsheet either. Set them by hand
in the dashboard after creation: Event App Settings > Linked Page Type, Hide Default Banner
Image, Attachment.

Only untargeted (all-attendee) announcements are automatable: the public API silently
drops the `tags`/`tickets` targeting field on write (verified), so targeting must be done
in the dashboard UI. All-attendee reminders like these work fully via the API.

Run modes
---------
    (default)        dry run   -> build + print payloads, no API calls
    --run            create scheduled announcements (skips ones already in the ledger)
    --limit N        cap how many to create/preview this run (e.g. --limit 10)
    --test           create one SILENT throwaway ~5 min out to confirm scheduling
    --status CODE    print an announcement's status/scheduling fields
    --delete CODE    delete an announcement by code

Credentials live in a local .env file (see .env.example).
"""

import argparse
import csv
import hashlib
import json
import os
import sys
from datetime import datetime, date, time, timedelta

import requests
from dotenv import load_dotenv

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

load_dotenv()

ORGANIZATION_CODE = os.getenv("PHEEDLOOP_ORG_CODE", "")
EVENT_CODE = os.getenv("PHEEDLOOP_EVENT_CODE", "")
API_KEY = os.getenv("PHEEDLOOP_API_KEY", "")
API_SECRET = os.getenv("PHEEDLOOP_API_SECRET", "")

BASE_URL = "https://api.pheedloop.com/"

SPREADSHEET_PATH = os.getenv("SPREADSHEET_PATH", "notifications.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "All Notifications")  # xlsx: preferred sheet

# Optional per-row targeting: a spreadsheet column holding tag name(s), comma-separated
# (e.g. "Marshal Exec"). Resolved to numeric tag IDs via the /tags/ endpoint.
# NOTE: the public API silently drops `tags` on write, so this "prepares" the target but
# you must confirm/apply it in the dashboard UI. A `--tag` CLI flag overrides per run.
TAG_COLUMN = os.getenv("TAG_COLUMN", "tag")

# Fallback only, if a row has no `date`: the Sunday orientation week starts on.
WEEK_START_SUNDAY = os.getenv("WEEK_START_SUNDAY", "2026-08-30")
WEEKDAY_OFFSETS = {"sunday": 0, "monday": 1, "tuesday": 2, "wednesday": 3,
                   "thursday": 4, "friday": 5, "saturday": 6}

# Only rows with one of these statuses get uploaded.
UPLOADABLE_STATUSES = {"draft"}

# Script-wide defaults for every announcement field that doesn't get its own spreadsheet
# column, because it's the same for every row in this event. Override a field for one row
# with the matching optional column (see the docstring's Field mapping) instead of editing
# this per-run — only fields listed here as False/"" have an optional column at all.
DEFAULT_SETTINGS = {
    "is_notification_push": True,       # column override: channel
    "is_notification_mail": False,      # column override: channel
    "is_multiple_send_enabled": False,
    "is_visible_app": True,
    "is_visible_attendee": True,
    "is_visible_virtual": False,
    "is_visible_exhibitor": False,
    "is_visible_speaker": False,
    "is_visible_sponsor": False,
    "cta_text": "",                     # column override: cta_text
    "cta_link": "",                     # column override: cta_link
    "is_credentials_included": False,
    "is_add_to_calendar": False,        # column override: add_to_calendar
    "is_ebadge_included": False,        # column override: ebadge
}

# channel column values -> (is_notification_push, is_notification_mail). Blank uses DEFAULT_SETTINGS.
CHANNEL_OVERRIDES = {
    "push": (True, False), "email": (False, True),
    "both": (True, True), "none": (False, False),
}

TRUE_VALUES = {"yes", "y", "true", "1", "x"}
FALSE_VALUES = {"no", "n", "false", "0"}

DESCRIPTION_SOFT_LIMIT = 115
SKIP_LOCATION_VALUES = {"", "tbd", "n/a", "na", "-"}

# Columns we need present (plus one of date/weekday, and one of title/event).
REQUIRED_COLUMNS = ["notify_time"]

LEDGER_PATH = "created_announcements.json"   # event_id -> {code, title, scheduled_for}
RESULTS_LOG = "announcement_results.csv"


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def _headers(json_body: bool = True) -> dict:
    h = {"X-API-KEY": API_KEY, "X-API-SECRET": API_SECRET}
    if json_body:
        h["Content-Type"] = "application/json"
    return h


def _announcements_url(code: str = "") -> str:
    url = (f"{BASE_URL}api/v3/organization/{ORGANIZATION_CODE}"
           f"/events/{EVENT_CODE}/announcements/")
    return url + f"{code}/" if code else url


def check_credentials() -> None:
    missing = [n for n, v in [
        ("PHEEDLOOP_ORG_CODE", ORGANIZATION_CODE),
        ("PHEEDLOOP_EVENT_CODE", EVENT_CODE),
        ("PHEEDLOOP_API_KEY", API_KEY),
        ("PHEEDLOOP_API_SECRET", API_SECRET),
    ] if not v]
    if missing:
        sys.exit("Missing credentials in .env: " + ", ".join(missing)
                 + "\nCopy .env.example to .env and fill in the values.")


def clean(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date, time)):
        return str(value)
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none") else text


def row_title(row: dict) -> str:
    """The announcement title. `title` is the current column name; `event` is
    accepted too for older sheets that still use it."""
    return clean(row.get("title")) or clean(row.get("event"))


def row_key(row: dict) -> str:
    """Stable dedup key for the ledger. An explicit `event_id` column (older
    sheets) wins; otherwise one is derived from the row's own identifying
    fields, so a fresh template with no event_id column still dedupes safely
    across re-runs without anyone having to invent an id by hand."""
    explicit = clean(row.get("event_id"))
    if explicit:
        return explicit
    basis = "|".join([clean(row.get("date")), clean(row.get("weekday")),
                      clean(row.get("notify_time")), row_title(row)])
    return hashlib.sha1(basis.encode("utf-8")).hexdigest()[:10]


def weekday_to_date(weekday: str) -> str:
    key = weekday.strip().lower()
    if key not in WEEKDAY_OFFSETS:
        raise ValueError(f"Unrecognized weekday {weekday!r}.")
    start = datetime.strptime(WEEK_START_SUNDAY, "%Y-%m-%d")
    return (start + timedelta(days=WEEKDAY_OFFSETS[key])).strftime("%Y-%m-%d")


def row_date(row: dict) -> str:
    """Prefer the row's `date`; fall back to weekday within the orientation week."""
    raw = row.get("date")
    if isinstance(raw, (datetime, date)):
        return raw.strftime("%Y-%m-%d")
    text = clean(raw)
    if text:
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        try:
            return datetime.fromisoformat(text).strftime("%Y-%m-%d")
        except ValueError:
            pass
    wd = clean(row.get("weekday"))
    if wd:
        return weekday_to_date(wd)
    raise ValueError("Row has no usable date or weekday.")


def parse_time(raw) -> str:
    if isinstance(raw, (datetime, time)):
        return raw.strftime("%H:%M:%S")
    text = clean(raw)
    if not text:
        raise ValueError("Empty notify_time.")
    normalized = text.upper().replace(".", "").replace("  ", " ")
    for fmt in ("%I:%M %p", "%I:%M%p", "%I %p", "%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(normalized, fmt).strftime("%H:%M:%S")
        except ValueError:
            continue
    raise ValueError(f"Could not parse time {raw!r} (try '9:45 AM' or '15:30').")


def parse_bool_override(raw, default: bool, field: str) -> bool:
    """Blank cell -> default. Otherwise a yes/no-ish value overrides it."""
    text = clean(raw).lower()
    if not text:
        return default
    if text in TRUE_VALUES:
        return True
    if text in FALSE_VALUES:
        return False
    raise ValueError(f"Unrecognized {field} value {raw!r} (use yes/no).")


def compose_description(row: dict) -> str:
    body = clean(row.get("description"))
    event_time = clean(row.get("event_time"))
    location = clean(row.get("location"))

    bits = []
    if event_time:
        bits.append(f" {event_time}")
    if location.lower() not in SKIP_LOCATION_VALUES:
        bits.append(f" {location}")
    if bits:
        line = " · ".join(bits)
        body = f"{body}\n{line}" if body else line
    return body


def build_payload(row: dict, default_tags: list = None) -> dict:
    title = row_title(row)
    if not title:
        raise ValueError("Row is missing a 'title'.")
    date_str = row_date(row)
    time_str = parse_time(row.get("notify_time"))

    # Tag targeting: per-row `tag` column wins; else the run-wide --tag default.
    tag_names = split_tags(row.get(TAG_COLUMN)) or (default_tags or [])
    tag_ids, unresolved = resolve_tag_names(tag_names)

    payload = dict(DEFAULT_SETTINGS)
    payload.update({
        "event": EVENT_CODE,
        "title": title,
        "description": compose_description(row),
        "is_released": False,  # scheduled; PheedLoop auto-publishes at the time
        "publish_date": date_str, "publish_time": time_str, "publis_date": date_str,
        "release_date": date_str, "release_time": time_str,
        "tags": tag_ids, "tickets": [],
    })

    # Optional per-row overrides of the script-wide defaults above.
    channel = clean(row.get("channel")).lower()
    if channel:
        if channel not in CHANNEL_OVERRIDES:
            raise ValueError(f"Unrecognized channel {channel!r} "
                             f"(use push/email/both/none).")
        payload["is_notification_push"], payload["is_notification_mail"] = CHANNEL_OVERRIDES[channel]
    if clean(row.get("cta_text")):
        payload["cta_text"] = clean(row.get("cta_text"))
    if clean(row.get("cta_link")):
        payload["cta_link"] = clean(row.get("cta_link"))
    payload["is_add_to_calendar"] = parse_bool_override(
        row.get("add_to_calendar"), DEFAULT_SETTINGS["is_add_to_calendar"], "add_to_calendar")
    payload["is_ebadge_included"] = parse_bool_override(
        row.get("ebadge"), DEFAULT_SETTINGS["is_ebadge_included"], "ebadge")

    payload["_scheduled_for"] = f"{date_str} {time_str}"
    payload["_event_id"] = row_key(row)
    payload["_tag_names"] = tag_names
    payload["_tag_ids"] = tag_ids
    payload["_unresolved_tags"] = unresolved
    return payload


def _post_body(payload: dict) -> dict:
    return {k: v for k, v in payload.items() if not k.startswith("_")}


# --------------------------------------------------------------------------- #
# API calls
# --------------------------------------------------------------------------- #

def create_announcement(payload: dict) -> dict:
    r = requests.post(_announcements_url(), headers=_headers(),
                      json=_post_body(payload), timeout=30)
    if not (200 <= r.status_code < 300):
        raise RuntimeError(f"Create failed (HTTP {r.status_code}): {r.text}")
    return r.json()


def get_announcement(code: str) -> dict:
    r = requests.get(_announcements_url(code), headers=_headers(False), timeout=30)
    if not (200 <= r.status_code < 300):
        raise RuntimeError(f"Get failed (HTTP {r.status_code}): {r.text}")
    return r.json()


def print_status(code: str) -> None:
    d = get_announcement(code)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] Announcement {code}:")
    for f in ("title", "status", "deliveries", "is_released",
              "publish_date", "publish_time", "release_date", "release_time"):
        print(f"    {f:14} = {d.get(f)!r}")


def delete_announcement(code: str) -> None:
    r = requests.delete(_announcements_url(code), headers=_headers(False), timeout=30)
    if not (200 <= r.status_code < 300):
        raise RuntimeError(f"Delete failed (HTTP {r.status_code}): {r.text}")
    print(f"Deleted announcement {code}.")


def do_delete_all() -> None:
    """Delete every announcement this script has created (per the ledger) and
    clear those entries, so a future --run recreates them fresh from whatever
    is currently in the spreadsheet. Use this instead of hand-editing an
    already-created row: editing title/date/notify_time changes its dedup key,
    so --run would create a duplicate rather than updating the old one."""
    ledger = load_ledger()
    if not ledger:
        print("Ledger is empty — nothing to delete.")
        return
    print(f"This will DELETE {len(ledger)} announcement(s) from PheedLoop (everything this "
          f"script has created) and clear them from {LEDGER_PATH}, so the next --run "
          f"recreates them fresh from the current spreadsheet:\n")
    for rec in ledger.values():
        print(f"  {rec.get('code')}  {rec.get('title')!r}  ({rec.get('scheduled_for')})")
    confirm = input(f"\nType DELETE to confirm deleting all {len(ledger)} announcement(s): ")
    if confirm.strip() != "DELETE":
        print("Aborted — nothing deleted.")
        return
    for key, rec in list(ledger.items()):
        code = rec.get("code")
        try:
            delete_announcement(code)
            del ledger[key]
        except RuntimeError as e:
            print(f"  ⚠️  failed to delete {code}: {e}")
    save_ledger(ledger)
    remaining = len(ledger)
    print(f"\nDone. {remaining} announcement(s) failed to delete and remain in the ledger."
          if remaining else "\nDone. All deleted and the ledger is now empty.")


# --------------------------------------------------------------------------- #
# Tag targeting (name -> numeric id). Read-only via API; UI applies it for real.
# --------------------------------------------------------------------------- #

def _tags_url() -> str:
    return (f"{BASE_URL}api/v3/organization/{ORGANIZATION_CODE}"
            f"/events/{EVENT_CODE}/tags/")


_TAG_MAP_CACHE = None


def get_tag_map() -> dict:
    """Return {tag_name_lower: numeric_id}, fetched once and cached."""
    global _TAG_MAP_CACHE
    if _TAG_MAP_CACHE is None:
        r = requests.get(_tags_url() + "?page_size=500", headers=_headers(False), timeout=30)
        if not (200 <= r.status_code < 300):
            raise RuntimeError(f"Tags fetch failed (HTTP {r.status_code}): {r.text}")
        data = r.json()
        results = data.get("results", data if isinstance(data, list) else [])
        _TAG_MAP_CACHE = {clean(t.get("name")).lower(): t.get("id") for t in results}
    return _TAG_MAP_CACHE


def split_tags(raw) -> list:
    """Split a tag cell / CLI value ('Marshal Exec, Marshal') into a list of names."""
    text = clean(raw)
    return [t.strip() for t in text.split(",") if t.strip()] if text else []


def resolve_tag_names(names: list) -> tuple:
    """Map tag names to numeric ids. Returns (ids, unresolved_names)."""
    if not names:
        return [], []
    tag_map = get_tag_map()
    ids, unresolved = [], []
    for name in names:
        tid = tag_map.get(name.strip().lower())
        (ids.append(tid) if tid is not None else unresolved.append(name.strip()))
    return ids, unresolved


# --------------------------------------------------------------------------- #
# Spreadsheet loading (header row auto-detected; leading title rows skipped)
# --------------------------------------------------------------------------- #

def _looks_like_header(cells) -> bool:
    lowered = {clean(c).lower() for c in cells}
    return ("title" in lowered or "event" in lowered) and "notify_time" in lowered


def _rows_from_matrix(matrix) -> list:
    header_idx = next((i for i, r in enumerate(matrix) if _looks_like_header(r)), None)
    if header_idx is None:
        sys.exit("Could not find a header row containing 'event' and 'notify_time'.")
    headers = [clean(h).lower() for h in matrix[header_idx]]
    rows = []
    for values in matrix[header_idx + 1:]:
        if all(clean(v) == "" for v in values):
            continue
        rows.append({h: v for h, v in zip(headers, values) if h})
    missing = [c for c in REQUIRED_COLUMNS if c not in headers]
    if "title" not in headers and "event" not in headers:
        missing.append("title")
    if missing:
        sys.exit(f"Spreadsheet missing required column(s): {', '.join(missing)}.\n"
                 f"Found: {', '.join(headers)}")
    return rows


def _load_xlsx(path: str) -> list:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    matrix = [list(r) for r in ws.iter_rows(values_only=True)]
    return _rows_from_matrix(matrix)


def _load_csv(path: str) -> list:
    with open(path, newline="", encoding="utf-8-sig") as f:
        matrix = [row for row in csv.reader(f)]
    return _rows_from_matrix(matrix)


def load_rows() -> list:
    if not os.path.exists(SPREADSHEET_PATH):
        sys.exit(f"Spreadsheet not found: {SPREADSHEET_PATH}")
    ext = os.path.splitext(SPREADSHEET_PATH)[1].lower()
    if ext == ".csv":
        return _load_csv(SPREADSHEET_PATH)
    if ext in (".xlsx", ".xlsm"):
        return _load_xlsx(SPREADSHEET_PATH)
    sys.exit(f"Unsupported spreadsheet type: {ext} (use .xlsx or .csv)")


def row_skip_reason(row: dict) -> str:
    """Return a reason to skip this row, or '' if it's uploadable."""
    status = clean(row.get("status")).lower()
    if UPLOADABLE_STATUSES and status not in UPLOADABLE_STATUSES:
        return f"status={status or 'blank'}"
    if not row_title(row):
        return "no title"
    if not clean(row.get("description")):
        return "no description"
    try:
        row_date(row)
        parse_time(row.get("notify_time"))
        channel = clean(row.get("channel")).lower()
        if channel and channel not in CHANNEL_OVERRIDES:
            raise ValueError(f"Unrecognized channel {channel!r} (use push/email/both/none).")
        parse_bool_override(row.get("add_to_calendar"), False, "add_to_calendar")
        parse_bool_override(row.get("ebadge"), False, "ebadge")
    except ValueError as e:
        return str(e)
    return ""


# --------------------------------------------------------------------------- #
# Ledger (dedup by event_id across runs)
# --------------------------------------------------------------------------- #

def load_ledger() -> dict:
    if os.path.exists(LEDGER_PATH):
        with open(LEDGER_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_ledger(ledger: dict) -> None:
    with open(LEDGER_PATH, "w", encoding="utf-8") as f:
        json.dump(ledger, f, indent=1, ensure_ascii=False)


# --------------------------------------------------------------------------- #
# Modes
# --------------------------------------------------------------------------- #

def _tag_line(p: dict) -> str:
    """One-line description of a payload's tag targeting, or '' if none."""
    if not p["_tag_names"]:
        return ""
    line = f"    tags: {p['_tag_names']} → ids {p['_tag_ids']}"
    if p["_unresolved_tags"]:
        line += f"   ⚠️ UNRESOLVED (no such tag): {p['_unresolved_tags']}"
    return line


def _override_line(p: dict) -> str:
    """One-line summary of fields this row overrode away from DEFAULT_SETTINGS, or '' if none."""
    bits = []
    if (p["is_notification_push"], p["is_notification_mail"]) != (
            DEFAULT_SETTINGS["is_notification_push"], DEFAULT_SETTINGS["is_notification_mail"]):
        channel = next(k for k, v in CHANNEL_OVERRIDES.items()
                       if v == (p["is_notification_push"], p["is_notification_mail"]))
        bits.append(f"channel={channel}")
    if p["cta_text"] or p["cta_link"]:
        bits.append(f"cta={p['cta_text']!r}→{p['cta_link']!r}")
    if p["is_add_to_calendar"] != DEFAULT_SETTINGS["is_add_to_calendar"]:
        bits.append("add_to_calendar=yes")
    if p["is_ebadge_included"] != DEFAULT_SETTINGS["is_ebadge_included"]:
        bits.append("ebadge=yes")
    return f"    overrides: {', '.join(bits)}" if bits else ""


def do_dry_run(limit, default_tags=None) -> None:
    rows = load_rows()
    ledger = load_ledger()
    uploadable = [r for r in rows if not row_skip_reason(r)]
    skipped = len(rows) - len(uploadable)
    pending = [r for r in uploadable if row_key(r) not in ledger]

    print(f"DRY RUN — {SPREADSHEET_PATH} [{SHEET_NAME}]: {len(rows)} data rows, "
          f"{len(uploadable)} uploadable, {skipped} skipped (status/empty), "
          f"{len(ledger)} already created, {len(pending)} pending.")
    if default_tags:
        print(f"Applying run-wide tag(s): {default_tags}")
    show = pending if limit is None else pending[:limit]
    print(f"Would create {len(show)} announcement(s){' (--limit)' if limit else ''}:\n")
    any_tags = False
    for i, row in enumerate(show, 1):
        p = build_payload(row, default_tags)
        b = _post_body(p)
        warn = " >115 chars" if len(b["description"]) > DESCRIPTION_SOFT_LIMIT else ""
        print(f"[{i}] {b['title']}  → {p['_scheduled_for']}  (event_id {p['_event_id']})")
        print(f"    {b['description']!r}{warn}")
        tag_line = _tag_line(p)
        if tag_line:
            any_tags = True
            print(tag_line)
        override_line = _override_line(p)
        if override_line:
            print(override_line)
        print()
    if any_tags or default_tags:
        print("⚠️  Reminder: the public API drops `tags` on write — these targets are shown/"
              "prepared here, but you must apply the tag in the PheedLoop dashboard for it to\n"
              "    actually restrict the audience. Untagged announcements go to all attendees.\n")
    print("Looks right? Create them with:  python3 announcement_scheduler.py --run"
          + (f" --limit {limit}" if limit else "")
          + (f" --tag {','.join(default_tags)!r}" if default_tags else ""))


def do_run(limit, default_tags=None) -> None:
    check_credentials()
    rows = load_rows()
    ledger = load_ledger()
    pending = [r for r in rows
               if not row_skip_reason(r) and row_key(r) not in ledger]
    todo = pending if limit is None else pending[:limit]

    print(f"RUN — {len(todo)} announcement(s) to create "
          f"({len(pending)} pending total, {len(ledger)} already done).")
    if default_tags:
        print(f"Requested run-wide tag(s): {default_tags}")
    print()

    tag_dropped = False
    results = []
    for i, row in enumerate(todo, 1):
        rec = {"row": i, "event_id": row_key(row), "title": "",
               "scheduled_for": "", "code": "", "status": "", "error": ""}
        try:
            p = build_payload(row, default_tags)
            rec["title"] = p["title"]
            rec["scheduled_for"] = p["_scheduled_for"]
            created = create_announcement(p)
            rec["code"] = created.get("code", "")
            rec["status"] = created.get("status", "")
            # Verify whether the tag actually stuck (the public API usually drops it).
            if p["_tag_ids"] and not created.get("tags"):
                tag_dropped = True
            ledger[p["_event_id"]] = {"code": rec["code"], "title": rec["title"],
                                      "scheduled_for": rec["scheduled_for"]}
            save_ledger(ledger)  # persist after each success (safe to interrupt)
            tag_note = f"  tags→{created.get('tags')}" if p["_tag_ids"] else ""
            print(f"[{i}] OK   {rec['title']} → {rec['code']} "
                  f"({rec['status']}, {rec['scheduled_for']}){tag_note}")
        except Exception as e:  # noqa: BLE001
            rec["error"] = str(e)
            print(f"[{i}] FAIL {rec['title'] or '(row)'}: {e}")
        results.append(rec)

    created = sum(1 for r in results if r["code"])
    print(f"\nDone. Created {created}, failed {sum(1 for r in results if r['error'])}.")
    if tag_dropped:
        print("⚠️  A requested tag was NOT applied (the API returned tags=[]). Open those "
              "announcements in the PheedLoop dashboard and set the target tag there.")
    _write_results(results)


def _write_results(results: list) -> None:
    fields = ["row", "event_id", "title", "scheduled_for", "code", "status", "error"]
    exists = os.path.exists(RESULTS_LOG)
    with open(RESULTS_LOG, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        if not exists:
            w.writeheader()
        w.writerows(results)
    print(f"Results appended to {RESULTS_LOG}")


def do_test() -> None:
    check_credentials()
    when = datetime.now() + timedelta(minutes=5)
    date_str, time_str = when.strftime("%Y-%m-%d"), when.strftime("%H:%M:%S")
    payload = {
        "event": EVENT_CODE,
        "title": f"TEST — please ignore ({when.strftime('%H:%M')})",
        "description": "Automated scheduling test. Safe to delete.",
        "is_released": False,
        "publish_date": date_str, "publish_time": time_str, "publis_date": date_str,
        "release_date": date_str, "release_time": time_str,
        "cta_text": "", "cta_link": "",
        "is_credentials_included": False, "is_add_to_calendar": False,
        "is_ebadge_included": False, "tags": [], "tickets": [],
        # SILENT: no push/email so auto-firing can't reach real attendees.
        "is_notification_push": False, "is_notification_mail": False,
        "is_multiple_send_enabled": False,
        "is_visible_app": True, "is_visible_attendee": False,
        "is_visible_virtual": False, "is_visible_exhibitor": False,
        "is_visible_speaker": False, "is_visible_sponsor": False,
    }
    print(f"Creating a SILENT test scheduled for {date_str} {time_str} (~5 min out)...\n")
    created = create_announcement(payload)
    code = created.get("code")
    print(f"Created {code}  status={created.get('status')!r}")
    print_status(code)
    print(f"\nWait ~5 min; it should flip to Sent. Delete with: "
          f"python3 announcement_scheduler.py --delete {code}")


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Bulk-schedule PheedLoop announcements from a spreadsheet.")
    g = parser.add_mutually_exclusive_group()
    g.add_argument("--run", action="store_true", help="Create scheduled announcements.")
    g.add_argument("--test", action="store_true", help="Silent scheduling test (~5 min out).")
    g.add_argument("--status", metavar="CODE", help="Print an announcement's status.")
    g.add_argument("--delete", metavar="CODE", help="Delete an announcement by code.")
    g.add_argument("--delete-all", action="store_true",
                    help="Delete every announcement this script has created (per the ledger) "
                         "and clear the ledger, so the next --run recreates fresh from the "
                         "current spreadsheet. Prompts for confirmation.")
    parser.add_argument("--limit", type=int, default=None,
                        help="Cap how many to create/preview this run (e.g. --limit 10).")
    parser.add_argument("--tag", default=None,
                        help="Run-wide target tag name(s), comma-separated (e.g. "
                             "--tag 'Marshal Exec'). Prepared in the payload, but the API "
                             "drops it — apply/confirm in the dashboard. A per-row 'tag' "
                             "column overrides this.")
    args = parser.parse_args()
    default_tags = split_tags(args.tag)

    if args.status:
        check_credentials(); print_status(args.status)
    elif args.delete:
        check_credentials(); delete_announcement(args.delete)
    elif args.delete_all:
        check_credentials(); do_delete_all()
    elif args.test:
        do_test()
    elif args.run:
        do_run(args.limit, default_tags)
    else:
        do_dry_run(args.limit, default_tags)


if __name__ == "__main__":
    main()
