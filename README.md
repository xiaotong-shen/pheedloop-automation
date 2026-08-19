# PheedLoop Automation Script  (SMC Orientation 2026)

Scripts that automate PheedLoop for one event: **SMC Orientation 2026**
(org `ORG2J3207CTQTEFVBT`, event `EVEMFFRKMCZTH`), because PheedLoop's
whitelabel app builder has no bulk-upload for announcements.

## Normal workflow (during orientation week)

1. The anyone can edit `notifications.xlsx` (sheet "All Notifications") —
   add/change rows for announcements: date, notify time, title,
   description, location, etc.
2. Notify Xiaotong of the change and I'll run it on my end (should always be an option, but if i'm not here see **[Backup tutorial](#backup-tutorial--if-xiaotong-is-unavailable)** below , you can also try for yourself for fun if you want after reading the doc!)
3. I'll run a dry run to review, then `--run` to actually schedule
   the announcements in PheedLoop. See [Commands](#commands) below to try for yourself:



## Files in this folder

| File | What it is |
|---|---|
| `announcement_scheduler.py` | The script. Bulk-creates scheduled PheedLoop announcements from `notifications.xlsx`. |
| `notifications.xlsx` | **Live working spreadsheet** — sheet "All Notifications" is what gets uploaded. This is the file the committee edits. |
| `announcement_template.xlsx` | Blank starter template (with an Instructions sheet) if you ever need to rebuild the spreadsheet from scratch. |
| `created_announcements.json` | **Dedup ledger — do not delete.** Every announcement the script has created gets recorded here so re-running `--run` never creates duplicates. |
| `.env` | Real API credentials (org code, event code, API key/secret). **Never commit this** — it's git-ignored. |
| `.env.example` | Template showing which variables `.env` needs. |
| `requirements.txt` | Python deps: `requests`, `openpyxl`, `python-dotenv`. |

btw i will share api keys seperately on teams kjust be safe, bc once it's committed it's in git history forever (recoverable even after you delete it later), and GitHub scans public repos for exposed keys within minutes. Whoever has our `API_KEY`/`API_SECRET` can hit the same endpoints this script does — no PheedLoop login needed — so they could blast a fake announcement to all attendees or delete/edit real ones. `.env.example` (field names only) is safe to commit; `.env` (real values) never is.

so even though the event isn't super high stakes,still  im scared of putting the keys on github ahaha

## Commands

```bash
pip install -r requirements.txt      # one-time setup
cp .env.example .env                 # then fill in real values (see below)

python3 announcement_scheduler.py              # dry run — prints what would be created, no API calls
python3 announcement_scheduler.py --run        # actually creates scheduled announcements
python3 announcement_scheduler.py --run --limit 5   # cap how many this run
python3 announcement_scheduler.py --test       # one throwaway silent test ~5 min out
python3 announcement_scheduler.py --status ANNOUNCEMENT-CODE   # check one announcement
python3 announcement_scheduler.py --delete ANNOUNCEMENT-CODE   # delete one
python3 announcement_scheduler.py --delete-all                 # wipe everything this script created, re-run to rebuild
```

Always dry-run first and read the output before `--run`.

Note: you don't
need to strip out old rows before re-uploading. `--run` already skips any row whose
dedup key (`date + notify_time + title`) is already in `created_announcements.json`
— confirmed in [announcement_scheduler.py](announcement_scheduler.py) (`do_run`
filters `row_key(r) not in ledger`). So uploading the *whole* sheet, old rows and
new ones together, is safe — only the new rows get created.

However, that dedup key is derived from the row's own title/date/time. If you **edit** a row that's already been created (fix a typo in the title, change the time), its key changes → the script won't recognize it as "already done" → `--run` creates a **second, duplicate** announcement instead of updating the old one, and the old one just sits there unchanged. There's no update-in-place yet.

So: adding brand-new rows and re-running is always safe. To *change* something already created, either edit it directly in the PheedLoop dashboard, or use `--delete-all` to wipe everything this script made and `--run` again to rebuild the whole sheet from scratch (asks for a typed confirmation first, so it's hard to fire by accident).

## Backup tutorial to run by yourself

**You will need:** the PheedLoop **API Key** and **API Secret** for this
event. Get these from the PheedLoop Dashboard → Settings → Integrations
→ REST API 

I'll also share them in chat (teams)

1. **Install Python** (skip if already installed): go to
   [python.org/downloads](https://www.python.org/downloads/), download
   and run the installer for your OS. Use any version 3.9+.
2. **Get this folder onto your computer.** Either:
   - Download the GitHub repo as a ZIP (green "Code" button → "Download ZIP") and unzip it, or
   - If you're comfortable with git: `git clone <repo-url>`
3. **Open a terminal in that folder.**
   - Mac: right-click the folder → "New Terminal at Folder" (or open
     Terminal and `cd` into it).
   - Windows: open the folder in File Explorer, click the address bar,
     type `cmd`, press Enter.
4. **Install the dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
5. **Set up credentials:**
   - Copy `.env.example` to a new file named `.env` in the same folder.
   - Open `.env` in any text editor and fill in `PHEEDLOOP_API_KEY` and
     `PHEEDLOOP_API_SECRET` with the values from the Dashboard (or what i pasted in the teams chat). The org/event codes are already filled in.
6. **Edit the spreadsheet:** edit and download the sheet as `notifications.xlsx`, go to the "All
   Notifications" sheet, add your row(s). Leave `status` as `draft` —
   only draft rows get uploaded.
7. **Preview, then run:**
   ```bash
   python3 announcement_scheduler.py
   ```
   Read the printed preview carefully — check dates, times, and titles.
   If it looks right:
   ```bash
   python3 announcement_scheduler.py --run
   ```
8. **Double check in the PheedLoop Dashboard** that the announcement
   shows up under Announcements with status "Scheduled" for the right time.