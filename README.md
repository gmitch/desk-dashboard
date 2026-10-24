# Desk Dashboard

A glanceable desk display: big clock, local weather, and your current /
next meeting with its room. Built for a cheap Android tablet in a stand,
hardwired and always on.

```
[Tablet kiosk page] --(JSONP, key in request)--> [Apps Script web app]
        ^                                               |
        |                                               v
[GitHub Pages: index.html]                  [Google Sheet of meetings]
Weather: Open-Meteo, fetched directly by the page (no key needed).
```

Security model: the page contains **no data and no secrets**. The Apps
Script checks a shared key on every request; the key lives in the
script's properties and in the tablet's local storage only — it is never
committed to this repo. The script returns only title, room, and times.

## The "Where's Mitch?" card

When Mitch is in a meeting, the main card shows the meeting (title,
room, start – end). When he isn't, it shows a one-line, AI-written guess
at where he is, based on the situation — after hours it's home/dinner/
family themed, the small hours are sleep themed, lunchtime is lunch,
work-hours gaps are break-room/walk themed, weekends are family themed.

- Generated inside the Apps Script by Gemini (`gemini-2.0-flash`) when
  the optional `AI_API_KEY` script property is set — get a free key at
  aistudio.google.com/apikey. The key lives only in Script Properties.
- Cached 30 minutes per situation, so the model is called a handful of
  times a day even though the page polls every minute.
- With no key (or if the call fails), built-in fallback lines rotate
  instead — the card never breaks.

## Files

- `index.html` — the dashboard page (GitHub Pages entry point).
- `config.example.js` — copy to `config.js` and set your endpoint URL.
- `apps-script/Code.gs` — the key-gated meeting endpoint.
- `apps-script/appsscript.json` — script manifest (timezone).

## Setup

### 1. The meetings sheet
Create a Google Sheet (on the same Google account that will own the
Apps Script — your personal one; no work account is involved). One row
per recurring weekly meeting:

| A (Day + start) | B (Room) | C (Title) | D (End, optional) |
|---|---|---|---|
| Fri 10:00 | Maple Room | Weekly Design Review | |
| Mon 9:30 AM | Room A | Standup | 9:45 AM |

Day is Mon–Sun, 24h or AM/PM times both work. Without an end time a
meeting is assumed to last 60 minutes. Note the sheet's ID from its URL.

### 2. The Apps Script endpoint (bound to the sheet, managed with clasp)
Bound is the easy setup here: the script lives inside the spreadsheet,
so it reads its own sheet — no SHEET_ID to configure — and you open the
editor right from the sheet.

1. In the meetings sheet: **Extensions → Apps Script** (this creates the
   bound project). Then **Project Settings** → note the **Script ID**,
   and under **Script Properties** add:
   - `DASH_KEY` — a long random key you invent. Treat it like a password.
   - `SHEET_TAB` — optional, if your data isn't on the first tab.
   - `AI_API_KEY` — optional, enables AI-written "Where's Mitch?" quips
     (free Gemini key from aistudio.google.com/apikey).
2. Push the code with clasp (from this repo):
   ```
   cd apps-script
   cp .clasp.json.example .clasp.json   # paste in the Script ID
   clasp push
   ```
   (`.clasp.json` is gitignored — it's per-machine config, not code.)
3. In the web editor: **Deploy → New deployment → Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Copy the deployment URL (`.../exec`).

   CLI alternative: the manifest (`appsscript.json`) already carries the
   web app config (access: ANYONE, execute as: you), so
   `clasp deploy --description "desk-dashboard"` creates the same web app
   deployment without the editor — take the `/exec` URL from the
   deployment ID it prints.

"Anyone" is safe here: the URL is unguessable, and the script returns
nothing without the key — it checks the key in constant time.

After future `clasp push`es, publish the new code via **Deploy →
Manage deployments → Edit → Version: New version** (same deployment,
same URL — the tablet never needs reconfiguring).

### 3. The page
Push this repo to GitHub and enable Settings → Pages → deploy from
branch (`main`). Your page: `https://<user>.github.io/desk-dashboard/`

Nothing deployment-specific goes in the repo — no key, and not even the
endpoint URL. Both are entered on the tablet in step 4 and stored only
there. (`config.js` is gitignored on purpose; `config.example.js`
documents the optional file-based config if you ever self-host instead.)

### 4. The tablet
1. Open the Pages URL in Chrome (or Fully Kiosk Browser for a stricter
   kiosk mode). The setup form asks for the endpoint URL (the `/exec`
   link from step 2) and the `DASH_KEY`. Both are stored on the tablet
   only.
   - Easier: build a one-time setup link on your computer —
     `https://<user>.github.io/desk-dashboard/?endpoint=<url-encoded /exec URL>&key=<your key>`
     — open it once on the tablet and the page saves both values, then
     scrubs them from the address bar.
2. Use the browser's "Add to Home screen" / kiosk mode, keep it plugged
   in, and set the screen to never sleep.

The page re-fetches meetings every minute, weather every 15 minutes,
and reloads itself every 6 hours. The ⚙ in the corner re-opens the key
prompt if you ever rotate the key.

## Changing the key
Update `DASH_KEY` in Script Properties, then tap ⚙ on the tablet and
enter the new key. No redeploy needed for property changes. (Code
changes do need a new deployment or "Manage deployments → edit".)
