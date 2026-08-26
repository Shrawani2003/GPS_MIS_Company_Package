# GPS vs MIS Fleet Dashboard — Streamlit version (Python, with admin login)

A private dashboard, protected by a single admin login, that runs on your
own machine or company server. Gives you a URL that works on your
**company Wi-Fi / LAN only** — never touches the public internet. Can also
be deployed to Streamlit Community Cloud for an internet-reachable,
login-protected URL (see the bottom of this file).

## Easiest way to run it: double-click `start.bat`

1. Double-click **`start.bat`** in this folder
2. First time only, it will ask you to choose a username and password —
   type them in and press Enter after each
3. It installs everything needed automatically, then starts the dashboard
4. It will print a `Network URL` like `http://192.168.1.42:8501` — that's
   the address to share with your team

Leave that black window open while people are using the dashboard. Next
time, just double-click `start.bat` again.

To change the password later, edit `.streamlit/credentials.json` in this
folder and restart `start.bat`, or use the **Change password** panel in
the app's sidebar once logged in.

---

## Manual steps (if you'd rather not use start.bat)

Requires Python 3.10+ installed (check with `python --version`).

```
pip install -r requirements.txt
```

Copy `.streamlit/credentials.example.json` to `.streamlit/credentials.json`
and set your own username/password inside it.

Run it:
```
streamlit run app.py
```

You'll see:
```
  Local URL:   http://localhost:8501
  Network URL: http://192.168.1.42:8501
```

- **Local URL** — only works on the machine running it
- **Network URL** — works for anyone on the same office Wi-Fi/LAN. Share
  this with your team. It will NOT work for anyone outside your office
  network (no internet exposure at all).

---

## For a URL that's always available (not just while your laptop is on)

Run this on a company server or an always-on PC instead of your personal
laptop, so the Network URL stays reachable all the time.

### Auto-start on Windows

1. Press `Win + R`, type `shell:startup`, press Enter
2. Put a shortcut to `start.bat` (from this folder) inside that Startup folder
3. The dashboard now starts automatically whenever that machine turns on

### Give it a friendly name instead of an IP address (optional)

Ask your IT team to assign the server a static/reserved IP, or add an
internal DNS entry so people can type something like
`http://fleet-dashboard.yourcompany.local:8501` instead of an IP.

---

## Deploying to Streamlit Community Cloud (internet-reachable, login-protected)

1. Push this folder's contents to a GitHub repo (private is fine)
2. Go to **share.streamlit.io**, sign in with GitHub, click **New app**
3. Pick the repo/branch, set **Main file path** to `app.py`
4. Under **Advanced settings → Secrets**, add:
   ```
   ADMIN_USERNAME = "your-username"
   ADMIN_PASSWORD = "your-strong-password"
   ```
   (this stays encrypted on Streamlit's servers, never goes into the repo)
5. Click **Deploy**

The app automatically prefers these Cloud secrets over the local
`credentials.json` file when both are present, so the same codebase works
both locally and on Cloud without changes.

**Important:** do not commit your real `.streamlit/credentials.json` to
GitHub — only commit `.streamlit/credentials.example.json`.

---

## What it shows

- KPI cards: Total GPS km, Total MIS km, Overall Diff %, Vehicles, Flagged
- Day-wise GPS vs MIS trend chart
- Site-wise GPS vs MIS volume chart (single chart, scrollable box) and a
  mismatch % chart, colored by severity
- Corrective-action cards (click one to filter the vehicle table below)
- Searchable/filterable vehicle table with color-coded Diff % and Action
- Vehicle drill-down with a daily GPS vs MIS chart
- Adjustable mismatch threshold slider

## Month-over-month comparison

Uploading a new file does **not** delete previous months — nothing is ever
deleted unless you delete it yourself.

- After uploading a file, confirm/edit the **month label** (guessed from
  the filename) — it's saved to history automatically, no button to click.
- The **Total GPS km** and **Total MIS km** cards show a delta versus the
  most recently-saved earlier month.
- Use the **"Viewing"** dropdown at the top to switch between the file you
  just uploaded and any previously saved month — this is your month filter.
- Scroll to the **Monthly history** section at the bottom for a trend
  chart and table across every month you've saved, and an option to
  remove a month if you saved one by mistake.
- History (including day-wise data for uploaded months) is stored in
  `history.json` in this folder — it persists across restarts.

## Notes

- Only one admin account is supported (single username/password).
- The Excel file you upload is processed entirely on the machine running
  the app — nothing is sent externally (except when using Streamlit Cloud,
  where it's processed on Streamlit's servers, not sent anywhere else).
- Company name and logo are set directly in `app.py` (search for
  `COMPANY_NAME` and `LOGO_PATH`) — edit and restart to change them.
