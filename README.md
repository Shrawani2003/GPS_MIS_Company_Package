# GPS vs MIS Fleet Dashboard — Streamlit version (Python, with admin login)

A private dashboard, protected by a single admin login, that runs on your
own machine or company server. Gives you a URL that works on your
**company Wi-Fi / LAN only** — never touches the public internet.

## Easiest way to run it: double-click `start.bat`

1. Double-click **`start.bat`** in this folder
2. First time only, it will ask you to choose a username and password —
   type them in and press Enter after each
3. It installs everything needed automatically, then starts the dashboard
4. It will print a `Network URL` like `http://192.168.1.42:8501` — that's
   the address to share with your team

Leave that black window open while people are using the dashboard. Next
time, just double-click `start.bat` again.

To change the password later, log in and use **"🔑 Change password"** in
the left sidebar — no file editing needed.

---

## Manual steps (if you'd rather not use start.bat)

**Step 1: Open the folder**
File → Open Folder → select the `streamlit-dashboard` folder in VS Code.

**Step 2: Open a terminal**
Terminal → New Terminal.

**Step 3: Install the requirements** (only needed once)
```
pip install -r requirements.txt
```

**Step 4: Set your admin username & password**
Inside the `.streamlit` folder there's a file called `credentials.example.json`.
1. Copy/rename it to `credentials.json` (remove `.example`)
2. Open it and set your own username and password:
```json
{
  "username": "admin",
  "password": "your-own-strong-password"
}
```
3. Save it.

⚠️ Keep this file private — never share it or upload it anywhere public.
It holds your login password in plain text.

**Step 5: Run the app**
```
streamlit run app.py
```

You'll see:
```
  Local URL:   http://localhost:8501
  Network URL: http://192.168.1.42:8501
```

**Step 6: Open it**
Ctrl+Click the `Local URL` (or paste it into your browser). You'll see the
login screen first — log in with the username/password you set in Step 4.

---

## Turn this into a permanent company URL

### Option A — Run it on an always-on company PC or server

1. Copy the whole `streamlit-dashboard` folder onto a company server or a
   PC that stays on and connected to the network (not your personal laptop)
2. On that machine, double-click `start.bat` (or follow the manual steps above)
3. Share that machine's **Network URL** with your team — this stays live
   as long as that machine is on

### Option B — Make it auto-start on Windows

On the server/PC from Option A:

1. Press `Win + R`, type `shell:startup`, press Enter — opens the Windows
   Startup folder
2. Create a shortcut to `start.bat` (from this folder) inside that Startup folder
3. Now the dashboard starts automatically every time that machine turns on
   or a user logs in

### Option C — Give it a friendly name instead of an IP address (optional)

IP addresses like `192.168.1.42` can change over time. Ask your IT team to:
- Assign that server a **static/reserved IP** in your router/DHCP settings, or
- Add an internal DNS entry so people can type something like
  `http://fleet-dashboard.yourcompany.local:8501` instead of an IP

---

## Sidebar: password, branding

After logging in, the left sidebar has two expandable sections:

- **🔑 Change password** — enter your current password once, then your new
  password twice. Takes effect immediately, no restart needed.
- **🏢 Company branding** — set your company name and upload a logo image
  (PNG/JPG). Both are saved to disk and show at the top of the sidebar for
  everyone who logs in afterward.

## Month-over-month comparison

Uploading a new file does **not** delete previous months — nothing is ever
deleted unless you delete it yourself.

- After uploading a file, confirm/edit the **month label** (guessed from the
  filename) — it's saved to history automatically, no button to click.
- The **Total GPS km** and **Total MIS km** cards then show a delta versus
  the most recently-saved earlier month.
- Use the **"Viewing"** dropdown at the top to switch between the file you
  just uploaded and any previously saved month.
- Scroll to the **Monthly history** section at the bottom for a trend chart
  and table across every saved month, with an option to remove a month.
- History is stored in `history.json` in this folder — persists across restarts.

## Notes

- Only one admin account is supported (as requested) — anyone with the
  username/password can log in from anywhere on the company network.
- All files (`credentials.json`, `branding.json`, `logo.png`, `history.json`)
  live in this folder / its `.streamlit` subfolder and persist across restarts.
- The Excel file you upload is processed entirely on the machine running
  the app — nothing is sent externally.
