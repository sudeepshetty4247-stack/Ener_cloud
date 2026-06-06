# EnerCloud — Complete Setup Guide

## What you have (5 files)

| File | Purpose | Language |
|------|---------|----------|
| `file1_iot_simulator.py` | Sends fake city energy data to Firebase every 5s | Python |
| `file3_ai_decision_engine.py` | Reads Firebase, runs Solar/Battery/Grid logic, writes decisions | Python |
| `file4_live_dashboard.html` | Real-time dashboard that auto-updates from Firebase | HTML/JS |
| `file5_report_generator.py` | Pulls all data from Firebase → generates Excel report | Python |
| `SETUP_GUIDE.md` | This file | — |

---

## Step 1 — Firebase Setup (do this FIRST, ~20 min)

1. Go to https://console.firebase.google.com
2. Click **"Add project"** → name it `EnerCloud` → Continue (no Google Analytics needed)
3. Inside the project → left sidebar → **Build → Realtime Database**
4. Click **"Create Database"** → choose any region → **"Start in test mode"** → Enable
5. **Get your Database URL** — it looks like:
   `https://enercloud-default-rtdb.firebaseio.com`
   Copy this. You'll paste it into all 3 Python files.

### Get your Firebase Config (for the dashboard HTML):
1. Click the gear icon ⚙️ → **Project settings**
2. Scroll to **"Your apps"** → click **`</>`** (Web app)
3. Name it `EnerCloud Dashboard` → Register app
4. Copy the `firebaseConfig` object — you'll paste it into `file4_live_dashboard.html`

### Get your Service Account Key (for Python files):
1. Click the gear icon ⚙️ → **Project settings** → **Service accounts** tab
2. Click **"Generate new private key"** → Download JSON
3. Rename the file to `serviceAccountKey.json`
4. Put it in the **same folder** as your Python scripts

---

## Step 2 — Install Python packages

Open terminal and run:

```bash
pip install firebase-admin openpyxl
```

---

## Step 3 — Update your Firebase credentials in every file

In **file1_iot_simulator.py** (line ~20):
```python
DATABASE_URL = "https://YOUR-PROJECT-ID-default-rtdb.firebaseio.com"
```

In **file3_ai_decision_engine.py** (line ~21):
```python
DATABASE_URL = "https://YOUR-PROJECT-ID-default-rtdb.firebaseio.com"
```

In **file4_live_dashboard.html** (the `firebaseConfig` block ~line 160):
```javascript
const firebaseConfig = {
    apiKey: "YOUR_API_KEY",
    authDomain: "YOUR_PROJECT_ID.firebaseapp.com",
    databaseURL: "https://YOUR_PROJECT_ID-default-rtdb.firebaseio.com",
    projectId: "YOUR_PROJECT_ID",
    ...
};
```

In **file5_report_generator.py** (line ~20):
```python
DATABASE_URL = "https://YOUR-PROJECT-ID-default-rtdb.firebaseio.com"
```

---

## Step 4 — Run everything (build order)

### Terminal 1 — IoT Simulator
```bash
python file1_iot_simulator.py
```
You'll see city data flowing every 5 seconds. Keep this running.

### Terminal 2 — AI Decision Engine
```bash
python file3_ai_decision_engine.py
```
This reads from Firebase and writes decisions back. Keep this running.

### Browser — Live Dashboard
Open `file4_live_dashboard.html` in Chrome. Charts will update automatically as data arrives.

### Generate Excel Report (whenever you want)
```bash
python file5_report_generator.py
```
Creates `EnerCloud_Energy_Report_<date>.xlsx` in the current folder.

---

## Step 5 — Demo script (what to say to judges)

1. **Open Firebase console** in browser — show database URL and live data appearing
2. **Open a terminal** → run `python file1_iot_simulator.py` → judges see city data streaming
3. **Open Chrome** → open `file4_live_dashboard.html` → charts updating live
4. **Open second terminal** → run `python file3_ai_decision_engine.py` → decisions appear on dashboard
5. **Run report** → `python file5_report_generator.py` → Excel file in 3 seconds
6. Say: *"This is exactly how a real government energy system works — IoT devices send data to the cloud, AI makes decisions, and everything is monitored in real time."*

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `FileNotFoundError: serviceAccountKey.json` | Make sure the JSON file is in the same folder as the Python scripts |
| Firebase permission denied | Go to Firebase → Realtime Database → Rules → set both read and write to `true` |
| Dashboard shows no data | Check that the `firebaseConfig` values are correct in the HTML file |
| `ModuleNotFoundError` | Run `pip install firebase-admin openpyxl` |

---

## File 6 — PPT + Video (do last, manually)

After everything works:
1. **Canva** (canva.com) → 8-slide pitch deck → use the flow: Problem → Solution → How It Works → Demo → Cities → Data → AI → Government Impact
2. **InVideo** (invideo.io) → 60-second AI video → paste your script → export

Good luck with the demo! 🚀
