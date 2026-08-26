# CSU Threat Intelligence Alert Automation Engine

An enterprise automation engine that monitors the Cyber Security Unit (CSU) Threat Intelligence tracker (`.xlsm` / `.xlsx`), automatically syncs real-time SharePoint data, identifies new **Alert** advisories across 25,000+ rows, formats them into branded advisories with **Dual Threat Meters** and **GSOC Assessments**, and dispatches master trigger notification emails with individual **`.msg` Outlook Item attachments** via Microsoft Outlook Desktop.

---

## 📁 Repository Structure & File Links

```text
├── advisory_alert_runner.py   # Primary automation engine (all logic in one file)
├── setup.bat                  # One-click Windows runner (auto-venv + auto-install)
├── requirements.txt           # Python dependencies (openpyxl, pywin32, Pillow)
├── TI_BG.png                  # Cyber Security Unit Threat Intelligence header banner
├── dummy_advisory.xlsx        # Sample Excel tracker workbook (38 columns)
├── .gitignore                 # Excludes .venv/, output_msg/, cache, tracking files
└── README.md                  # Comprehensive documentation and setup guide
```

### How the Files Link Together:
1. **`setup.bat` ➔ `requirements.txt` & `.venv`**:
   Automatically detects your Python installation (`py` or `python`), creates a local isolated virtual environment (`.venv`), installs all dependencies from `requirements.txt`, and launches `advisory_alert_runner.py`.
2. **`advisory_alert_runner.py` ➔ `Advisory_Tracker.xlsm` / `dummy_advisory.xlsx`**:
   Triggers Excel's background `RefreshAll()` macro to pull the latest real-time rows from SharePoint, then fast-scans all 25k+ rows.
3. **`advisory_alert_runner.py` ➔ `TI_BG.png`**:
   Encodes the CSU branded banner as an inline high-resolution asset inside the HTML template.
4. **`advisory_alert_runner.py` ➔ `output_msg/` & Microsoft Outlook**:
   Generates individual `.msg` files inside `output_msg/` and connects to Outlook Desktop via COM to transmit the master notification email with `.msg` files attached.
5. **`advisory_alert_runner.py` ➔ `.processed_ids.txt`**:
   Persists row fingerprints to prevent duplicate emails from ever being sent.

---

## ⚙️ How the Automation Works (3-Hour Loop)

```
┌────────────────────────────────────────────────────────┐
│        3-HOUR RECURRING AUTOMATION WORKFLOW            │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
 1. [Macro Auto-Refresh] Opens Excel in background, triggers 'Refresh All'
    (Power Query) to fetch live SharePoint records, saves & closes.
                           │
                           ▼
 2. [Fast Scan 25k+ Rows] Scans 25,000+ rows in read-only streaming mode (< 1s).
                           │
                           ▼
 3. [Delta Filter] Checks rows where Advisory Type == 'Alert'.
    Filters out any rows already listed in .processed_ids.txt.
                           │
                           ▼
 4. [Generate .msg Files] Builds branded individual Outlook .msg items in
    output_msg/ with TI_BG banner, Dual Threat Meters & GSOC Assessment.
                           │
                           ▼
 5. [Outlook Dispatch] Attaches all new .msg files to the master trigger email
    and sends it directly from your work Outlook profile.
                           │
                           ▼
 6. [Sleep 3 Hours] Waits 3 hours, then automatically repeats Step 1.
```

---

## 🚀 How to Run on your Work Laptop

### Prerequisites on Work Laptop:
* Windows OS
* Microsoft Outlook Desktop (signed into your work email)
* Python 3.9+ installed (with "Add Python to PATH" checked)

---

### Step 1: Clone or Copy the Repository
```bash
git clone https://github.com/chiroo09/advisorycode.git
cd advisorycode
```

### Step 2: Configure Settings (in `advisory_alert_runner.py`)
Open `advisory_alert_runner.py` in any text editor and check the configuration at the top:
```python
# Excel tracker filename
EXCEL_FILE = "Advisory_Tracker.xlsm"  # Or your actual tracker file name

# Target recipient
EMAIL_TO = "tejesh988@outlook.com"
EMAIL_CC = ""

# For initial 1-time test run, set to 0 (Set to 3 for continuous 3-hour loop)
SCHEDULE_INTERVAL_HOURS = 0

# Set to True to send automatically through Outlook
AUTO_SEND = True

# Optional: Shared mailbox address (leave empty to use default signed-in profile)
EMAIL_SENT_ON_BEHALF = ""
```

### Step 3: Run Initial One-Off Test
1. Double-click **`setup.bat`**.
2. It will:
   * Create the `.venv` and install `openpyxl`, `pywin32`, and `Pillow`.
   * Refresh the Excel tracker.
   * Generate the `.msg` files.
   * Send the master trigger email through Outlook Desktop.
3. Check your Outlook Inbox to confirm the email arrived with the `.msg` attachments!

### Step 4: Enable Continuous 3-Hour Automation
Once verified:
1. In `advisory_alert_runner.py`, set:
   ```python
   SCHEDULE_INTERVAL_HOURS = 3
   ```
2. Double-click **`setup.bat`**.
3. Leave the batch window minimized. It will now continuously monitor your Excel sheet every 3 hours.

---

## 🧪 Test Scenarios & Expected Behavior

| # | Test Scenario | Expected Result |
|---|---------------|-----------------|
| **1** | **Initial Run with Existing Alerts** | Processes all unprocessed Alert rows, creates `.msg` files in `output_msg/`, sends 1 master trigger email with all `.msg` attachments, and records IDs in `.processed_ids.txt`. |
| **2** | **Next 3-Hour Cycle (No New Alerts)** | Triggers macro refresh, scans 25k+ rows, detects no new unrecorded alert IDs, logs `[*] No new unprocessed 'Alert' records found`, and sleeps without sending duplicate emails. |
| **3** | **New Alert Added to SharePoint / Excel** | On next cycle, identifies only the new Alert row, generates its `.msg` file, and sends a trigger email with only the new alert attached. |
| **4** | **Non-Alert Rows (e.g. 'News', 'Update')** | Automatically skipped; only rows where `Advisory Type == 'Alert'` are processed. |
| **5** | **Large Workbook (25,000+ Rows)** | Fast read-only streaming iterator processes the entire workbook in under 1 second with low memory usage. |

---

## 🛡️ Email & Advisory Design Highlights

* **Master Trigger Email**:
  * Clean, professional notification referencing the date and list of new `.msg` advisory files.
  * Native Outlook `.msg` item attachments with envelope icons.
* **Individual Advisory (`.msg` file)**:
  * **Branded Header Banner**: Embedded `TI_BG.png` with Capgemini Cyber Security Unit Threat Intelligence logo.
  * **Advisory Table**: Advisory Number, Title, Summary, Severity Badge, Impact Analysis, and Vendor Solutions.
  * **Dual Threat Meters**:
    * *Vendor Meter* (Left): Scale 1, 4, 7, 9, 10 with colored indicator arrow.
    * *GSOC Meter* (Right): Scale 1, 4, 7, 9, 10 with colored indicator arrow.
  * **GSOC Assessment Section**: Risk statement, patch guidelines, and warning note:  
    `/* Test changes on non-production systems before applying on production systems */`
  * **Footer**: Official Capgemini Confidentiality and proprietary notice.

---

## 🔒 Security & Best Practices
* **No hardcoded credentials**: The engine integrates directly with Outlook Desktop's secure Windows COM profile.
* **Tracked State**: `.processed_ids.txt` ensures deduplication across power interruptions and reboots.
