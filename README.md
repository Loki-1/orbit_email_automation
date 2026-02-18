# ORBIT Power BI — Email Automation Tool

Automatically reads `.xlsx` onboarding files, builds personalised subject lines,
and sends the standard ORBIT Power BI welcome email via **Outlook** (desktop) or **SMTP**.

---

## Folder Structure

```
orbit_email_automation/
├── send_emails.py          ← Main script (run this)
├── config_template.py      ← Copy → config.py and fill in your settings
├── config.py               ← Your actual settings (NOT committed to git)
├── requirements.txt        ← Python dependencies
├── assets/
│   └── orbit_banner.jpeg   ← Yellow Power BI banner (embedded in every email)
├── input_files/            ← DROP YOUR .xlsx FILES HERE
│   └── RITM7831632.xlsx    ← Example file
└── logs/                   ← Auto-created: .log + .csv result files
```

---

## Quick Start

### Step 1 — Install dependencies
```bash
pip install -r requirements.txt
```

### Step 2 — Configure
```bash
cp config_template.py config.py
```
Open `config.py` and set:
| Setting | What to fill in |
|---|---|
| `SEND_METHOD` | `"OUTLOOK"` (Windows, no creds needed) or `"SMTP"` |
| `EMAIL_DOMAIN` | Your company domain, e.g. `uhg.com` |
| `CC_EMAIL` | CC address, e.g. `Orbit_PowerBI_Onboarding@uhg.com` (or `""` to skip) |
| `SMTP_*` | Only if using SMTP mode |

### Step 3 — Add your XLSX files
Drop all your `.xlsx` files into the `input_files/` folder.

### Step 4 — Run
```bash
python send_emails.py
```

---

## What the script does

1. Scans every `.xlsx` in `input_files/`.
2. Reads these fields from each file:

| Field | Excel row label | Example value |
|---|---|---|
| RITM | `RITM` | RITM7831632 |
| AIDE ID | `AIDE_ID` | AIDE_0094185 |
| AIDE Name | `AIDE NAME` | Intelligent Auth Accel |
| Recipient | `Application Owner` | nkelly13 |

3. Builds the subject line:
   ```
   Welcome to ORBIT Power BI - Your Guide to Getting Started - [RITM7831632][AIDE_0094185][Intelligent Auth Accel]
   ```

4. Resolves the email address:
   - If `Application Owner` already contains `@`, it is used as-is.
   - Otherwise, `@EMAIL_DOMAIN` is appended (e.g. `nkelly13` → `nkelly13@uhg.com`).

5. Sends the standard HTML email with the embedded yellow Power BI banner.

6. Writes a CSV result log to `logs/results_YYYYMMDD_HHMMSS.csv`.

---

## XLSX File Format

The script expects the **standard ORBIT onboarding spreadsheet** format, where data is
stored in key-value pairs in columns A and B:

| Column A | Column B |
|---|---|
| RITM | RITM7831632 |
| AIDE_ID | AIDE_0094185 |
| AIDE NAME | Intelligent Auth Accel |
| Application Owner | nkelly13 |

Any file that does not contain all four of these fields will be **skipped** (logged as SKIPPED).

---

## Logs

After each run you get two files in `logs/`:

- **`.log`** — Timestamped console output (SENT / FAILED / SKIPPED per file).
- **`.csv`** — Machine-readable results with columns:
  `file, ritm, aide_id, aide_name, to_email, status, error`

---

## Modes

### Outlook mode (recommended on Windows)
```python
SEND_METHOD = "OUTLOOK"
```
Uses your already-signed-in Outlook desktop client via Windows COM automation.
No SMTP credentials required. Emails appear in your Sent Items folder.
**Requires:** Windows + Outlook desktop app installed.

### SMTP mode (cross-platform)
```python
SEND_METHOD = "SMTP"
```
Connects to any SMTP relay. Fill in `SMTP_*` settings in `config.py`.
Works on Windows, macOS, and Linux.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `config.py not found` | Copy `config_template.py` → `config.py` |
| `win32com not available` | `pip install pywin32` (Windows only) |
| All files SKIPPED | Check that your XLSX has the exact row labels: `RITM`, `AIDE_ID`, `AIDE NAME`, `Application Owner` |
| Emails land in Spam | Ask your IT team to whitelist the sender or use your corporate SMTP relay |
| Wrong recipient email | Set `EMAIL_DOMAIN` in `config.py` to your company's domain |
