"""
ORBIT Power BI - Email Automation Tool
========================================
Reads all .xlsx files from the 'input_files' folder,
extracts RITM, AIDE_ID, AIDE NAME, and Application Owner (email),
builds a personalized subject line, and sends a styled HTML email
via Outlook (win32com) or SMTP depending on configuration.

Usage:
    python send_emails.py
"""

import os
import sys
import pandas as pd
import logging
import csv
from datetime import datetime
from pathlib import Path

# ─── Configuration ──────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
INPUT_DIR     = BASE_DIR / "input_files"
LOGS_DIR      = BASE_DIR / "logs"
ASSETS_DIR    = BASE_DIR / "assets"
CONFIG_FILE   = BASE_DIR / "config.py"
BANNER_PATH   = ASSETS_DIR / "orbit_banner.jpeg"

LOGS_DIR.mkdir(exist_ok=True)

# ─── Logging setup ──────────────────────────────────────────────────────────
log_filename = LOGS_DIR / f"email_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# ─── Import config ──────────────────────────────────────────────────────────
try:
    sys.path.insert(0, str(BASE_DIR))
    import config as cfg
except ImportError:
    log.error("config.py not found. Copy config_template.py → config.py and fill in your settings.")
    sys.exit(1)


# ─── Parse XLSX ─────────────────────────────────────────────────────────────
def parse_xlsx(filepath: Path) -> dict | None:
    """
    Extract RITM, AIDE_ID, AIDE NAME, and Application Owner email from
    the specific key-value layout used by ORBIT onboarding spreadsheets.
    Returns a dict or None if parsing fails.
    """
    try:
        df = pd.read_excel(filepath, header=None)
    except Exception as e:
        log.error(f"Cannot open {filepath.name}: {e}")
        return None

    data = {}
    # Build a lookup: first column value → second column value
    for _, row in df.iterrows():
        key = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        val = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
        if key and val and val.lower() not in ("nan", ""):
            data[key] = val

    ritm          = data.get("RITM", "")
    aide_id       = data.get("AIDE_ID", "")
    aide_name     = data.get("AIDE NAME", "")
    app_owner_raw = data.get("Application Owner", "")

    if not all([ritm, aide_id, aide_name, app_owner_raw]):
        log.warning(
            f"[{filepath.name}] Missing fields — "
            f"RITM='{ritm}' AIDE_ID='{aide_id}' AIDE_NAME='{aide_name}' Owner='{app_owner_raw}'"
        )
        return None

    # Resolve email: if it looks like an MSID (no @), append the domain
    to_email = app_owner_raw if "@" in app_owner_raw else f"{app_owner_raw}@{cfg.EMAIL_DOMAIN}"

    return {
        "file":      filepath.name,
        "ritm":      ritm,
        "aide_id":   aide_id,
        "aide_name": aide_name,
        "to_email":  to_email,
    }


# ─── Build email body ────────────────────────────────────────────────────────
def build_html_body(banner_cid: str = "orbit_banner") -> str:
    """Return the full HTML email body with embedded banner (via CID)."""
    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body       {{ font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #1a1a1a; margin: 0; padding: 0; background: #f4f4f4; }}
  .wrapper   {{ max-width: 720px; margin: 20px auto; background: #ffffff; border-radius: 6px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,.12); }}
  .banner    {{ width: 100%; display: block; }}
  .title-bar {{ background: #1e3a5f; color: #ffffff; padding: 18px 28px; }}
  .title-bar h1 {{ margin: 0; font-size: 22px; font-weight: bold; }}
  .title-bar h2 {{ margin: 4px 0 0; font-size: 16px; font-weight: normal; opacity: 0.9; }}
  .body-wrap {{ padding: 24px 28px; line-height: 1.7; }}
  .body-wrap p {{ margin: 0 0 12px; }}
  .step-header {{ background: #f0b429; color: #1a1a1a; padding: 8px 14px; border-radius: 4px; font-weight: bold; margin: 20px 0 8px; }}
  ul  {{ margin: 6px 0 12px 20px; padding: 0; }}
  li  {{ margin-bottom: 5px; }}
  a   {{ color: #0063b1; }}
  .footer {{ background: #1e3a5f; color: #cce0ff; padding: 14px 28px; font-size: 12px; }}
  .footer a {{ color: #90c5f7; }}
  .sig {{ margin-top: 20px; color: #555; font-style: italic; }}
</style>
</head>
<body>
<div class="wrapper">

  <!-- Yellow Power BI banner -->
  <img src="cid:{banner_cid}" class="banner" alt="ORBIT Power BI Banner" />

  <!-- Dark-blue title bar -->
  <div class="title-bar">
    <h1>Welcome to ORBIT Power BI</h1>
    <h2>Your Guide to Getting Started</h2>
  </div>

  <!-- Main body -->
  <div class="body-wrap">
    <p>Dear Customer,</p>
    <p>Welcome to Power BI! We're thrilled to have you onboard.</p>
    <p>The requests for developers and license provisions in SECURE have been processed. An email should
    have been sent to you for the approval of your application transition from platform global groups to
    application roles within SECURE. If you haven't checked yet, please look for this email in your inbox.
    Once you receive and approve it, new developers and users can then follow the instructions provided to
    request their roles in SECURE moving forward.</p>
    <p>As the new workspace owner, your team now has the responsibility to guide users and provide them
    with access instructions. It will also be your team's duty to support their content.</p>
    <p>It is important to remind you not to share the report URL or approve requests via emails.
    Report viewers should request their access via SECURE and access the reports directly from the
    application, not the workspace. Please refer to the instructions below for both users and developers.</p>
    <p>Here are some steps to help you get started:</p>

    <!-- Step 1 -->
    <div class="step-header">1. Request Access</div>
    <p>Access to Power BI is managed via the SECURE application. Please follow one of these guides to
    request access:</p>
    <ul>
      <li>Watch <a href="https://uhgazure.sharepoint.com/teams/orbit-powerbi/Shared%20Documents/Forms/AllItems.aspx?id=%2Fteams%2Forbit%2Dpowerbi%2FShared%20Documents%2FApps%2FYammer%2FUserGuide%2Fguide%2Dsecure%2Drequest%2Daccess%2Dapplication%2Ehtm&parent=%2Fteams%2Forbit%2Dpowerbi%2FShared%20Documents%2FApps%2FYammer%2FUserGuide&p=true&ga=1">Request Access</a> animation.</li>
      <li>Read <a href="https://uhgazure.sharepoint.com/teams/orbit-powerbi/Shared%20Documents/Apps/Yammer/Orbit%20PowerBI%20Secure%20Access%20Guide.pdf">Request Access</a> instructions.</li>
    </ul>

    <!-- Step 2 -->
    <div class="step-header">2. Access Power BI</div>
    <p>Visit <a href="https://powerbi.com">https://powerbi.com</a>.</p>

    <!-- Step 3 -->
    <div class="step-header">3. User Instructions</div>
    <p>Here are two methods to access Power BI:</p>
    <ul>
      <li><strong>Option 1 (Preferred):</strong> Pin Power BI to your Teams for easy access.
          Watch <a href="https://uhgazure.sharepoint.com/:u:/t/orbit-powerbi/EY8o5BVC6H9MsIlVDoi-klkBXiyIhr-MCvKkdmix_UY5OQ?e=zOPQCa">Access via Teams</a> animation.</li>
      <li><strong>Option 2:</strong> Visit <a href="https://powerbi.com">https://powerbi.com</a>. Select "Apps" on the left panel.</li>
    </ul>

    <!-- Step 4 -->
    <div class="step-header">4. Developer Instructions</div>
    <ul>
      <li><strong>Accessing Workspaces:</strong> Visit <a href="https://powerbi.com">https://powerbi.com</a>.
          You can see your workspaces (Production and Non-Production) from your navigation pane.</li>
      <li><strong>Publishing a Power BI Report:</strong>
        <ul>
          <li>Open Power BI Desktop → File → Publish to Power BI.</li>
          <li>Select the name of the App Workspace you want to publish to.</li>
        </ul>
      </li>
      <li><strong>Publishing workspace as an application:</strong>
        <ul>
          <li>Visit <a href="https://powerbi.com">https://powerbi.com</a> → select your workspace → Update app.</li>
          <li>Setup tab: name, description, logo, color.</li>
          <li>Content tab: add reports. Security tab: enable reports. Click <em>Publish app</em>.</li>
          <li>Watch <a href="https://uhgazure.sharepoint.com/:i:/t/orbit-powerbi/ES7z3gdR1G5Hi9k0Oz0-QeQBAAqmt-_FXU83u1SPdb1wAw?e=396kka">Publish Workspace App</a> animation.</li>
        </ul>
      </li>
      <li><strong>Connecting Datasets to the Database:</strong>
        <ul>
          <li><u>On-Premises / Cloud (Snowflake, Databricks, etc.):</u>
            Cloud connections only — <a href="https://uhgazure.sharepoint.com/teams/orbit-powerbi/ESTK87R2jXdKnFa8EXrmt-IBgih4wr9BPzLdQfuVcEvHZA?e=FT4Gvq">Set up a VNET</a>.<br>
            To request a gateway connection, provide a non-user ID with database access and submit a
            <a href="https://helpdesktools.uhg.com/">ServiceNow</a> ticket assigned to the
            <em>ORBIT - POWER BI - PLATFORM ADMIN</em> team.</li>
          <li><u>SharePoint / OneDrive connections:</u> Select <em>Data source credentials</em> → choose OAuth2 and Sign in to authenticate.</li>
        </ul>
      </li>
      <li><strong>Optional — Schedule dataset refresh:</strong> You MUST configure the dataset connection first.</li>
      <li><strong>Optional — Request audiences:</strong> For new onboarded tenants, contact ORBIT. To add new audiences after onboarding, follow the
          <a href="https://uhgazure.sharepoint.com/:p:/t/orbit-powerbi/EWOGesD4DcxDtKBhg7a_hOcBIW6jet24qAwAsUpq6t7kMw?e=pkNTG5">How to request Audiences</a> instruction.</li>
    </ul>

    <!-- Step 5 -->
    <div class="step-header">5. Additional Resources</div>
    <ul>
      <li><a href="https://uhgazure.sharepoint.com/sites/Welcome-To-OrbitPower-BI/SitePages/Training.aspx">Free Training</a></li>
      <li><a href="https://uhgazure.sharepoint.com/sites/Welcome-To-OrbitPower-BI/SitePages/Other-ORBIT-Tools.aspx">Free Software</a></li>
      <li><a href="https://uhgazure.sharepoint.com/sites/Welcome-To-OrbitPower-BI/SitePages/Support.aspx">Platform Support</a></li>
      <li><a href="https://uhgazure.sharepoint.com/sites/Welcome-To-OrbitPower-BI/_layouts/15/Events.aspx?ListGuid=49238fc8-e89a-420c-9ff1-b5e1badff306">Office Hours</a></li>
      <li><a href="mailto:odmbusinessintelligence_dl@ds.uhc.com">Development Services</a></li>
      <li><a href="https://uhgazure.sharepoint.com/sites/Welcome-To-OrbitPower-BI/SitePages/Cost%20and%20Access.aspx">Cost</a></li>
      <li><a href="https://uhgazure.sharepoint.com/sites/Welcome-To-OrbitPower-BI/SitePages/Support.aspx">ORBIT Welcome Page</a></li>
      <li><a href="https://uhgazure.sharepoint.com/teams/OrbitPowerBIForms/SitePages/Power-BI-Rules.aspx#development-best-practices">ORBIT Power BI Rules &amp; Best Practices</a></li>
    </ul>

    <p class="sig">Best,<br><strong>[Dharani]</strong></p>
  </div>

  <!-- Footer bar -->
  <div class="footer">
    ORBIT Power BI Platform &nbsp;|&nbsp;
    <a href="https://uhgazure.sharepoint.com/sites/Welcome-To-OrbitPower-BI">Welcome Page</a> &nbsp;|&nbsp;
    <a href="mailto:odmbusinessintelligence_dl@ds.uhc.com">Contact Us</a>
  </div>

</div>
</body>
</html>
"""


# ─── Send via Outlook (win32com) ─────────────────────────────────────────────
def send_via_outlook(record: dict, html_body: str) -> bool:
    try:
        import win32com.client as win32
    except ImportError:
        log.error("win32com not available. Install with: pip install pywin32")
        return False

    try:
        outlook = win32.Dispatch("outlook.application")
        mail    = outlook.CreateItem(0)
        mail.To = record["to_email"]

        if cfg.CC_EMAIL:
            mail.CC = cfg.CC_EMAIL

        mail.Subject = (
            f"Welcome to ORBIT Power BI - Your Guide to Getting Started - "
            f"[{record['ritm']}][{record['aide_id']}][{record['aide_name']}]"
        )

        # Attach banner as inline image
        attachment = mail.Attachments.Add(str(BANNER_PATH))
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "orbit_banner"
        )

        mail.HTMLBody = html_body
        mail.Send()
        log.info(f"[SENT] {record['file']} → {record['to_email']}")
        return True

    except Exception as e:
        log.error(f"[FAILED] {record['file']} → {record['to_email']} : {e}")
        return False


# ─── Send via SMTP ────────────────────────────────────────────────────────────
def send_via_smtp(record: dict, html_body: str) -> bool:
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText
    from email.mime.image     import MIMEImage

    subject = (
        f"Welcome to ORBIT Power BI - Your Guide to Getting Started - "
        f"[{record['ritm']}][{record['aide_id']}][{record['aide_name']}]"
    )

    msg = MIMEMultipart("related")
    msg["From"]    = cfg.SMTP_FROM
    msg["To"]      = record["to_email"]
    msg["Subject"] = subject
    if cfg.CC_EMAIL:
        msg["Cc"] = cfg.CC_EMAIL

    alt = MIMEMultipart("alternative")
    msg.attach(alt)
    alt.attach(MIMEText(html_body, "html"))

    # Embed banner image
    try:
        with open(BANNER_PATH, "rb") as img_f:
            img = MIMEImage(img_f.read(), _subtype="jpeg")
            img.add_header("Content-ID", "<orbit_banner>")
            img.add_header("Content-Disposition", "inline", filename="orbit_banner.jpeg")
            msg.attach(img)
    except FileNotFoundError:
        log.warning("Banner image not found — email will send without it.")

    try:
        with smtplib.SMTP(cfg.SMTP_HOST, cfg.SMTP_PORT) as server:
            if cfg.SMTP_USE_TLS:
                server.starttls()
            if cfg.SMTP_USER and cfg.SMTP_PASSWORD:
                server.login(cfg.SMTP_USER, cfg.SMTP_PASSWORD)
            recipients = [record["to_email"]]
            if cfg.CC_EMAIL:
                recipients.append(cfg.CC_EMAIL)
            server.sendmail(cfg.SMTP_FROM, recipients, msg.as_string())
        log.info(f"[SENT] {record['file']} → {record['to_email']}")
        return True
    except Exception as e:
        log.error(f"[FAILED] {record['file']} → {record['to_email']} : {e}")
        return False


# ─── CSV Results logger ───────────────────────────────────────────────────────
def write_csv_log(results: list[dict]):
    csv_path = LOGS_DIR / f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    fieldnames = ["file", "ritm", "aide_id", "aide_name", "to_email", "status", "error"]
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(results)
    log.info(f"Results saved → {csv_path}")


# ─── Main ────────────────────────────────────────────────────────────────────
def main():
    xlsx_files = sorted(INPUT_DIR.glob("*.xlsx"))
    if not xlsx_files:
        log.warning(f"No .xlsx files found in {INPUT_DIR}")
        return

    log.info(f"Found {len(xlsx_files)} file(s) in {INPUT_DIR}")
    html_body = build_html_body()
    results   = []

    for xlsx_path in xlsx_files:
        log.info(f"Processing: {xlsx_path.name}")
        record = parse_xlsx(xlsx_path)
        if record is None:
            results.append({"file": xlsx_path.name, "ritm": "", "aide_id": "",
                             "aide_name": "", "to_email": "", "status": "SKIPPED",
                             "error": "Missing required fields"})
            continue

        if cfg.SEND_METHOD.upper() == "OUTLOOK":
            ok = send_via_outlook(record, html_body)
        else:
            ok = send_via_smtp(record, html_body)

        results.append({**record,
                        "status": "SENT" if ok else "FAILED",
                        "error":  "" if ok else "See log"})

    write_csv_log(results)
    sent    = sum(1 for r in results if r["status"] == "SENT")
    failed  = sum(1 for r in results if r["status"] == "FAILED")
    skipped = sum(1 for r in results if r["status"] == "SKIPPED")
    log.info(f"Done — Sent: {sent} | Failed: {failed} | Skipped: {skipped}")


if __name__ == "__main__":
    main()
