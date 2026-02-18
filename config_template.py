"""
config.py  — ORBIT Email Automation Configuration
===================================================
Copy this file to 'config.py' (in the same folder as send_emails.py)
and fill in your actual values before running.

DO NOT commit real credentials to version control.
"""

# ─── Send Method ──────────────────────────────────────────────────────────────
# "OUTLOOK" → uses the locally installed Outlook desktop app (Windows only, no credentials needed)
# "SMTP"    → sends via any SMTP relay (works on Windows/Mac/Linux)
SEND_METHOD = "OUTLOOK"    # or "SMTP"


# ─── SMTP Settings (only needed if SEND_METHOD = "SMTP") ─────────────────────
SMTP_HOST     = "smtp.office365.com"   # Your SMTP server
SMTP_PORT     = 587                    # 587 for TLS, 465 for SSL
SMTP_USE_TLS  = True
SMTP_USER     = "your.email@company.com"    # Leave "" for anonymous relay
SMTP_PASSWORD = "your_password_here"        # Leave "" for anonymous relay
SMTP_FROM     = "your.email@company.com"    # Sender address shown in To field


# ─── Email Defaults ───────────────────────────────────────────────────────────
# Domain appended to MSID usernames that don't contain "@"
# e.g. "nkelly13" → "nkelly13@yourcompany.com"
EMAIL_DOMAIN = "yourcompany.com"

# Optional CC address (leave "" to skip)
CC_EMAIL = "Orbit_PowerBI_Onboarding@yourcompany.com"
