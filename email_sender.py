"""
Resume Email Automation Tool
=============================
Sends your resume PDF to all HR contacts in your Excel/CSV list.
Supports Gmail (App Password) or any SMTP server.
Handles 1800+ emails with rate limiting, retries, and progress tracking.

SETUP STEPS (read README.txt first):
1. pip install -r requirements.txt
2. Edit config.py with your details
3. Place your resume PDF in this folder
4. Place your contacts Excel/CSV file in this folder
5. Run: python email_sender.py
"""

import smtplib
import time
import os
import sys
import csv
import logging
import json
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
from datetime import datetime
from pathlib import Path

import pandas as pd

# ── Import your config ──────────────────────────────────────────────────────
from config import (
    SMTP_HOST, SMTP_PORT, EMAIL_ADDRESS, EMAIL_PASSWORD,
    SENDER_NAME, SUBJECT_TEMPLATE, EMAIL_BODY_TEMPLATE,
    RESUME_PDF_PATH, CONTACTS_FILE_PATH,
    DELAY_BETWEEN_EMAILS, DELAY_JITTER,
    BATCH_SIZE, BATCH_PAUSE_SECONDS,
    MAX_RETRIES, RETRY_DELAY,
    PROGRESS_FILE, LOG_FILE,
    EMAIL_COLUMN, NAME_COLUMN, COMPANY_COLUMN, TITLE_COLUMN
)

# ── Logging Setup ────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger(__name__)


# ── Progress Tracking ────────────────────────────────────────────────────────
def load_progress():
    """Load previously sent emails to skip them on resume."""
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r") as f:
            data = json.load(f)
        log.info(f"Resuming — {len(data['sent'])} emails already sent previously.")
        return set(data["sent"]), set(data["failed"])
    return set(), set()


def save_progress(sent: set, failed: set):
    """Save progress after each email so we can resume if interrupted."""
    with open(PROGRESS_FILE, "w") as f:
        json.dump({"sent": list(sent), "failed": list(failed)}, f, indent=2)


# ── Contact Loading ──────────────────────────────────────────────────────────
def load_contacts(filepath: str) -> list[dict]:
    """Load contacts from Excel or CSV file."""
    path = Path(filepath)
    if not path.exists():
        log.error(f"Contacts file not found: {filepath}")
        sys.exit(1)

    ext = path.suffix.lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(filepath, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(filepath, dtype=str)
    else:
        log.error("Contacts file must be .xlsx, .xls, or .csv")
        sys.exit(1)

    # Strip whitespace from all values
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    df.columns = [c.strip() for c in df.columns]

    required = [EMAIL_COLUMN]
    for col in required:
        if col not in df.columns:
            log.error(f"Column '{col}' not found in contacts file. Available: {list(df.columns)}")
            sys.exit(1)

    # Drop rows with no email
    df = df[df[EMAIL_COLUMN].notna() & (df[EMAIL_COLUMN] != "")]

    contacts = df.to_dict("records")
    log.info(f"Loaded {len(contacts)} contacts from {filepath}")
    return contacts


# ── Email Building ───────────────────────────────────────────────────────────
def build_email(contact: dict, resume_bytes: bytes, resume_filename: str) -> MIMEMultipart:
    """Build a personalised email with resume attached."""
    name    = contact.get(NAME_COLUMN, "Hiring Manager") or "Hiring Manager"
    company = contact.get(COMPANY_COLUMN, "your company") or "your company"
    title   = contact.get(TITLE_COLUMN, "") or ""
    to_email = contact[EMAIL_COLUMN]

    # Personalise subject and body
    subject = SUBJECT_TEMPLATE.format(name=name, company=company, title=title)
    body    = EMAIL_BODY_TEMPLATE.format(name=name, company=company, title=title)

    msg = MIMEMultipart()
    msg["From"]    = formataddr((SENDER_NAME, EMAIL_ADDRESS))
    msg["To"]      = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Attach resume PDF
    part = MIMEBase("application", "octet-stream")
    part.set_payload(resume_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{resume_filename}"')
    msg.attach(part)

    return msg


# ── SMTP Connection ──────────────────────────────────────────────────────────
def create_smtp_connection():
    """Create and return an authenticated SMTP connection."""
    log.info(f"Connecting to SMTP: {SMTP_HOST}:{SMTP_PORT}")
    if SMTP_PORT == 465:
        server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30)
    else:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
        server.ehlo()
        server.starttls()
        server.ehlo()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    log.info("SMTP login successful.")
    return server


# ── Send One Email ───────────────────────────────────────────────────────────
def send_one(server, msg: MIMEMultipart, to_email: str):
    """Send a single email. Returns (success: bool, server) — server may be reconnected."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            server.sendmail(EMAIL_ADDRESS, [to_email], msg.as_string())
            return True, server
        except smtplib.SMTPRecipientsRefused:
            log.warning(f"  ✗ Recipient refused: {to_email}")
            return False, server  # No point retrying a bad address
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPException) as e:
            log.warning(f"  Attempt {attempt}/{MAX_RETRIES} failed for {to_email}: {e}")
            if attempt < MAX_RETRIES:
                log.info(f"  Retrying in {RETRY_DELAY}s...")
                time.sleep(RETRY_DELAY)
                try:
                    server = create_smtp_connection()
                except Exception as conn_err:
                    log.error(f"  Reconnect failed: {conn_err}")
    return False, server


# ── Main Sender Loop ─────────────────────────────────────────────────────────
def run():
    log.info("=" * 60)
    log.info("  Resume Email Automation — Starting")
    log.info("=" * 60)

    # Validate resume
    resume_path = Path(RESUME_PDF_PATH)
    if not resume_path.exists():
        log.error(f"Resume PDF not found: {RESUME_PDF_PATH}")
        sys.exit(1)
    resume_bytes    = resume_path.read_bytes()
    resume_filename = resume_path.name
    log.info(f"Resume loaded: {resume_filename} ({len(resume_bytes)//1024} KB)")

    # Load contacts
    contacts = load_contacts(CONTACTS_FILE_PATH)

    # Load progress (skip already-sent)
    sent_emails, failed_emails = load_progress()

    # Filter out already processed
    pending = [c for c in contacts if c[EMAIL_COLUMN] not in sent_emails and c[EMAIL_COLUMN] not in failed_emails]
    log.info(f"Pending: {len(pending)} | Already sent: {len(sent_emails)} | Failed: {len(failed_emails)}")

    if not pending:
        log.info("Nothing to send. All contacts already processed!")
        return

    # Ask for confirmation
    print(f"\n{'─'*60}")
    print(f"  Ready to send {len(pending)} emails from: {EMAIL_ADDRESS}")
    print(f"  Resume: {resume_filename}")
    print(f"  Estimated time: ~{len(pending) * (DELAY_BETWEEN_EMAILS + DELAY_JITTER/2) / 60:.0f} minutes")
    print(f"{'─'*60}")
    confirm = input("  Type YES to start sending: ").strip()
    if confirm.upper() != "YES":
        log.info("Cancelled by user.")
        return

    # Connect to SMTP
    server = create_smtp_connection()

    success_count = 0
    fail_count    = 0
    start_time    = datetime.now()

    try:
        for i, contact in enumerate(pending, 1):
            to_email = contact[EMAIL_COLUMN]

            log.info(f"[{i}/{len(pending)}] Sending to: {to_email} ({contact.get(COMPANY_COLUMN, '')})")

            msg = build_email(contact, resume_bytes, resume_filename)
            ok, server = send_one(server, msg, to_email)

            if ok:
                sent_emails.add(to_email)
                success_count += 1
                log.info(f"  ✓ Sent")
            else:
                failed_emails.add(to_email)
                fail_count += 1
                log.warning(f"  ✗ Failed — added to failed list")

            # Save progress every 10 emails
            if i % 10 == 0:
                save_progress(sent_emails, failed_emails)
                elapsed  = (datetime.now() - start_time).seconds
                eta_secs = (elapsed / i) * (len(pending) - i) if i > 0 else 0
                log.info(f"  Progress: {i}/{len(pending)} | Success: {success_count} | Failed: {fail_count} | ETA: {eta_secs/60:.1f} min")

            # Batch pause every BATCH_SIZE emails (avoids spam filters)
            if i % BATCH_SIZE == 0 and i < len(pending):
                log.info(f"  Batch of {BATCH_SIZE} done — pausing {BATCH_PAUSE_SECONDS}s to avoid spam filters...")
                time.sleep(BATCH_PAUSE_SECONDS)

            # Delay between emails with jitter (human-like pattern)
            delay = DELAY_BETWEEN_EMAILS + random.uniform(0, DELAY_JITTER)
            time.sleep(delay)

    except KeyboardInterrupt:
        log.info("\nInterrupted by user — saving progress...")
    finally:
        save_progress(sent_emails, failed_emails)
        server.quit()

    # Final report
    total_time = (datetime.now() - start_time).seconds / 60
    log.info("=" * 60)
    log.info(f"  DONE — Total: {len(pending)} | Sent: {success_count} | Failed: {fail_count}")
    log.info(f"  Time taken: {total_time:.1f} minutes")
    log.info(f"  Log saved to: {LOG_FILE}")
    log.info(f"  Progress saved to: {PROGRESS_FILE}")
    log.info("=" * 60)

    # Save failed list to CSV for review
    if failed_emails:
        failed_contacts = [c for c in contacts if c[EMAIL_COLUMN] in failed_emails]
        pd.DataFrame(failed_contacts).to_csv("failed_emails.csv", index=False)
        log.info(f"  Failed emails saved to: failed_emails.csv — review and retry manually")


if __name__ == "__main__":
    run()