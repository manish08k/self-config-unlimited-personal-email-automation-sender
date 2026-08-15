"""
Resume Email Automation - Multi-Account Web App
Rotates across multiple Gmail accounts (each with its own daily limit).
When an account hits its daily cap, it auto-switches to the next one.
If all accounts are exhausted, it waits and resumes automatically once
limits reset (new day) - no manual restart needed as long as the app
process is running.
"""
import os
import json
import time
import random
import smtplib
import threading
from pathlib import Path
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr

import pandas as pd
from flask import Flask, request, jsonify, render_template, send_file

APP_DIR = Path(__file__).parent
UPLOAD_DIR = APP_DIR / "uploads"
DATA_DIR = APP_DIR / "data"
UPLOAD_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

ACCOUNTS_FILE = DATA_DIR / "accounts.json"
USAGE_FILE = DATA_DIR / "usage.json"
PROGRESS_FILE = DATA_DIR / "progress.json"

app = Flask(__name__)
lock = threading.Lock()

state = {
    "running": False,
    "waiting": False,      # true when all accounts are out of quota, waiting for reset
    "total": 0,
    "sent": 0,
    "failed": 0,
    "current_account": "",
    "current_contact": "",
    "log": [],
    "done": False,
    "failed_csv": None,
    "stop": False,
}


def log_line(msg):
    with lock:
        state["log"].append(msg)
        state["log"] = state["log"][-300:]


# ── Accounts persistence ─────────────────────────────────────────────────
def load_accounts():
    if ACCOUNTS_FILE.exists():
        return json.loads(ACCOUNTS_FILE.read_text())
    return []


def save_accounts(accounts):
    ACCOUNTS_FILE.write_text(json.dumps(accounts, indent=2))


def load_usage():
    if USAGE_FILE.exists():
        return json.loads(USAGE_FILE.read_text())
    return {}


def save_usage(usage):
    USAGE_FILE.write_text(json.dumps(usage, indent=2))


def usage_today(usage, email):
    return usage.get(email, {}).get(date.today().isoformat(), 0)


def increment_usage(email):
    with lock:
        usage = load_usage()
        today = date.today().isoformat()
        usage.setdefault(email, {})
        usage[email][today] = usage[email].get(today, 0) + 1
        save_usage(usage)


def pick_account(accounts):
    """Return first account that still has quota left today, else None."""
    usage = load_usage()
    for acct in accounts:
        if usage_today(usage, acct["email"]) < acct.get("daily_limit", 500):
            return acct
    return None


# ── Progress persistence (so restarts don't resend) ─────────────────────
def load_progress():
    if PROGRESS_FILE.exists():
        d = json.loads(PROGRESS_FILE.read_text())
        return set(d.get("sent", [])), set(d.get("failed", []))
    return set(), set()


def save_progress(sent, failed):
    PROGRESS_FILE.write_text(json.dumps({"sent": list(sent), "failed": list(failed)}, indent=2))


# ── Contacts / email building ────────────────────────────────────────────
def load_contacts(filepath, email_col):
    path = Path(filepath)
    ext = path.suffix.lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(filepath, dtype=str)
    else:
        df = pd.read_csv(filepath, dtype=str)
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    df.columns = [c.strip() for c in df.columns]
    if email_col not in df.columns:
        raise ValueError(f"Column '{email_col}' not found. Available: {list(df.columns)}")
    df = df[df[email_col].notna() & (df[email_col] != "")]
    return df.to_dict("records")


def build_email(contact, resume_bytes, resume_filename, cfg, acct):
    name = contact.get(cfg["name_col"], "Hiring Manager") or "Hiring Manager"
    company = contact.get(cfg["company_col"], "your company") or "your company"
    title = contact.get(cfg["title_col"], "") or ""
    to_email = contact[cfg["email_col"]]

    subject = cfg["subject_template"].format(name=name, company=company, title=title)
    body = cfg["body_template"].format(name=name, company=company, title=title)

    sender_name = cfg.get("sender_name") or acct["email"]
    msg = MIMEMultipart()
    msg["From"] = formataddr((sender_name, acct["email"]))
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    part = MIMEBase("application", "octet-stream")
    part.set_payload(resume_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{resume_filename}"')
    msg.attach(part)
    return msg


def create_smtp_connection(acct, smtp_host, smtp_port):
    if smtp_port == 465:
        server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30)
    else:
        server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
        server.ehlo()
        server.starttls()
        server.ehlo()
    server.login(acct["email"], acct["password"])
    return server


# ── Main background job ──────────────────────────────────────────────────
def run_job(cfg, resume_path, contacts_path):
    try:
        accounts = load_accounts()
        if not accounts:
            log_line("No accounts configured. Add at least one Gmail account first.")
            return

        resume_bytes = Path(resume_path).read_bytes()
        resume_filename = Path(resume_path).name
        contacts = load_contacts(contacts_path, cfg["email_col"])

        sent, failed = load_progress()
        pending = [c for c in contacts if c[cfg["email_col"]] not in sent and c[cfg["email_col"]] not in failed]

        with lock:
            state["total"] = len(contacts)
            state["sent"] = len(sent)
            state["failed"] = len(failed)

        log_line(f"Loaded {len(contacts)} contacts. Pending: {len(pending)}. Accounts: {len(accounts)}.")

        server = None
        current_email = None

        for i, contact in enumerate(pending, 1):
            with lock:
                if state["stop"]:
                    log_line("Stopped by user.")
                    break

            # find an account with quota, waiting (and auto-resuming) if all are exhausted
            acct = pick_account(accounts)
            while acct is None:
                with lock:
                    state["waiting"] = True
                    if state["stop"]:
                        break
                log_line("All accounts hit their daily limit. Waiting for reset (checks every 5 min)...")
                time.sleep(300)
                with lock:
                    if state["stop"]:
                        break
                acct = pick_account(accounts)
            with lock:
                if state["stop"]:
                    break
                state["waiting"] = False

            to_email = contact[cfg["email_col"]]
            with lock:
                state["current_contact"] = to_email
                state["current_account"] = acct["email"]

            if acct["email"] != current_email:
                try:
                    if server:
                        server.quit()
                except Exception:
                    pass
                server = create_smtp_connection(acct, cfg["smtp_host"], cfg["smtp_port"])
                current_email = acct["email"]
                log_line(f"Using account: {acct['email']}")

            try:
                msg = build_email(contact, resume_bytes, resume_filename, cfg, acct)
                server.sendmail(acct["email"], [to_email], msg.as_string())
                sent.add(to_email)
                increment_usage(acct["email"])
                with lock:
                    state["sent"] += 1
                log_line(f"[{i}/{len(pending)}] ✓ {to_email}  (via {acct['email']})")
            except Exception as e:
                failed.add(to_email)
                with lock:
                    state["failed"] += 1
                log_line(f"[{i}/{len(pending)}] ✗ {to_email}: {e}")
                try:
                    server = create_smtp_connection(acct, cfg["smtp_host"], cfg["smtp_port"])
                except Exception:
                    pass

            if i % 10 == 0:
                save_progress(sent, failed)

            if i % cfg["batch_size"] == 0 and i < len(pending):
                log_line(f"Batch pause {cfg['batch_pause']}s...")
                time.sleep(cfg["batch_pause"])

            time.sleep(cfg["delay"] + random.uniform(0, cfg["jitter"]))

        save_progress(sent, failed)
        try:
            if server:
                server.quit()
        except Exception:
            pass

        if failed:
            fdf = pd.DataFrame([c for c in contacts if c[cfg["email_col"]] in failed])
            fpath = DATA_DIR / "failed_emails.csv"
            fdf.to_csv(fpath, index=False)
            with lock:
                state["failed_csv"] = str(fpath)

        log_line("DONE.")
    except Exception as e:
        log_line(f"FATAL ERROR: {e}")
    finally:
        with lock:
            state["running"] = False
            state["waiting"] = False
            state["done"] = True


# ── Routes: accounts ──────────────────────────────────────────────────────
@app.route("/accounts", methods=["GET"])
def get_accounts():
    accounts = load_accounts()
    usage = load_usage()
    out = []
    for a in accounts:
        out.append({
            "email": a["email"],
            "daily_limit": a.get("daily_limit", 500),
            "used_today": usage_today(usage, a["email"]),
        })
    return jsonify(out)


@app.route("/accounts/add", methods=["POST"])
def add_account():
    data = request.json
    email = data.get("email", "").strip()
    password = data.get("password", "").strip()
    daily_limit = int(data.get("daily_limit", 500))
    if not email or not password:
        return jsonify({"error": "Email and app password required."}), 400
    accounts = load_accounts()
    if any(a["email"] == email for a in accounts):
        return jsonify({"error": "Account already added."}), 400
    accounts.append({"email": email, "password": password, "daily_limit": daily_limit})
    save_accounts(accounts)
    return jsonify({"ok": True})


@app.route("/accounts/remove", methods=["POST"])
def remove_account():
    email = request.json.get("email")
    accounts = [a for a in load_accounts() if a["email"] != email]
    save_accounts(accounts)
    return jsonify({"ok": True})


# ── Routes: campaign ──────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start():
    with lock:
        if state["running"]:
            return jsonify({"error": "A campaign is already running."}), 400

    resume = request.files.get("resume")
    contacts = request.files.get("contacts")
    if not resume or not contacts:
        return jsonify({"error": "Resume and contacts file are required."}), 400

    resume_path = UPLOAD_DIR / resume.filename
    contacts_path = UPLOAD_DIR / contacts.filename
    resume.save(resume_path)
    contacts.save(contacts_path)

    cfg = {
        "smtp_host": request.form.get("smtp_host", "smtp.gmail.com"),
        "smtp_port": int(request.form.get("smtp_port", 587)),
        "sender_name": request.form.get("sender_name", ""),
        "subject_template": request.form.get("subject_template", "Application for {title} at {company}"),
        "body_template": request.form.get("body_template", "Dear {name},\n\nI am interested in opportunities at {company}."),
        "email_col": request.form.get("email_col", "Email"),
        "name_col": request.form.get("name_col", "Name"),
        "company_col": request.form.get("company_col", "Company"),
        "title_col": request.form.get("title_col", "Title"),
        "delay": float(request.form.get("delay", 3)),
        "jitter": float(request.form.get("jitter", 2)),
        "batch_size": int(request.form.get("batch_size", 40)),
        "batch_pause": float(request.form.get("batch_pause", 60)),
    }

    with lock:
        state.update({
            "running": True, "waiting": False, "total": 0, "sent": 0, "failed": 0,
            "current_account": "", "current_contact": "", "log": [], "done": False,
            "failed_csv": None, "stop": False,
        })

    t = threading.Thread(target=run_job, args=(cfg, resume_path, contacts_path), daemon=True)
    t.start()
    return jsonify({"ok": True})


@app.route("/stop", methods=["POST"])
def stop():
    with lock:
        state["stop"] = True
    return jsonify({"ok": True})


@app.route("/status")
def status():
    with lock:
        s = dict(state)
    accounts = load_accounts()
    usage = load_usage()
    s["accounts"] = [{
        "email": a["email"],
        "daily_limit": a.get("daily_limit", 500),
        "used_today": usage_today(usage, a["email"]),
    } for a in accounts]
    return jsonify(s)


@app.route("/download-failed")
def download_failed():
    with lock:
        path = state["failed_csv"]
    if not path or not os.path.exists(path):
        return jsonify({"error": "No failed emails file available."}), 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000, threaded=True)
