import argparse
import csv
import json
import logging
import os
import smtplib
import time
import threading
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger
from pymongo import MongoClient
import pandas as pd
import re
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

# Logging Configuration
logging.basicConfig(
    filename="email_automation.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Persistent Task Storage File
TASK_FILE = "email_tasks.json"

# Initialize Scheduler
scheduler = BackgroundScheduler()

# MongoDB Configuration
MONGO_URI = "mongodb://localhost:27017/"
client = MongoClient(MONGO_URI)
db = client["email_automation_db"]
logs_collection = db["email_logs"]

# Gmail SMTP Configuration
GMAIL_SMTP_SERVER = "smtp.gmail.com"
GMAIL_SMTP_PORT = 587

def log_to_mongodb(task_name, details, status, level="INFO"):
    log_entry = {
        "task_name": task_name,
        "details": details,
        "status": status,
        "level": level,
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    logs_collection.insert_one(log_entry)

def load_tasks():
    try:
        with open(TASK_FILE, "r") as f:
            tasks = json.load(f)
            return tasks if isinstance(tasks, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_tasks(tasks):
    with open(TASK_FILE, "w") as f:
        json.dump(tasks, f, indent=4)

def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

def send_email(recipient_email, subject, message, attachments=None):
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        logging.error("Missing email credentials in .env file.")
        return False

    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg["Subject"] = subject

    msg.attach(MIMEText(message, "plain"))

    if attachments:
        for attachment in attachments:
            try:
                with open(attachment, "rb") as file:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment)}")
                msg.attach(part)
            except FileNotFoundError:
                logging.error(f"Attachment '{attachment}' not found.")

    try:
        server = smtplib.SMTP(GMAIL_SMTP_SERVER, GMAIL_SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        text = msg.as_string()
        server.sendmail(SENDER_EMAIL, recipient_email, text)
        server.quit()
        logging.info(f"Email sent to {recipient_email}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email to {recipient_email}: {e}")
        return False

def email_task(task_name, email_list_path, message_path, subject, attachments):
    try:
        if email_list_path.endswith(".csv"):
            df = pd.read_csv(email_list_path)
        elif email_list_path.endswith(".xlsx"):
            df = pd.read_excel(email_list_path)
        else:
            raise ValueError("Unsupported email list format. Only CSV and XLSX are supported.")

        invalid_emails = []
        with open(message_path, "r") as f:
            message_template = f.read()

        for index, row in df.iterrows():
            email = row.get("email")
            name = row.get("name", "")

            if not is_valid_email(email):
                invalid_emails.append(email)
                continue

            message = message_template.replace("{name}", name)
            if send_email(email, subject, message, attachments):
                log_to_mongodb(task_name, {"recipient": email, "subject": subject}, "Email sent")
            else:
                log_to_mongodb(task_name, {"recipient": email, "subject": subject}, "Email failed", level="ERROR")

        if invalid_emails:
            with open("invalid_emails.log", "a") as f:
                for email in invalid_emails:
                    f.write(f"{email}\n")
            logging.warning(f"Invalid emails found: {', '.join(invalid_emails)}")
            log_to_mongodb(task_name, {"invalid_emails": invalid_emails}, "Invalid emails", level="WARNING")

    except Exception as e:
        logging.error(f"Task '{task_name}' failed: {e}")
        log_to_mongodb(task_name, {"email_list": email_list_path, "message_file": message_path, "subject": subject, "attachments": attachments}, f"Error: {e}", level="ERROR")

def add_task(interval, unit, email_list, message_file, subject, attachments):
    tasks = load_tasks()
    task_name = f"task_{len(tasks) + 1}"

    new_task_details = {
        "interval": interval,
        "unit": unit,
        "email_list": email_list,
        "message_file": message_file,
        "subject": subject,
        "attachments": attachments
    }

    # Check for duplicates
    for existing_task_details in tasks.values():
        if new_task_details == existing_task_details:
            print("⚠️ Task with the same interval and details already exists.")
            return

    tasks[task_name] = new_task_details
    save_tasks(tasks)

    trigger = IntervalTrigger(**{unit: interval})
    scheduler.add_job(email_task, trigger, args=[task_name, email_list, message_file, subject, attachments], id=task_name)

    print(f"✅ Task '{task_name}' added successfully.")

def list_tasks():
    tasks = load_tasks()
    if not tasks:
        print("⚠️ No scheduled tasks found.")
        return

    print("\n📌 Scheduled Email Tasks:")
    for task_name, details in tasks.items():
        print(f"🔹 {task_name} - Every {details['interval']} {details['unit']}")

def remove_task(task_name):
    tasks = load_tasks()
    if task_name not in tasks:
        print(f"⚠️ Task '{task_name}' not found.")
        return

    del tasks[task_name]
    save_tasks(tasks)

    try:
        scheduler.remove_job(task_name)
        print(f"✅ Task '{task_name}' removed successfully.")
    except Exception:
        print(f"⚠️ Task '{task_name}' was not running but removed from saved tasks.")

# CLI Argument Parsing
parser = argparse.ArgumentParser(description="Email Automation Scheduler")
parser.add_argument("--add", type=int, help="Add a new task with interval")
parser.add_argument("--unit", type=str, choices=["seconds", "minutes", "hours", "days"], help="Time unit for the interval")
parser.add_argument("--email-list", type=str, help="Path to email list file (CSV/XLSX)")
parser.add_argument("--message-file", type=str, help="Path to message file (TXT)")
parser.add_argument("--subject", type=str, help="Email subject")
parser.add_argument("--attachments", nargs="*", default=[], help="List of attachment file paths")
parser.add_argument("--list", action="store_true", help="List all scheduled tasks")
parser.add_argument("--remove", type=str, help="Remove a scheduled task by name")

args = parser.parse_args()

if args.add:
    if not all((args.unit, args.email_list, args.message_file, args.subject)):
        print("⚠️ Please provide --unit, --email-list, --message-file, and --subject.")
        exit(1)
    add_task(args.add, args.unit, args.email_list, args.message_file, args.subject, args.attachments)
elif args.list:
    list_tasks()
elif args.remove:
    remove_task(args.remove)


def load_and_schedule_tasks():
    tasks = load_tasks()
    for task_name, details in tasks.items():
        trigger = IntervalTrigger(**{details["unit"]: details["interval"]})
        scheduler.add_job(
            email_task,
            trigger,
            args=[
                task_name,
                details["email_list"],
                details["message_file"],
                details["subject"],
                details["attachments"],
            ],
            id=task_name,
        )
# Load and schedule tasks before parsing commands
load_and_schedule_tasks()
def start_scheduler():
    """Runs the scheduler in a separate thread."""
    scheduler.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("🛑 Scheduler stopped.")
        scheduler.shutdown()

# Start the scheduler thread only if no other command is given
if not (args.add or args.list or args.remove):
    scheduler_thread = threading.Thread(target=start_scheduler, daemon=True)
    scheduler_thread.start()

    # Keep the main thread alive
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("🛑 Scheduler stopped.")
        scheduler.shutdown()

