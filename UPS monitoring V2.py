#-------------------------------------------------------------------------------
# Name:        UPS(APC Power Chute) monitoring with E-mail notifycation
# Purpose:
#
# Author:      9A4AM
#
# Created:     07.10.2025
# Copyright:   (c) 9A4AM 2025
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import time
import wmi
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
import os
import sys

# ------------------------------
# Gmail SMTP settings
# ------------------------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_USER = "your_email@gmail.com"
EMAIL_PASS = "your_password"  # 16-znamenkasti App Password
TO_EMAIL = "your_email@gmail.com"

# ------------------------------
# Log file location (folder aplikacije)
# ------------------------------
if getattr(sys, 'frozen', False):
    # Ako je EXE
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    # Ako je .py
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_FILE = os.path.join(SCRIPT_DIR, "UPS_log.txt")

# ------------------------------
# Delay before sending email when UPS goes on battery (in seconds)
# ------------------------------
BATTERY_DELAY = 20

# ------------------------------
# Helper functions
# ------------------------------
def log_event(text):
    """Append event with timestamp to log file"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"{timestamp} - {text}\n"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(entry)
    print(entry.strip())

def send_email(subject, body):
    """Send an email using Gmail SMTP with App Password"""
    msg = MIMEText(body)
    msg["From"] = EMAIL_USER
    msg["To"] = TO_EMAIL
    msg["Subject"] = subject
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
        log_event(f"📧 Email poslan: {subject}")
    except Exception as e:
        log_event(f"❌ Greška prilikom slanja emaila: {e}")

def get_ups_status():
    """Return UPS status via WMI (1 = battery, 2 = AC)"""
    c = wmi.WMI(namespace="root\\CIMV2")
    for battery in c.Win32_Battery():
        return battery.BatteryStatus
    return None

# ------------------------------
# Main loop
# ------------------------------
on_battery_sent = False
battery_start_time = None

log_event("⚙️ UPS monitoring pokrenut.")

while True:
    status = get_ups_status()

    # UPS switched to battery
    if status == 1 and not on_battery_sent:
        if battery_start_time is None:
            battery_start_time = time.time()
            log_event("⚠️ UPS na bateriji, čekam 20s potvrdu...")
        elif time.time() - battery_start_time >= BATTERY_DELAY:
            send_email("⚠️ UPS na bateriji", "UPS je prešao na baterijsko napajanje (nakon 20s).")
            on_battery_sent = True
            battery_start_time = None

    # UPS back on AC power
    elif status == 2 and on_battery_sent:
        send_email("✅ Struja vraćena", "UPS je ponovno na mrežnom napajanju.")
        on_battery_sent = False
        battery_start_time = None

    # UPS se vratio prije 20s, reset timer
    elif status == 2 and battery_start_time is not None:
        log_event("ℹ️ Struja se vratila prije isteka 20s.")
        battery_start_time = None

    time.sleep(5)

