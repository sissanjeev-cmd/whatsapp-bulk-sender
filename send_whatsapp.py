"""
WhatsApp Bulk Message Sender
============================
Reads contacts from Excel, personalizes message from Word template,
and sends via WhatsApp Web using pywhatkit.

Requirements:
    pip install pandas openpyxl python-docx pywhatkit

Usage:
    python send_whatsapp.py

Notes:
    - WhatsApp Web must be logged in (it opens automatically on first run)
    - Keep your phone connected / WhatsApp active
    - Numbers are treated as Indian (+91). Edit COUNTRY_CODE below if needed.
"""

import time
import re
import sys
import logging
from datetime import datetime
from pathlib import Path

# ── third-party (install once with pip) ──────────────────────────────────────
try:
    import pandas as pd
    from docx import Document
    import pywhatkit as kit
except ImportError as e:
    print(f"\n[ERROR] Missing library: {e}")
    print("Run:  pip install pandas openpyxl python-docx pywhatkit\n")
    sys.exit(1)

# ═══════════════════════ CONFIGURATION ═══════════════════════════════════════

EXCEL_FILE   = r"/Users/sanjeevgarg/Videos Using Claude/ca-lead-capture-demo-booking/CA_Delhi_Contacts V2.0.xlsx"
WORD_FILE    = r"/Users/sanjeevgarg/Videos Using Claude/ca-lead-capture-demo-booking/Message Draft.docx"
LOG_FILE     = r"/Users/sanjeevgarg/Videos Using Claude/ca-lead-capture-demo-booking/whatsapp_send_log.txt"

COUNTRY_CODE = "+91"          # India — change if needed (e.g. "+1" for USA)

# Seconds to wait after opening each WhatsApp Web tab before closing it.
# Increase if your internet is slow.
WAIT_TIME    = 20             # seconds per message

# Set to True to do a DRY RUN (print messages without actually sending)
DRY_RUN      = False

# ═════════════════════════════════════════════════════════════════════════════


def setup_logging(log_path: str) -> logging.Logger:
    """Configure logger to write to both console and a log file."""
    logger = logging.getLogger("wa_sender")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")

    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    return logger


def read_contacts(excel_path: str) -> pd.DataFrame:
    """Load Excel contacts, keeping only rows with a valid mobile number."""
    df = pd.read_excel(excel_path, dtype=str)
    df.columns = df.columns.str.strip()

    # Flexible column detection
    name_col   = next((c for c in df.columns if "name"   in c.lower()), None)
    mobile_col = next((c for c in df.columns if "mobile" in c.lower()
                        or "phone" in c.lower()
                        or "number" in c.lower()), None)

    if not name_col or not mobile_col:
        raise ValueError(
            f"Could not find Name/Mobile columns. Found: {list(df.columns)}"
        )

    df = df[[name_col, mobile_col]].rename(
        columns={name_col: "Name", mobile_col: "Mobile"}
    )
    df.dropna(subset=["Mobile"], inplace=True)
    df["Mobile"] = df["Mobile"].str.strip()
    df.dropna(subset=["Name"], inplace=True)
    df["Name"] = df["Name"].str.strip()
    return df.reset_index(drop=True)


def read_template(word_path: str) -> str:
    """Extract full text from a Word document, preserving line breaks."""
    doc = Document(word_path)
    lines = []
    for para in doc.paragraphs:
        lines.append(para.text)
    return "\n".join(lines)


def personalise(template: str, name: str) -> str:
    """Replace all {{Name}} (case-insensitive) placeholders with the contact name."""
    return re.sub(r"\{\{\s*[Nn]ame\s*\}\}", name, template)


def format_number(raw: str, country_code: str) -> str:
    """
    Ensure the number has the country code prefix and contains only digits after +.
    E.g.  '9212007566' → '+919212007566'
         '+919212007566' → '+919212007566'
    """
    digits = re.sub(r"\D", "", raw)        # strip spaces, dashes, etc.
    cc_digits = re.sub(r"\D", "", country_code)

    if digits.startswith(cc_digits):
        return "+" + digits
    return country_code + digits


def send_message(phone: str, message: str, wait: int, logger: logging.Logger):
    """Send a single WhatsApp message via pywhatkit (instant mode)."""
    # pywhatkit.sendwhatmsg_instantly opens WhatsApp Web and sends immediately.
    kit.sendwhatmsg_instantly(
        phone_no   = phone,
        message    = message,
        wait_time  = wait,      # seconds to wait for WhatsApp Web to load
        tab_close  = True,      # close the tab after sending
        close_time = 3          # seconds after sending before tab closes
    )


def main():
    logger = setup_logging(LOG_FILE)
    logger.info("=" * 60)
    logger.info("WhatsApp Bulk Sender — started")
    logger.info(f"Excel  : {EXCEL_FILE}")
    logger.info(f"Word   : {WORD_FILE}")
    logger.info(f"Dry run: {DRY_RUN}")
    logger.info("=" * 60)

    # ── Load contacts ─────────────────────────────────────────────────────────
    logger.info("Reading contacts …")
    contacts = read_contacts(EXCEL_FILE)
    logger.info(f"  → {len(contacts)} contacts loaded")

    # ── Load message template ─────────────────────────────────────────────────
    logger.info("Reading message template …")
    template = read_template(WORD_FILE)
    logger.info(f"  → Template preview: {template[:80].strip()} …")

    # ── Send loop ─────────────────────────────────────────────────────────────
    sent = 0; failed = 0; skipped = 0

    for idx, row in contacts.iterrows():
        name   = str(row["Name"]).strip()
        mobile = str(row["Mobile"]).strip()
        phone  = format_number(mobile, COUNTRY_CODE)

        # Basic sanity check — must have at least 10 digits after country code
        if len(re.sub(r"\D", "", phone)) < 10:
            logger.warning(f"[{idx+1:>3}] SKIP  {name:<30} | invalid number: {mobile}")
            skipped += 1
            continue

        message = personalise(template, name)

        if DRY_RUN:
            logger.info(f"[{idx+1:>3}] DRY   {name:<30} | {phone}")
            logger.debug(f"          Message preview:\n{message[:120]} …\n")
            sent += 1
            continue

        logger.info(f"[{idx+1:>3}] SEND  {name:<30} | {phone}")
        try:
            send_message(phone, message, WAIT_TIME, logger)
            logger.info(f"       ✓ Sent successfully")
            sent += 1
        except Exception as e:
            logger.error(f"       ✗ Failed: {e}")
            failed += 1

        # Brief pause between messages to avoid rate-limiting
        time.sleep(5)

    # ── Summary ───────────────────────────────────────────────────────────────
    logger.info("=" * 60)
    logger.info(f"DONE — Sent: {sent}  |  Failed: {failed}  |  Skipped: {skipped}")
    logger.info(f"Full log saved to: {LOG_FILE}")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
