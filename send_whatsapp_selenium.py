"""
WhatsApp Bulk Message Sender — Selenium Edition
================================================
Opens WhatsApp Web ONCE in a single Chrome tab and sends all messages
from there — no repeated tab opening/closing.

Requirements:
    pip install pandas openpyxl python-docx selenium webdriver-manager

Usage:
    python send_whatsapp_selenium.py

How it works:
    1. Opens Chrome → web.whatsapp.com
    2. Shows QR code — scan with your phone (first run only; session is saved)
    3. Sends each personalised message one by one in the same tab
    4. Logs results to whatsapp_selenium_log.txt
"""

import time
import re
import sys
import logging
import urllib.parse
from pathlib import Path

# ── third-party ───────────────────────────────────────────────────────────────
try:
    import pandas as pd
    from docx import Document
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError as e:
    print(f"\n[ERROR] Missing library: {e}")
    print("Run:  pip install pandas openpyxl python-docx selenium webdriver-manager\n")
    sys.exit(1)

# ═══════════════════════ CONFIGURATION ═══════════════════════════════════════

EXCEL_FILE    = r"/Users/sanjeevgarg/Videos Using Claude/ca-lead-capture-demo-booking/CA_Delhi_Contacts V2.0.xlsx"
WORD_FILE     = r"/Users/sanjeevgarg/Videos Using Claude/ca-lead-capture-demo-booking/Message Draft.docx"
LOG_FILE      = r"/Users/sanjeevgarg/Videos Using Claude/ca-lead-capture-demo-booking/whatsapp_selenium_log.txt"

# Session profile — saves your WhatsApp Web login so you only scan QR once
CHROME_PROFILE = str(Path.home() / "whatsapp_chrome_profile")

COUNTRY_CODE   = "+91"   # India — change if needed (e.g. "+1" for USA)

# Seconds to wait for each chat to open. Increase on slow internet.
PAGE_LOAD_WAIT = 30

# Seconds to pause after hitting Send before moving to next contact
SEND_WAIT      = 3

# Skip the first N rows (useful for resuming after a partial send)
# e.g. set to 33 to skip contacts 1–33 and start from contact 34
START_FROM_ROW = 0

# Set True to print messages without sending (safe test)
DRY_RUN        = False

# ═════════════════════════════════════════════════════════════════════════════


def setup_logging(log_path: str) -> logging.Logger:
    logger = logging.getLogger("wa_selenium")
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
    df = pd.read_excel(excel_path, dtype=str)
    df.columns = df.columns.str.strip()
    name_col   = next(c for c in df.columns if "name"   in c.lower())
    mobile_col = next(c for c in df.columns if "mobile" in c.lower()
                      or "phone" in c.lower() or "number" in c.lower())
    df = df[[name_col, mobile_col]].rename(
        columns={name_col: "Name", mobile_col: "Mobile"})
    df.dropna(inplace=True)
    df["Name"]   = df["Name"].str.strip()
    df["Mobile"] = df["Mobile"].str.strip()
    return df.reset_index(drop=True)


def read_template(word_path: str) -> str:
    doc = Document(word_path)
    return "\n".join(p.text for p in doc.paragraphs)


def personalise(template: str, name: str) -> str:
    return re.sub(r"\{\{\s*[Nn]ame\s*\}\}", name, template)


def format_number(raw: str, country_code: str) -> str:
    digits    = re.sub(r"\D", "", raw)
    cc_digits = re.sub(r"\D", "", country_code)
    if digits.startswith(cc_digits):
        return "+" + digits
    return country_code + digits


def launch_browser(profile_dir: str) -> webdriver.Chrome:
    """Launch Chrome with a persistent profile so QR login is remembered."""
    opts = Options()
    opts.add_argument(f"--user-data-dir={profile_dir}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    # Keep the window visible so you can scan the QR code
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=opts)
    driver.maximize_window()
    return driver


def wait_for_whatsapp(driver: webdriver.Chrome, timeout: int,
                      logger: logging.Logger) -> bool:
    """
    Poll multiple selectors until WhatsApp Web is ready.
    Gives up to `timeout` seconds — enough time to scan the QR code.
    """
    logger.info(f"Waiting up to {timeout}s for WhatsApp Web "
                f"(scan QR code if prompted) …")
    selectors = [
        (By.XPATH,        '//div[@contenteditable="true"][@data-tab="3"]'),
        (By.XPATH,        '//div[@data-tab="3"]'),
        (By.CSS_SELECTOR, 'div[data-tab="3"]'),
        (By.XPATH,        '//div[@aria-label="Search input textbox"]'),
        (By.XPATH,        '//*[@id="side"]'),
        (By.CSS_SELECTOR, '#side'),
    ]
    deadline = time.time() + timeout
    while time.time() < deadline:
        for by, sel in selectors:
            try:
                el = driver.find_element(by, sel)
                if el.is_displayed():
                    logger.info("✓ WhatsApp Web is ready")
                    return True
            except Exception:
                pass
        time.sleep(2)
    return False


def send_message_selenium(driver: webdriver.Chrome, phone: str, message: str,
                          page_wait: int, send_wait: int,
                          logger: logging.Logger):
    """
    Navigate to the chat for `phone` via wa.me URL with pre-filled text,
    then hit Enter to send — all in the same tab.
    """
    url = (f"https://web.whatsapp.com/send"
           f"?phone={phone}&text={urllib.parse.quote(message)}")
    driver.get(url)

    # Try multiple XPaths for the message input box
    box = None
    for xpath in [
        '//div[@contenteditable="true"][@data-tab="10"]',
        '//div[@data-tab="10"]',
        '//footer//div[@contenteditable="true"]',
    ]:
        try:
            box = WebDriverWait(driver, page_wait).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            if box:
                break
        except Exception:
            pass

    if not box:
        raise RuntimeError("Message input box not found — "
                           "number may not be on WhatsApp")

    time.sleep(1.5)
    box.send_keys(Keys.ENTER)
    time.sleep(send_wait)


def main():
    logger = setup_logging(LOG_FILE)
    logger.info("=" * 60)
    logger.info("WhatsApp Selenium Sender — started")
    logger.info(f"Excel      : {EXCEL_FILE}")
    logger.info(f"Word       : {WORD_FILE}")
    logger.info(f"Start row  : {START_FROM_ROW + 1}")
    logger.info(f"Dry run    : {DRY_RUN}")
    logger.info("=" * 60)

    contacts = read_contacts(EXCEL_FILE)
    logger.info(f"Total contacts loaded: {len(contacts)}")

    template = read_template(WORD_FILE)
    logger.info(f"Template: {template[:60].strip()} …")

    remaining = contacts.iloc[START_FROM_ROW:]
    logger.info(f"Sending to {len(remaining)} contacts "
                f"(rows {START_FROM_ROW + 1}–{len(contacts)})")

    # ── Dry run ───────────────────────────────────────────────────────────────
    if DRY_RUN:
        logger.info("DRY RUN — printing messages only, not sending.")
        for idx, row in remaining.iterrows():
            msg   = personalise(template, row["Name"])
            phone = format_number(row["Mobile"], COUNTRY_CODE)
            logger.info(f"[{idx+1:>3}] {row['Name']:<30} | {phone}")
            logger.debug(f"        {msg[:80]} …")
        logger.info("Dry run complete.")
        return

    # ── Launch browser & wait for login ──────────────────────────────────────
    logger.info("Launching Chrome …")
    driver = launch_browser(CHROME_PROFILE)
    driver.get("https://web.whatsapp.com")

    if not wait_for_whatsapp(driver, timeout=180, logger=logger):
        logger.error("Timed out — WhatsApp Web did not load. "
                     "Please scan QR and retry.")
        driver.quit()
        sys.exit(1)

    time.sleep(3)   # let the UI fully settle after login

    # ── Send loop ─────────────────────────────────────────────────────────────
    sent = 0; failed = 0; skipped = 0

    for idx, row in remaining.iterrows():
        name   = str(row["Name"]).strip()
        mobile = str(row["Mobile"]).strip()
        phone  = format_number(mobile, COUNTRY_CODE)

        if len(re.sub(r"\D", "", phone)) < 10:
            logger.warning(f"[{idx+1:>3}] SKIP  {name:<30} | invalid: {mobile}")
            skipped += 1
            continue

        message = personalise(template, name)
        logger.info(f"[{idx+1:>3}] SEND  {name:<30} | {phone}")

        try:
            send_message_selenium(driver, phone, message,
                                  PAGE_LOAD_WAIT, SEND_WAIT, logger)
            logger.info(f"       ✓ Sent")
            sent += 1
        except Exception as e:
            logger.error(f"       ✗ Failed: {e}")
            failed += 1

        time.sleep(2)   # brief pause between contacts

    # ── Done ──────────────────────────────────────────────────────────────────
    driver.quit()
    logger.info("=" * 60)
    logger.info(f"DONE — Sent: {sent}  |  Failed: {failed}  |  Skipped: {skipped}")
    logger.info(f"Log: {LOG_FILE}")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
