import os
import json
import time
import imaplib
import email
import schedule
import logging
import io
from email.header import decode_header
from datetime import datetime
from anthropic import Anthropic
import gspread
from google.oauth2.service_account import Credentials
import resend

ZOHO_EMAIL        = os.environ.get("ZOHO_EMAIL", "internal@scope-iq.io")
ZOHO_APP_PASSWORD = os.environ.get("ZOHO_APP_PASSWORD")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
RESEND_API_KEY    = os.environ.get("RESEND_API_KEY")
GOOGLE_SHEET_ID   = os.environ.get("GOOGLE_SHEET_ID", "1_4L63VqFN6etwLRWW2zT0WyJTHUT8LLTWwqiNPJJR4w")
GOOGLE_CREDS_FILE = "primordial-mile-495807-k9-0217981265dd.json"
MODEL             = "claude-haiku-4-5-20251001"

IMAP_HOST = "imappro.zoho.com"
IMAP_PORT = 993

SCOPE_TEAM_EMAILS = [
    e.strip().lower() for e in
    os.environ.get("SCOPE_TEAM_EMAILS", "").split(",")
    if e.strip()
]

REPORT_RECIPIENTS = [
    "alishir.aliyev@scopeconsulting.az",
    "afgan.mammadov@scopeconsulting.az",
    "techoffice@scopeconsulting.az"
]

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

anthropic_client     = Anthropic(api_key=ANTHROPIC_API_KEY)
_processed_ids_cache = set()
_cache_loaded        = False

# In-memory file cache — sender email -> list of files
_file_memory_cache = {}

SYSTEM_PROMPT = """You are Alex Rivera, Construction Expert at SCOPE Consulting MMC.

You have 25 years of experience across commercial, residential, infrastructure, oil and gas, and high-end fit-out projects in Europe, the Middle East, and CIS countries including Azerbaijan.

Your areas of expertise include Quantity Surveying, Cost Management, Quality Assurance and Quality Control, Contract Administration under FIDIC and NEC4, technical document review across all disciplines, materials and specifications, handover and commissioning, claims and variations, and all major construction standards including Eurocodes, British Standards, ISO, GOST, AzDTN and SNiP.

LANGUAGE RULES:
Default language is English. Always reply in English unless the incoming email is written entirely in Azerbaijani.
If the email is in Azerbaijani, reply in Azerbaijani using correct formal register and proper special characters: ə ı ö ü ğ ş ç.
Never mix languages in one reply.

EMAIL WRITING STYLE - CRITICAL:
Write exactly like a senior construction professional writing a formal business email.
Use plain professional prose only. No bullet points, no symbols, no emojis, no checkmarks, no arrows, no dashes used as decoration, no hashtags.
Paragraphs separated by blank lines.
Begin with a formal salutation such as Dear Alishir, or Dear Team, as appropriate.
End with a formal closing such as Kind regards or Yours sincerely.
Be direct, confident and senior in tone.
Keep responses concise and actionable.

RESPONSE TIME - CRITICAL:
Never say you will review later or that you need more time.
Never say within the working day, within two working days, or any future time commitment.
Always analyse immediately and include full findings in your reply right now.
Never ask for files that were already sent previously. If files were previously received from this sender, use them to answer the question immediately.
Only request files if you have genuinely never received any from this sender.

DOCUMENT ANALYSIS - CRITICAL:
When any BOQ, Smeta, Excel, cost schedule, or pricing document is provided analyse every single sheet and every single position immediately.
Do not skip any sheet. Do not skip mezzanine, fit-out, MEP, civil, structural, external works, or any discipline.
For every position state whether the rate is within market range, above market, or below market and give the Baku market range.
Flag all missing items and scope gaps across all sheets.
Give total risk exposure at the end.

SIGNATURE — always end every email with exactly this:

Alex Rivera
Construction Expert
SCOPE Consulting MMC
internal@scope-iq.io

BAKU MARKET RATES:
Mechanical excavation 8 to 15 AZN per cubic metre, Manual excavation 60 to 90 AZN per cubic metre, Concrete C25/30 190 to 240 AZN per cubic metre, Reinforcement 1200 to 1500 AZN per tonne, Formwork 18 to 28 AZN per square metre, External brickwork 25 to 40 AZN per square metre, Internal brickwork 20 to 32 AZN per square metre, Plastering 16 to 24 AZN per square metre, Paint 12 to 18 AZN per square metre, Ceramic tiles 25 to 40 AZN per square metre, Premium tiles 45 to 80 AZN per square metre, Gypsum partition 32 to 48 AZN per square metre, Armstrong ceiling 28 to 42 AZN per square metre, Raised access floor 55 to 85 AZN per square metre, Epoxy floor 35 to 55 AZN per square metre, Carpet tiles 40 to 70 AZN per square metre, Aluminium glazing 180 to 280 AZN per square metre, Timber door 350 to 550 AZN each, Fire door 600 to 1200 AZN each, Metal door 4500 to 6500 AZN each, Aluminium door 700 to 900 AZN each, HVAC ductwork 45 to 75 AZN per square metre, Fan coil unit 350 to 600 AZN each, Chiller 120 to 200 AZN per kW, Plumbing pipework 25 to 55 AZN per metre, Sanitary point 180 to 350 AZN per point, Cable tray 35 to 65 AZN per metre, LV cable 8 to 25 AZN per metre, Distribution board 800 to 3500 AZN each, Lighting fixture 45 to 120 AZN each, Emergency lighting 80 to 180 AZN each, Fire alarm 35 to 65 AZN per square metre, Sprinkler system 45 to 75 AZN per square metre, Access control 800 to 2500 AZN per door, CCTV 250 to 600 AZN per camera, Concrete paving 35 to 55 AZN per square metre, Natural stone paving 85 to 150 AZN per square metre, Soft landscape 25 to 55 AZN per square metre, Site fencing 45 to 120 AZN per metre, Passenger lift 45000 to 80000 AZN each, Freight lift 60000 to 120000 AZN each."""


def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds = Credentials.from_service_account_info(
            json.loads(creds_json), scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(
            GOOGLE_CREDS_FILE, scopes=scopes)
    return gspread.authorize(creds)


def get_sheet(tab="Sheet1"):
    try:
        client = get_gspread_client()
        return client.open_by_key(GOOGLE_SHEET_ID).worksheet(tab)
    except Exception as e:
        logger.error(f"Sheet error: {e}")
        return None


def get_or_create_sheet(title, rows=1000, cols=4):
    """Get existing sheet tab or create it"""
    try:
        client      = get_gspread_client()
        spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
        try:
            return spreadsheet.worksheet(title)
        except:
            sheet = spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
            logger.info(f"Created sheet tab: {title}")
            return sheet
    except Exception as e:
        logger.error(f"Get/create sheet error ({title}): {e}")
        return None


def save_files_to_memory(sender, attachments):
    """
    Save attachment content to Google Sheet AND in-memory cache.
    Replaces all previous files from same sender.
    """
    global _file_memory_cache
    sender = sender.strip().lower()

    # Save to in-memory cache immediately
    _file_memory_cache[sender] = [
        {
            "name":     att["name"],
            "content":  att["content"],
            "saved_at": datetime.now().strftime("%d.%m.%Y %H:%M")
        }
        for att in attachments
    ]
    logger.info(f"Saved {len(attachments)} files to in-memory cache for {sender}")

    # Also persist to Google Sheet
    try:
        sheet = get_or_create_sheet("File Memory", rows=2000, cols=4)
        if not sheet:
            return

        # Initialize header if empty
        all_values = sheet.get_all_values()
        if not all_values:
            sheet.append_row(["Sender", "Filename", "Content", "Saved At"])
            all_values = [["Sender", "Filename", "Content", "Saved At"]]

        # Delete existing rows for this sender
        rows_to_delete = []
        for i, row in enumerate(all_values):
            if i == 0:
                continue
            if row and len(row) > 0 and row[0].strip().lower() == sender:
                rows_to_delete.append(i + 1)

        for row_num in reversed(rows_to_delete):
            sheet.delete_rows(row_num)

        # Append new files
        saved_at = datetime.now().strftime("%d.%m.%Y %H:%M")
        for att in attachments:
            sheet.append_row([
                sender,
                att["name"],
                att["content"][:10000],
                saved_at
            ])
        logger.info(f"Saved {len(attachments)} files to sheet for {sender}")

    except Exception as e:
        logger.error(f"Save files to sheet error: {e}")


def load_files_from_memory(sender):
    """
    Load files for sender — check in-memory cache first, then Google Sheet.
    """
    global _file_memory_cache
    sender = sender.strip().lower()

    # Check in-memory cache first (fastest)
    if sender in _file_memory_cache and _file_memory_cache[sender]:
        files = _file_memory_cache[sender]
        logger.info(f"Loaded {len(files)} files from in-memory cache for {sender}")
        return files

    # Fall back to Google Sheet
    try:
        sheet = get_or_create_sheet("File Memory", rows=2000, cols=4)
        if not sheet:
            return []

        all_values = sheet.get_all_values()
        files = []
        for row in all_values[1:]:
            if row and len(row) >= 3 and row[0].strip().lower() == sender:
                files.append({
                    "name":     row[1] if len(row) > 1 else "file",
                    "content":  row[2] if len(row) > 2 else "",
                    "saved_at": row[3] if len(row) > 3 else ""
                })

        if files:
            # Populate in-memory cache from sheet
            _file_memory_cache[sender] = files
            logger.info(f"Loaded {len(files)} files from sheet for {sender}")
        else:
            logger.info(f"No files in memory for {sender}")

        return files

    except Exception as e:
        logger.error(f"Load files from sheet error: {e}")
        return []


def get_processed_sheet():
    return get_or_create_sheet("Processed Emails", rows=5000, cols=2)


def load_processed_ids():
    global _processed_ids_cache, _cache_loaded
    if _cache_loaded:
        return _processed_ids_cache
    try:
        sheet = get_processed_sheet()
        if sheet:
            # Add header if empty
            all_values = sheet.get_all_values()
            if not all_values:
                sheet.append_row(["Message ID", "Processed At"])
            else:
                for row in all_values[1:]:
                    if row and row[0] and row[0].strip():
                        _processed_ids_cache.add(row[0].strip())
            logger.info(f"Loaded {len(_processed_ids_cache)} processed IDs")
    except Exception as e:
        logger.error(f"Load IDs error: {e}")
    _cache_loaded = True
    return _processed_ids_cache


def mark_as_processed(msg_id):
    global _processed_ids_cache
    if not msg_id:
        return
    msg_id = str(msg_id).strip()
    if msg_id in _processed_ids_cache:
        return
    _processed_ids_cache.add(msg_id)
    try:
        sheet = get_processed_sheet()
        if sheet:
            sheet.append_row([msg_id, datetime.now().strftime("%d.%m.%Y %H:%M")])
    except Exception as e:
        logger.error(f"Mark processed error: {e}")


def is_processed(msg_id):
    if not msg_id:
        return False
    return str(msg_id).strip() in _processed_ids_cache


def save_to_memory(sender, subject, summary, action, status="Open"):
    try:
        sheet = get_sheet("Sheet1")
        if sheet:
            sheet.append_row([
                datetime.now().strftime("%d.%m.%Y %H:%M"),
                sender, subject, summary, action, status
            ])
            logger.info(f"Saved: {subject}")
    except Exception as e:
        logger.error(f"Save error: {e}")


def read_memory_for_report():
    try:
        sheet = get_sheet("Sheet1")
        if sheet:
            records = sheet.get_all_records()
            pending = [r for r in records if r.get("Status") in ["Open", "Monitoring"]]
            closed  = [r for r in records if r.get("Status") == "Closed"]
            return pending, closed
        return [], []
    except Exception as e:
        logger.error(f"Read memory error: {e}")
        return [], []


def
