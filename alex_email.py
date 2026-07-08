import os
import json
import time
import imaplib
import email
import schedule
import logging
import io
from email.header import decode_header
from datetime import datetime, timedelta
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

IMAP_HOST        = "imappro.zoho.com"
IMAP_PORT        = 993
REMINDER_1_DAYS  = 3
REMINDER_2_DAYS  = 7
REMINDER_3_DAYS  = 14
AUTO_CLOSE_DAYS  = 21
ALWAYS_CC        = "alishir.aliyev@scopeconsulting.az"

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
_file_memory_cache   = {}

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
When any BOQ, Smeta, Excel, cost schedule, or pricing document is provided you must analyse every single discipline, every single sheet, and every single line item with equal depth and rigour. No discipline is summarised. No discipline is skipped.

Civil works: review every excavation, concrete, formwork, reinforcement, and masonry position individually.
Fit-out works: review every partition, ceiling, flooring, tiling, painting, joinery, and door position individually.
MEP works — IDENTICAL DEPTH TO CIVIL AND FIT-OUT — NEVER SUMMARISE AS A BLOCK:
Mechanical: every ductwork item, FCU, chiller, AHU, VAV box, grille, diffuser, insulation, pipework, pump, valve individually.
Electrical: every cable run, tray, board, fixture, emergency light, socket, switch, earthing item individually.
Fire protection: every sprinkler head, pipe, pump, panel, detector individually.
Low voltage: every access control point, camera, cabling run, rack individually.
Plumbing: every pipe run, fitting, fixture, valve individually.

For every line item: description, quantity, unit rate, market assessment, Baku market range, quantity concerns.
Flag all missing scope. Discipline-by-discipline risk summary. Combined total risk exposure at end.

SIGNATURE — always end every email with exactly this:

Alex Rivera
Construction Expert
SCOPE Consulting MMC
internal@scope-iq.io

BAKU MARKET RATES:
Mechanical excavation 8 to 15 AZN per cubic metre, Manual excavation 60 to 90 AZN per cubic metre, Concrete C25/30 190 to 240 AZN per cubic metre, Reinforcement 1200 to 1500 AZN per tonne, Formwork 18 to 28 AZN per square metre, External brickwork 25 to 40 AZN per square metre, Internal brickwork 20 to 32 AZN per square metre, Plastering 16 to 24 AZN per square metre, Paint 12 to 18 AZN per square metre, Ceramic tiles 25 to 40 AZN per square metre, Premium tiles 45 to 80 AZN per square metre, Gypsum partition 32 to 48 AZN per square metre, Armstrong ceiling 28 to 42 AZN per square metre, Raised access floor 55 to 85 AZN per square metre, Epoxy floor 35 to 55 AZN per square metre, Carpet tiles 40 to 70 AZN per square metre, Aluminium glazing 180 to 280 AZN per square metre, Timber door 350 to 550 AZN each, Fire door 600 to 1200 AZN each, Metal door 4500 to 6500 AZN each, Aluminium door 700 to 900 AZN each, HVAC ductwork 45 to 75 AZN per square metre, Fan coil unit 350 to 600 AZN each, Chiller 120 to 200 AZN per kW, AHU 800 to 2500 AZN each, VAV box 250 to 600 AZN each, Grille and diffuser 35 to 85 AZN each, Duct insulation 15 to 28 AZN per square metre, Plumbing pipework 25 to 55 AZN per metre, Sanitary fixture 180 to 450 AZN each, Pump 800 to 3500 AZN each, Cable tray 35 to 65 AZN per metre, LV cable 8 to 25 AZN per metre, Distribution board 800 to 3500 AZN each, Lighting fixture 45 to 120 AZN each, Emergency lighting 80 to 180 AZN each, Socket and switch 25 to 65 AZN each, Fire alarm panel 1500 to 8000 AZN each, Smoke detector 45 to 120 AZN each, Sprinkler head 25 to 55 AZN each, Sprinkler pipework 18 to 45 AZN per metre, Fire pump 3500 to 12000 AZN each, Access control 800 to 2500 AZN per door, CCTV camera 250 to 600 AZN each, Structured cabling point 85 to 180 AZN each, Concrete paving 35 to 55 AZN per square metre, Natural stone paving 85 to 150 AZN per square metre, Passenger lift 45000 to 80000 AZN each, Freight lift 60000 to 120000 AZN each."""


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


def get_or_create_sheet(title, rows=1000, cols=5):
    try:
        client      = get_gspread_client()
        spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
        try:
            return spreadsheet.worksheet(title)
        except:
            sheet = spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
            logger.info(f"Created: {title}")
            return sheet
    except Exception as e:
        logger.error(f"Sheet create error ({title}): {e}")
        return None


def save_files_to_memory(sender, attachments):
    global _file_memory_cache
    sender = sender.strip().lower()
    _file_memory_cache[sender] = [
        {"name": att["name"], "content": att["content"],
         "saved_at": datetime.now().strftime("%d.%m.%Y %H:%M")}
        for att in attachments
    ]
    try:
        sheet = get_or_create_sheet("File Memory", rows=2000, cols=4)
        if not sheet:
            return
        all_values = sheet.get_all_values()
        if not all_values:
            sheet.append_row(["Sender", "Filename", "Content", "Saved At"])
            all_values = [["Sender", "Filename", "Content", "Saved At"]]
        rows_to_delete = [i + 1 for i, row in enumerate(all_values)
                          if i > 0 and row and row[0].strip().lower() == sender]
        for r in reversed(rows_to_delete):
            sheet.delete_rows(r)
        saved_at = datetime.now().strftime("%d.%m.%Y %H:%M")
        for att in attachments:
            sheet.append_row([sender, att["name"], att["content"][:10000], saved_at])
    except Exception as e:
        logger.error(f"Save files error: {e}")


def load_files_from_memory(sender):
    global _file_memory_cache
    sender = sender.strip().lower()
    if sender in _file_memory_cache and _file_memory_cache[sender]:
        return _file_memory_cache[sender]
    try:
        sheet = get_or_create_sheet("File Memory", rows=2000, cols=4)
        if not sheet:
            return []
        files = []
        for row in sheet.get_all_values()[1:]:
            if row and len(row) >= 3 and row[0].strip().lower() == sender:
                files.append({"name": row[1], "content": row[2],
                               "saved_at": row[3] if len(row) > 3 else ""})
        if files:
            _file_memory_cache[sender] = files
        return files
    except Exception as e:
        logger.error(f"Load files error: {e}")
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


def save_to_monitoring(sender, subject, summary, action, thread_id, status="Monitoring"):
    try:
        sheet = get_sheet("Sheet1")
        if sheet:
            headers = sheet.row_values(1)
            for col, name in [(7, "Thread-ID"), (8, "Last Reminded"), (9, "Reminder Count")]:
                if len(headers) < col:
                    sheet.update_cell(1, col, name)
            sheet.append_row([
                datetime.now().strftime("%d.%m.%Y %H:%M"),
                sender, subject, summary, action, status,
                thread_id, "", "0"
            ])
    except Exception as e:
        logger.error(f"Save monitoring error: {e}")


def save_to_memory(sender, subject, summary, action, status="Open"):
    try:
        sheet = get_sheet("Sheet1")
        if sheet:
            sheet.append_row([
                datetime.now().strftime("%d.%m.%Y %H:%M"),
                sender, subject, summary, action, status, "", "", "0"
            ])
    except Exception as e:
        logger.error(f"Save error: {e}")


def update_row(row_number, status=None, last_reminded=None, reminder_count=None):
    try:
        sheet = get_sheet("Sheet1")
        if sheet:
            if status:
                sheet.update_cell(row_number, 6, status)
            if last_reminded:
                sheet.update_cell(row_number, 8, last_reminded)
            if reminder_count is not None:
                sheet.update_cell(row_number, 9, str(reminder_count))
    except Exception as e:
        logger.error(f"Update row error: {e}")


def read_memory_for_report():
    try:
        sheet = get_sheet("Sheet1")
        if sheet:
            records = sheet.get_all_records()
            pending = [r for r in records if r.get("Status") in ["Open", "Monitoring"]]
            closed  = [r for r in records if r.get("Status") in ["Closed", "Closed — No Response"]]
            return pending, closed
        return [], []
    except Exception as e:
        logger.error(f"Read memory error: {e}")
        return [], []


def get_first_name(email_address):
    try:
        local = email_address.split("@")[0]
        parts = local.replace(".", " ").replace("_", " ").split()
        return parts[0].capitalize() if parts else "Colleague"
    except:
        return "Colleague"


def build_reply_html(body_text):
    """Wrap Alex reply in branded HTML template"""
    today    = datetime.now().strftime("%d %B %Y")
    time_now = datetime.now().strftime("%H:%M")

    # Split into paragraphs
    paragraphs = body_text.strip().split("\n\n")
    html_body  = ""

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue

        # Detect signature block
        if "Alex Rivera" in para and "SCOPE Consulting" in para:
            lines    = para.split("\n")
            sig_html = ""
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if "Alex Rivera" in line:
                    sig_html += f'<div style="font-size:13px;font-weight:600;color:#1a2942;">{line}</div>'
                elif "internal@scope-iq.io" in line:
                    sig_html += f'<div style="font-size:12px;color:#3CB496;">{line}</div>'
                else:
                    sig_html += f'<div style="font-size:12px;color:#666;">{line}</div>'
            html_body += f'<div style="margin-top:24px;padding-top:16px;border-top:1px solid #f0f0f0;line-height:1.8;">{sig_html}</div>'

        else:
            # Regular paragraph — handle single line breaks within
            lines     = para.split("\n")
            para_html = "<br>".join(line.strip() for line in lines if line.strip())
            html_body += f'<p style="font-size:14px;color:#333;line-height:1.8;margin:0 0 16px;">{para_html}</p>'

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
<div style="max-width:620px;margin:0 auto;padding:20px 0;">

  <!-- Header -->
  <div style="background:#1a2942;border-radius:12px 12px 0 0;padding:18px 28px;">
    <div style="display:flex;justify-content:space-between;align-items:center;">
      <div style="color:#fff;font-size:20px;font-weight:600;letter-spacing:1px;">
        SCOPE <span style="color:#3CB496;">IQ</span>
      </div>
      <div style="font-size:11px;color:#8facc8;">
        {today} &nbsp;·&nbsp; {time_now} Baku
      </div>
    </div>
    <div style="font-size:12px;color:#8facc8;margin-top:6px;">
      Response from Alex Rivera &nbsp;·&nbsp; Construction Expert
    </div>
  </div>

  <!-- Body -->
  <div style="background:#fff;border:1px solid #e8e8e8;border-top:none;border-radius:0 0 12px 12px;padding:28px 28px 24px;">

    {html_body}

    <!-- Branded footer bar -->
    <div style="background:#f8f9fa;border-radius:6px;padding:10px 14px;margin-top:20px;">
      <div style="font-size:11px;color:#888;line-height:1.6;">
        This response was prepared by
        <strong style="color:#1a2942;">Alex Rivera</strong>,
        Construction Expert at SCOPE Consulting MMC, using
        <strong style="color:#3CB496;">SCOPE IQ</strong> — Intelligent Email Monitoring.
      </div>
    </div>

  </div>
</div>
</body>
</html>"""

    return html


def build_report_html(pending, closed, today, day_name, time_now):
    n_open    = len([r for r in pending if r.get("Status") == "Open"])
    n_monitor = len([r for r in pending if r.get("Status") == "Monitoring"])
    n_closed  = len(closed)
    n_total   = len(pending) + n_closed

    items_html = ""
    if pending:
        for i, r in enumerate(pending[-15:], 1):
            subj    = r.get("Subject") or "No subject"
            sender  = r.get("Sender")  or "Unknown"
            date    = r.get("Date")    or "Not recorded"
            status  = r.get("Status")  or "Open"
            summary = r.get("Summary") or "No summary available"
            action  = r.get("Action")  or "Review and action required"

            status_bg   = "#fff8ee" if status == "Open" else "#e1f5ee"
            status_text = "#9a6000" if status == "Open" else "#0f6e56"

            items_html += f"""
            <div style="border:1px solid #e8e8e8;border-radius:8px;margin-bottom:16px;overflow:hidden;">
              <div style="background:#f8f9fa;padding:12px 16px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #e8e8e8;">
                <span style="font-size:12px;color:#888;font-weight:500;">ITEM {i} OF {len(pending)}</span>
                <span style="background:{status_bg};color:{status_text};font-size:11px;padding:3px 10px;border-radius:20px;font-weight:500;">{status}</span>
              </div>
              <div style="padding:16px;">
                <table style="width:100%;font-size:13px;border-collapse:collapse;">
                  <tr><td style="color:#888;padding:4px 0;width:120px;">Subject</td><td style="color:#1a2942;font-weight:600;padding:4px 0;">{subj}</td></tr>
                  <tr><td style="color:#888;padding:4px 0;">From</td><td style="color:#333;padding:4px 0;">{sender}</td></tr>
                  <tr><td style="color:#888;padding:4px 0;">Date received</td><td style="color:#333;padding:4px 0;">{date}</td></tr>
                </table>
                <div style="border-top:1px solid #f0f0f0;margin:12px 0;"></div>
                <div style="margin-bottom:10px;">
                  <div style="font-size:11px;color:#888;font-weight:500;letter-spacing:0.5px;text-transform:uppercase;margin-bottom:4px;">Content summary</div>
                  <div style="font-size:13px;color:#444;line-height:1.6;">{summary}</div>
                </div>
                <div style="background:#fff8ee;border:1px solid #f0c060;border-radius:6px;padding:10px 12px;">
                  <div style="font-size:10px;font-weight:600;color:#9a6000;letter-spacing:0.5px;text-transform:uppercase;margin-bottom:4px;">Action required</div>
                  <div style="font-size:13px;color:#5a3a00;line-height:1.5;">{action}</div>
                </div>
              </div>
            </div>"""

    no_items_html = ""
    if not pending:
        no_items_html = """
        <div style="text-align:center;padding:32px;">
          <div style="width:48px;height:48px;border-radius:50%;background:#e1f5ee;display:flex;align-items:center;justify-content:center;margin:0 auto 12px;font-size:22px;color:#3CB496;">&#10003;</div>
          <div style="font-size:15px;color:#333;font-weight:500;margin-bottom:6px;">All clear</div>
          <div style="font-size:13px;color:#888;">No outstanding emails or open action items as of today.</div>
        </div>"""

    greeting = f"Good morning. Please find below the daily email monitoring report for <strong>{today}</strong>. The following <strong>{len(pending)} item(s)</strong> require your attention." if pending else f"Good morning. This is your daily email monitoring report for <strong>{today}</strong>. All monitored threads are closed or have received responses."

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
<div style="max-width:620px;margin:0 auto;padding:20px 0;">

  <!-- Header -->
  <div style="background:#1a2942;border-radius:12px 12px 0 0;padding:24px 28px;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
      <div style="color:#fff;font-size:22px;font-weight:600;letter-spacing:1px;">SCOPE <span style="color:#3CB496;">IQ</span></div>
      <div style="background:#3CB496;color:#fff;font-size:11px;padding:4px 12px;border-radius:20px;font-weight:500;">Daily Report</div>
    </div>
    <table style="width:100%;font-size:12px;border-collapse:collapse;">
      <tr>
        <td style="color:#8facc8;padding:2px 0;width:50%;">Date &nbsp;<strong style="color:#c8ddf0;">{day_name}, {today}</strong></td>
        <td style="color:#8facc8;padding:2px 0;">Time &nbsp;<strong style="color:#c8ddf0;">{time_now} Baku</strong></td>
      </tr>
      <tr>
        <td style="color:#8facc8;padding:2px 0;">Prepared by &nbsp;<strong style="color:#c8ddf0;">Alex Rivera</strong></td>
        <td style="color:#8facc8;padding:2px 0;">Status &nbsp;<strong style="color:#c8ddf0;">{"All clear" if not pending else f"{len(pending)} item(s) open"}</strong></td>
      </tr>
    </table>
  </div>

  <!-- Stats bar -->
  <div style="background:#243550;padding:12px 28px;display:flex;">
    <div style="flex:1;text-align:center;border-right:1px solid #1a2942;">
      <div style="font-size:22px;font-weight:600;color:#f0a030;">{n_open}</div>
      <div style="font-size:11px;color:#8facc8;margin-top:2px;">Open</div>
    </div>
    <div style="flex:1;text-align:center;border-right:1px solid #1a2942;">
      <div style="font-size:22px;font-weight:600;color:#3CB496;">{n_monitor}</div>
      <div style="font-size:11px;color:#8facc8;margin-top:2px;">Monitoring</div>
    </div>
    <div style="flex:1;text-align:center;border-right:1px solid #1a2942;">
      <div style="font-size:22px;font-weight:600;color:#6ab87a;">{n_closed}</div>
      <div style="font-size:11px;color:#8facc8;margin-top:2px;">Closed</div>
    </div>
    <div style="flex:1;text-align:center;">
      <div style="font-size:22px;font-weight:600;color:#8facc8;">{n_total}</div>
      <div style="font-size:11px;color:#8facc8;margin-top:2px;">Total</div>
    </div>
  </div>

  <!-- Body -->
  <div style="background:#fff;border:1px solid #e8e8e8;border-top:none;border-radius:0 0 12px 12px;padding:24px 28px;">
    <p style="font-size:14px;color:#444;line-height:1.7;margin:0 0 20px;">{greeting}</p>
    {"<div style='font-size:11px;font-weight:600;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;'>Outstanding items</div>" if pending else ""}
    {items_html}
    {no_items_html}

    <!-- Footer -->
    <div style="border-top:1px solid #f0f0f0;margin-top:24px;padding-top:20px;display:flex;justify-content:space-between;align-items:flex-end;">
      <div style="font-size:12px;color:#666;line-height:1.8;">
        <strong style="color:#1a2942;font-size:13px;">Alex Rivera</strong><br>
        Construction Expert<br>
        SCOPE Consulting MMC<br>
        <span style="color:#3CB496;">internal@scope-iq.io</span>
      </div>
      <div style="font-size:11px;color:#aaa;text-align:right;line-height:1.7;">
        Generated automatically<br>
        SCOPE IQ Email Monitoring<br>
        09:00 Baku daily
      </div>
    </div>

    <!-- Chase protocol -->
    <div style="background:#f8f9fa;border-radius:6px;padding:10px 14px;margin-top:16px;">
      <div style="font-size:11px;color:#888;line-height:1.6;">
        <strong style="color:#555;">Chase protocol:</strong>
        Reminder at day 3 &nbsp;·&nbsp; Second reminder at day 7 &nbsp;·&nbsp;
        Escalation at day 14 &nbsp;·&nbsp; Auto-close at day 21
      </div>
    </div>
  </div>
</div>
</body>
</html>"""

    return html


def check_thread_replies(thread_ids):
    replied_threads = set()
    if not thread_ids:
        return replied_threads
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
        mail.select("INBOX")
        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            mail.logout()
            return replied_threads
        for eid in data[0].split()[-100:]:
            try:
                typ, msg_data = mail.fetch(eid, "(RFC822)")
                if typ != "OK" or not msg_data or not msg_data[0]:
                    continue
                msg         = email.message_from_bytes(msg_data[0][1])
                in_reply_to = safe_decode(msg.get("In-Reply-To", "")).strip()
                references  = safe_decode(msg.get("References",  "")).strip()
                sender      = extract_email_address(msg.get("From", ""))
                if ZOHO_EMAIL.lower() in sender:
                    continue
                for tid in thread_ids:
                    if tid and (tid in in_reply_to or tid in references):
                        replied_threads.add(tid)
            except:
                continue
        mail.logout()
    except Exception as e:
        logger.error(f"Check replies error: {e}")
    return replied_threads


def check_followup_reminders():
    logger.info("Checking follow-up reminders...")
    try:
        sheet = get_sheet("Sheet1")
        if not sheet:
            return

        all_values = sheet.get_all_values()
        if not all_values or len(all_values) < 2:
            return

        today = datetime.now()

        monitoring_threads = {}
        for i, row in enumerate(all_values[1:], start=2):
            if len(row) >= 6 and row[5].strip() == "Monitoring":
                tid = row[6].strip() if len(row) > 6 else ""
                if tid:
                    monitoring_threads[tid] = i

        replied = check_thread_replies(set(monitoring_threads.keys()))
        for tid, row_num in monitoring_threads.items():
            if tid in replied:
                update_row(row_num, status="Closed")

        all_values = sheet.get_all_values()

        for i, row in enumerate(all_values[1:], start=2):
            try:
                if len(row) < 6:
                    continue

                date_str       = row[0].strip() if row[0] else ""
                sender         = row[1].strip() if len(row) > 1 else ""
                subject        = row[2].strip() if len(row) > 2 else ""
                summary        = row[3].strip() if len(row) > 3 else ""
                action         = row[4].strip() if len(row) > 4 else ""
                status         = row[5].strip() if len(row) > 5 else ""
                last_reminded  = row[7].strip() if len(row) > 7 else ""
                reminder_count = int(row[8].strip()) if len(row) > 8 and row[8].strip().isdigit() else 0

                if status not in ["Open", "Monitoring"]:
                    continue
                if not sender or "@" not in sender:
                    continue
                if not any(t in sender for t in SCOPE_TEAM_EMAILS):
                    continue

                try:
                    logged_date = datetime.strptime(date_str, "%d.%m.%Y %H:%M")
                except:
                    continue

                days_open = (today - logged_date).days

                if days_open >= AUTO_CLOSE_DAYS:
                    update_row(i, status="Closed — No Response")
                    notice  = f"Dear Team,\n\nThe following email thread has been automatically closed after {AUTO_CLOSE_DAYS} days with no response.\n\n"
                    notice += f"Subject: {subject}\nFrom: {sender}\nDate raised: {date_str}\nDays open: {days_open}\n\n"
                    notice += f"Summary:\n{summary}\n\nNo further reminders will be sent.\n\n"
                    notice += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    send_email(REPORT_RECIPIENTS,
                               f"Auto-Closed — No Response — {subject}",
                               notice,
                               html_body=build_reply_html(notice))
                    continue

                reminder_due = None
                if days_open >= REMINDER_3_DAYS and reminder_count < 3:
                    reminder_due = 3
                elif days_open >= REMINDER_2_DAYS and reminder_count < 2:
                    reminder_due = 2
                elif days_open >= REMINDER_1_DAYS and reminder_count < 1:
                    reminder_due = 1

                if not reminder_due:
                    continue

                if last_reminded:
                    try:
                        last_date = datetime.strptime(last_reminded, "%d.%m.%Y %H:%M")
                        if (today - last_date).days < 3:
                            continue
                    except:
                        pass

                first_name = get_first_name(sender)

                if reminder_due == 1:
                    subject_line = f"Follow-up — {subject}"
                    body  = f"Dear {first_name},\n\n"
                    body += f"I am writing to follow up on the email referenced below, which was sent {days_open} days ago and appears to be awaiting a response.\n\n"
                    body += f"Subject: {subject}\nDate raised: {date_str}\n\nSummary:\n{summary}\n\nAction required:\n{action}\n\n"
                    body += f"I would be grateful if you could review this matter and respond at your earliest convenience.\n\n"
                    body += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    cc = [ALWAYS_CC] if sender.lower() != ALWAYS_CC.lower() else []

                elif reminder_due == 2:
                    subject_line = f"Second Follow-up — {subject}"
                    body  = f"Dear {first_name},\n\n"
                    body += f"This is a second follow-up regarding the matter below, which has now been open for {days_open} days without a response.\n\n"
                    body += f"Subject: {subject}\nDate raised: {date_str}\n\nSummary:\n{summary}\n\nAction required:\n{action}\n\n"
                    body += f"This matter requires your urgent attention. Please respond or confirm the current status as soon as possible.\n\n"
                    body += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    cc = [ALWAYS_CC] if sender.lower() != ALWAYS_CC.lower() else []

                elif reminder_due == 3:
                    subject_line = f"Escalation Notice — {subject}"
                    body  = f"Dear {first_name},\n\n"
                    body += f"This is a formal escalation notice. The email thread referenced below has been open for {days_open} days and has not received a response despite two previous reminders.\n\n"
                    body += f"Subject: {subject}\nDate raised: {date_str}\n\nSummary:\n{summary}\n\nAction required:\n{action}\n\n"
                    body += f"Please be advised that if no response is received within 7 days this matter will be automatically closed and recorded as unresolved. This notice has been copied to SCOPE Consulting management for awareness.\n\n"
                    body += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    cc = [r for r in REPORT_RECIPIENTS if r.lower() != sender.lower()]

                sent = send_email([sender], subject_line, body,
                                  cc_emails=cc,
                                  html_body=build_reply_html(body))
                if sent:
                    update_row(i, last_reminded=today.strftime("%d.%m.%Y %H:%M"),
                               reminder_count=reminder_due)
                    logger.info(f"Reminder {reminder_due} sent: {subject}")

            except Exception as e:
                logger.error(f"Reminder row error: {e}")
                continue

    except Exception as e:
        logger.error(f"Follow-up check error: {e}")


def safe_decode(value, fallback=""):
    if value is None:
        return fallback
    try:
        parts = decode_header(str(value))
        result = ""
        for part, enc in parts:
            if isinstance(part, bytes):
                result += part.decode(enc or "utf-8", errors="replace")
            else:
                result += str(part)
        return result
    except:
        return str(value) if value else fallback


def extract_email_address(header_value):
    if not header_value:
        return ""
    header_str = safe_decode(header_value)
    if "<" in header_str and ">" in header_str:
        start = header_str.rfind("<") + 1
        end   = header_str.rfind(">")
        return header_str[start:end].strip().lower()
    return header_str.strip().lower()


def extract_all_emails(header_value):
    if not header_value:
        return []
    header_str = safe_decode(header_value)
    addresses  = []
    for part in header_str.split(","):
        addr = extract_email_address(part.strip())
        if addr and "@" in addr:
            addresses.append(addr)
    return addresses


def extract_attachments(msg):
    attachments = []
    try:
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition", ""))
            filename = part.get_filename()
            if not filename and "attachment" not in content_disposition:
                continue
            filename = safe_decode(filename, "attachment")
            payload  = part.get_payload(decode=True)
            if not payload:
                continue
            ext = filename.lower().split(".")[-1] if "." in filename else ""

            if ext in ["xlsx", "xls"]:
                try:
                    import openpyxl
                    wb   = openpyxl.load_workbook(io.BytesIO(payload))
                    text = f"Excel file: {filename}\nSheets: {', '.join(wb.sheetnames)}\n\n"
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        text += f"\n{'='*40}\nSHEET: {sheet_name}\n{'='*40}\n"
                        for row in ws.iter_rows(max_row=500, values_only=True):
                            row_data = [str(c) for c in row if c is not None]
                            if row_data:
                                text += " | ".join(row_data) + "\n"
                    attachments.append({"name": filename, "content": text[:15000]})
                except Exception as e:
                    logger.error(f"Excel error: {e}")

            elif ext == "pdf":
                try:
                    import fitz
                    doc  = fitz.open(stream=payload, filetype="pdf")
                    text = f"PDF: {filename}\n" + "".join(p.get_text() for p in doc)
                    doc.close()
                    attachments.append({"name": filename, "content": text[:5000]})
                except Exception as e:
                    logger.error(f"PDF error: {e}")

            elif ext in ["docx", "doc"]:
                try:
                    import docx
                    document = docx.Document(io.BytesIO(payload))
                    text = f"Word: {filename}\n" + "\n".join(p.text for p in document.paragraphs)
                    attachments.append({"name": filename, "content": text[:5000]})
                except Exception as e:
                    logger.error(f"Word error: {e}")

    except Exception as e:
        logger.error(f"Attachment error: {e}")
    return attachments


def get_email_body(msg):
    body = ""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                try:
                    if part.get_content_type() == "text/plain" and not part.get_filename():
                        payload = part.get_payload(decode=True)
                        if payload:
                            body += payload.decode(
                                part.get_content_charset() or "utf-8", errors="replace")
                except:
                    continue
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = payload.decode(
                    msg.get_content_charset() or "utf-8", errors="replace")
    except Exception as e:
        logger.error(f"Body error: {e}")
    return body[:2000]


def send_email(to_emails, subject, body, reply_to_msg_id=None,
               references=None, cc_emails=None, html_body=None):
    try:
        resend.api_key = RESEND_API_KEY
        if isinstance(to_emails, str):
            to_emails = [to_emails]
        to_emails = [e for e in to_emails if e.lower() != ZOHO_EMAIL.lower()]
        if not to_emails:
            return False
        params = {
            "from":    f"Alex Rivera <{ZOHO_EMAIL}>",
            "to":      to_emails,
            "subject": subject,
            "text":    body,
            "headers": {}
        }
        if html_body:
            params["html"] = html_body
        if cc_emails:
            cc_emails = [e for e in cc_emails if e.lower() != ZOHO_EMAIL.lower()]
            if cc_emails:
                params["cc"] = cc_emails
        if reply_to_msg_id:
            params["headers"]["In-Reply-To"] = reply_to_msg_id
            params["headers"]["References"]  = references or reply_to_msg_id
        resend.Emails.send(params)
        logger.info(f"Sent to {to_emails}")
        return True
    except Exception as e:
        logger.error(f"Resend error: {e}")
        return False


def analyse_email(sender, subject, body, attachments=None, memory_files=None, is_cc=False):
    try:
        active_files = attachments if attachments else memory_files
        using_memory = not attachments and bool(memory_files)

        if is_cc:
            full_content = body
            if active_files:
                full_content += "\n\nFILES:\n" + "\n".join(
                    f"{f['name']}:\n{f['content']}" for f in active_files)
            prompt = f"""CC'd email — internal analysis only. Do not reply to sender.
From: {sender}
Subject: {subject}
Content: {full_content}

Full internal assessment:
1. Email type.
2. Summary in three to five sentences.
3. Commercial, technical, contractual or programme implications.
4. Specific actions — what, who, by when.
5. Risk level: High, Medium or Low.
6. Recommended response deadline.
Plain professional prose. Numbered paragraphs. No symbols."""

        else:
            if active_files:
                memory_note = ""
                if using_memory:
                    saved_at   = active_files[0].get("saved_at", "previously")
                    file_names = ", ".join(f["name"] for f in active_files)
                    memory_note = f"Note: Using files previously received from this sender ({file_names}, saved {saved_at}).\n\n"
                files_content = "\n".join(
                    f"{f['name']}:\n{f['content']}" for f in active_files)
                prompt = f"""Email from SCOPE team member.
From: {sender}
Subject: {subject}
Email body: {body}

{memory_note}FILE CONTENT — ALL SHEETS:
{files_content}

Analyse every discipline, every sheet, every line item equally. Never summarise MEP as a block.
For each item: description, quantity, unit rate, market assessment, Baku market range, quantity concerns.
Flag missing scope. Discipline risk summary. Total risk exposure at end.
Write complete formal professional reply now."""
            else:
                prompt = f"""Email from SCOPE team member.
From: {sender}
Subject: {subject}
Content: {body}
Write complete formal professional reply. If files needed and never provided, request them and confirm reply within minutes."""

        response = anthropic_client.messages.create(
            model=MODEL,
            max_tokens=4000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        logger.error(f"Analysis error: {e}")
        return None


def process_emails():
    logger.info("Checking emails via IMAP...")
    load_processed_ids()
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
        logger.info("IMAP login successful")
        mail.select("INBOX")

        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            mail.logout()
            return

        all_ids = data[0].split()
        if not all_ids:
            mail.logout()
            return

        recent    = all_ids[-50:]
        new_count = 0
        logger.info(f"Checking {len(recent)} emails")

        for eid in reversed(recent):
            try:
                eid_str       = eid.decode() if isinstance(eid, bytes) else str(eid)
                typ, msg_data = mail.fetch(eid, "(RFC822)")
                if typ != "OK" or not msg_data or not msg_data[0]:
                    continue
                raw = msg_data[0][1]
                if not raw:
                    continue

                msg        = email.message_from_bytes(raw)
                msg_id_hdr = safe_decode(msg.get("Message-ID"), "").strip()
                unique_id  = msg_id_hdr or eid_str

                if is_processed(unique_id):
                    continue

                mark_as_processed(unique_id)

                sender       = extract_email_address(msg.get("From", ""))
                subject      = safe_decode(msg.get("Subject"), "No subject")
                to_field     = safe_decode(msg.get("To"),      "").lower()
                cc_field     = safe_decode(msg.get("CC"),      "").lower()
                references   = safe_decode(msg.get("References"), "")
                to_addresses = extract_all_emails(msg.get("To",  ""))
                cc_addresses = extract_all_emails(msg.get("CC",  ""))

                if ZOHO_EMAIL.lower() in sender:
                    continue

                is_direct   = ZOHO_EMAIL.lower() in to_field
                is_cc_email = ZOHO_EMAIL.lower() in cc_field
                is_internal = any(t in sender for t in SCOPE_TEAM_EMAILS)

                if not is_direct and not is_cc_email:
                    continue

                new_count  += 1
                body        = get_email_body(msg)
                attachments = extract_attachments(msg)

                if attachments:
                    save_files_to_memory(sender, attachments)
                    memory_files = None
                else:
                    memory_files = load_files_from_memory(sender)

                logger.info(f"Processing: {sender} | {subject}")

                if is_cc_email:
                    logger.info(f"CC — silent log: {sender}")
                    analysis = analyse_email(sender, subject, body,
                                             attachments=attachments,
                                             memory_files=memory_files, is_cc=True)
                    if analysis:
                        save_to_monitoring(sender, subject, analysis[:400],
                                           analysis[:500], msg_id_hdr, "Monitoring")

                elif is_direct and is_internal:
                    analysis = analyse_email(sender, subject, body,
                                             attachments=attachments,
                                             memory_files=memory_files, is_cc=False)
                    if analysis:
                        reply_sub = f"Re: {subject}" if not subject.startswith("Re:") else subject
                        all_recipients = list(set(
                            [sender] +
                            [a for a in to_addresses if a != ZOHO_EMAIL.lower()] +
                            [a for a in cc_addresses if a != ZOHO_EMAIL.lower()]
                        ))
                        new_references = f"{references} {msg_id_hdr}".strip() if references else msg_id_hdr

                        # Build branded HTML reply
                        html_reply = build_reply_html(analysis)

                        sent = send_email(
                            all_recipients, reply_sub, analysis,
                            reply_to_msg_id=msg_id_hdr,
                            references=new_references,
                            html_body=html_reply
                        )
                        save_to_memory(sender, subject, analysis[:400],
                                       "Replied by Alex",
                                       "Closed" if sent else "Open")

                elif is_direct and not is_internal:
                    logger.info(f"Ignoring external: {sender}")

            except Exception as e:
                logger.error(f"Email error: {e}")
                continue

        mail.logout()
        logger.info(f"Done — {new_count} new emails processed")

    except imaplib.IMAP4.error as e:
        logger.error(f"IMAP auth error: {e}")
    except Exception as e:
        logger.error(f"IMAP error: {e}")


def send_morning_report():
    logger.info("Sending morning report...")
    try:
        pending, closed = read_memory_for_report()
        today    = datetime.now().strftime("%d %B %Y")
        day_name = datetime.now().strftime("%A")
        time_now = datetime.now().strftime("%H:%M")
        n_open   = len([r for r in pending if r.get("Status") == "Open"])

        html_body = build_report_html(pending, closed, today, day_name, time_now)

        plain = f"SCOPE IQ Daily Report — {today}\n\n"
        if pending:
            plain += f"{len(pending)} item(s) require attention.\n\n"
            for i, r in enumerate(pending[-15:], 1):
                plain += f"{i}. {r.get('Subject','?')} | {r.get('Sender','?')} | {r.get('Status','?')}\n"
                plain += f"   Action: {r.get('Action','')}\n\n"
        else:
            plain += "All clear — no outstanding items.\n\n"
        plain += "Alex Rivera | SCOPE Consulting MMC | internal@scope-iq.io"

        subject_line = f"SCOPE IQ Daily Report — {today}"
        subject_line += f" — {len(pending)} Open Item(s)" if pending else " — All Clear"

        send_email(REPORT_RECIPIENTS, subject_line, plain, html_body=html_body)
        logger.info("Morning report sent")

    except Exception as e:
        logger.error(f"Morning report error: {e}")


def main():
    logger.info("Alex Email Service starting")
    logger.info(f"Monitoring: {ZOHO_EMAIL}")
    logger.info(f"Authorised team: {SCOPE_TEAM_EMAILS}")
    logger.info(f"Chase protocol: Day {REMINDER_1_DAYS}/{REMINDER_2_DAYS}/{REMINDER_3_DAYS}/Close {AUTO_CLOSE_DAYS}")

    load_processed_ids()
    logger.info("Duplicate protection active")

    schedule.every(10).minutes.do(process_emails)
    schedule.every().day.at("05:00").do(send_morning_report)
    schedule.every(6).hours.do(check_followup_reminders)

    process_emails()

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
