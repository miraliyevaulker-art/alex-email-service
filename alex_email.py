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
IMAP_TIMEOUT     = 30
REMINDER_1_DAYS  = 3
REMINDER_2_DAYS  = 7
REMINDER_3_DAYS  = 14
AUTO_CLOSE_DAYS  = 21
ALWAYS_CC        = "alishir.aliyev@scopeconsulting.az"
INTERNAL_DOMAIN  = os.environ.get("INTERNAL_DOMAIN", "scopeconsulting.az")
CLIENT_DOMAINS   = [
    d.strip().lower() for d in
    os.environ.get("CLIENT_DOMAINS", "").split(",")
    if d.strip()
]

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


def get_imap_connection():
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.socket().settimeout(IMAP_TIMEOUT)
    mail.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
    return mail


def safe_logout(mail):
    try:
        mail.logout()
    except Exception:
        pass


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


def get_action_tracker_sheet():
    try:
        client      = get_gspread_client()
        spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
        try:
            sheet   = spreadsheet.worksheet("Action Tracker")
            headers = sheet.row_values(1)
            for col, name in [(15, "All Thread Participants"),
                              (16, "Client Emails"),
                              (17, "Responsible Name")]:
                if len(headers) < col:
                    sheet.update_cell(1, col, name)
            return sheet
        except:
            sheet = spreadsheet.add_worksheet(
                title="Action Tracker", rows=2000, cols=17)
            sheet.append_row([
                "Date Logged", "Meeting Reference", "Action Item",
                "Responsible Party", "Responsible Email", "Due Date",
                "Status", "Last Reminded", "Reminder Count",
                "Thread ID", "MOM Sender", "Draft Sent",
                "External Reply", "Notes", "All Thread Participants",
                "Client Emails", "Responsible Name"
            ])
            logger.info("Created Action Tracker tab")
            return sheet
    except Exception as e:
        logger.error(f"Action tracker error: {e}")
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


def save_action_item(meeting_ref, action_item, responsible_party,
                     responsible_email, responsible_name, due_date,
                     thread_id, mom_sender, all_participants,
                     client_emails, status="Open"):
    try:
        sheet = get_action_tracker_sheet()
        if sheet:
            sheet.append_row([
                datetime.now().strftime("%d.%m.%Y %H:%M"),
                meeting_ref, action_item,
                responsible_party, responsible_email, due_date,
                status, "", "0", thread_id, mom_sender, "", "",
                "", ",".join(all_participants),
                ",".join(client_emails),
                responsible_name
            ])
            logger.info(f"Action saved: {action_item[:50]} → {responsible_name} <{responsible_email}>")
    except Exception as e:
        logger.error(f"Save action error: {e}")


def update_action_row(row_number, status=None, last_reminded=None,
                      reminder_count=None, draft_sent=None,
                      external_reply=None, notes=None):
    try:
        sheet = get_action_tracker_sheet()
        if sheet:
            if status:
                sheet.update_cell(row_number, 7, status)
            if last_reminded:
                sheet.update_cell(row_number, 8, last_reminded)
            if reminder_count is not None:
                sheet.update_cell(row_number, 9, str(reminder_count))
            if draft_sent:
                sheet.update_cell(row_number, 12, draft_sent)
            if external_reply:
                sheet.update_cell(row_number, 13, external_reply[:500])
            if notes:
                sheet.update_cell(row_number, 14, notes)
    except Exception as e:
        logger.error(f"Update action row error: {e}")


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


def read_actions_for_report():
    try:
        sheet = get_action_tracker_sheet()
        if sheet:
            records      = sheet.get_all_records()
            open_actions = [r for r in records if r.get("Status") in
                            ["Open", "Draft Pending", "Draft Sent", "Reminded"]]
            closed       = [r for r in records if r.get("Status") in
                            ["Closed", "Closed — No Response"]]
            flagged      = [r for r in records if r.get("Status") == "Email Unknown"]
            return open_actions, closed, flagged
        return [], [], []
    except Exception as e:
        logger.error(f"Read actions error: {e}")
        return [], [], []


def get_first_name(email_address):
    try:
        local = email_address.split("@")[0]
        parts = local.replace(".", " ").replace("_", " ").split()
        return parts[0].capitalize() if parts else "Colleague"
    except:
        return "Colleague"


def is_internal_email(email_addr):
    return INTERNAL_DOMAIN.lower() in email_addr.lower()


def is_client_email(email_addr):
    if not CLIENT_DOMAINS:
        return False
    return any(d in email_addr.lower() for d in CLIENT_DOMAINS)


def is_approval_reply(body_text):
    approval_keywords = [
        "approve", "approved", "send", "go ahead", "ok", "okay",
        "confirmed", "confirm", "yes", "proceed", "looks good",
        "please send", "send it", "agreed"
    ]
    return any(kw in body_text.lower() for kw in approval_keywords)


def is_mom_email(subject, body, attachments):
    mom_keywords = [
        "minutes of meeting", "mom", "meeting minutes",
        "iclasın protokolu", "görüş protokolu",
        "meeting notes", "action items", "action points",
        "minutes from", "meeting summary"
    ]
    text = (subject + " " + body).lower()
    if any(kw in text for kw in mom_keywords):
        return True
    for att in attachments:
        if any(kw in att.get("name", "").lower()
               for kw in ["mom", "minutes", "meeting", "protocol"]):
            return True
    return False


def extract_display_name(header_value):
    if not header_value:
        return ""
    header_str = safe_decode(header_value)
    if "<" in header_str:
        name = header_str[:header_str.rfind("<")].strip()
        return name.strip('"').strip("'").strip()
    return ""


def extract_email_with_name(header_value):
    return {
        "email": extract_email_address(header_value),
        "name":  extract_display_name(header_value)
    }


def extract_all_emails_with_names(header_value):
    if not header_value:
        return []
    header_str = safe_decode(header_value)
    results    = []
    parts      = []
    current    = ""
    in_quote   = False
    for char in header_str:
        if char == '"':
            in_quote = not in_quote
        if char == "," and not in_quote:
            parts.append(current.strip())
            current = ""
        else:
            current += char
    if current.strip():
        parts.append(current.strip())

    for part in parts:
        info = extract_email_with_name(part)
        if info["email"] and "@" in info["email"]:
            results.append(info)
    return results


def classify_thread_participants(all_participants_with_names):
    internal    = []
    clients     = []
    contractors = []
    for p in all_participants_with_names:
        addr = p["email"]
        if addr == ZOHO_EMAIL.lower():
            continue
        if is_internal_email(addr):
            internal.append(p)
        elif is_client_email(addr):
            clients.append(p)
        else:
            contractors.append(p)
    return internal, clients, contractors


def find_action_by_thread(in_reply_to, references):
    try:
        sheet = get_action_tracker_sheet()
        if not sheet:
            return None
        all_values = sheet.get_all_values()
        for i, row in enumerate(all_values[1:], start=2):
            if len(row) < 10:
                continue
            thread_id = row[9].strip() if len(row) > 9 else ""
            status    = row[6].strip() if len(row) > 6 else ""
            if not thread_id or status in ["Closed", "Closed — No Response"]:
                continue
            if thread_id in in_reply_to or thread_id in references:
                data = get_action_data_from_row(row)
                return {"row": i, **data}
        return None
    except Exception as e:
        logger.error(f"Find action error: {e}")
        return None


def extract_mom_actions(mom_content, thread_participants_with_names, subject):
    try:
        internal, clients, contractors = classify_thread_participants(
            thread_participants_with_names)

        def fmt_list(lst):
            return "\n".join([
                f"  - {p['name']} <{p['email']}>" if p['name']
                else f"  - {p['email']}"
                for p in lst
            ]) or "  None detected"

        context = f"""SCOPE team (internal PMC):
{fmt_list(internal)}

Known clients (CC for awareness only, never action owners):
{fmt_list(clients)}

External contractors/consultants (potential action owners):
{fmt_list(contractors)}"""

        prompt = f"""Analyse this Minutes of Meeting and extract all action items.

Meeting subject: {subject}

Participants in this email thread:
{context}

MOM Content:
{mom_content[:8000]}

For each action item:
1. Action description
2. Responsible party name as mentioned in MOM
3. Responsible email — match to contractor list. If matched use email. If no match write UNKNOWN.
4. Responsible display name — full name if known, else empty string
5. Due date or NOT SPECIFIED
6. Meeting reference
7. Party role — CONTRACTOR, CLIENT, or SCOPE

Also identify client company and contractor company from MOM content.

Respond in this exact JSON only:
{{"meeting_reference": "...", "client_identified": "...", "contractor_identified": "...", "actions": [{{"action": "...", "responsible_party": "...", "responsible_email": "...", "responsible_name": "...", "due_date": "...", "party_role": "CONTRACTOR"}}]}}"""

        response = anthropic_client.messages.create(
            model=MODEL,
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        text = response.content[0].text.strip()
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        data = json.loads(text.strip())
        logger.info(f"Extracted {len(data.get('actions', []))} actions")
        return data
    except Exception as e:
        logger.error(f"MOM extraction error: {e}")
        return {"meeting_reference": subject, "actions": [],
                "client_identified": "Unknown",
                "contractor_identified": "Unknown"}


def analyse_external_reply(action_item, reply_content):
    try:
        prompt = f"""Review this external party reply against a required action item.

Action required: {action_item}
Reply received: {reply_content[:3000]}

Has this reply fully and satisfactorily addressed the action?

Respond in this exact JSON only:
{{"satisfied": true or false, "analysis": "2-3 sentence assessment.", "outstanding_items": "What remains or NONE"}}"""

        response = anthropic_client.messages.create(
            model=MODEL,
            max_tokens=500,
            messages=[{"role": "user", "content": prompt}]
        )
        text = response.content[0].text.strip()
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text.strip())
    except Exception as e:
        logger.error(f"Reply analysis error: {e}")
        return {"satisfied": False, "analysis": "Unable to analyse.",
                "outstanding_items": "Unknown"}


def draft_external_reminder(action_item, responsible_name, responsible_party,
                             due_date, meeting_ref, reminder_number,
                             outstanding=None, on_behalf_of=None):
    try:
        tone_map = {1: "polite and professional",
                    2: "firm and urgent",
                    3: "formal escalation"}
        if responsible_name and responsible_name.strip():
            first      = responsible_name.strip().split()[0]
            salutation = f"Dear {first},"
        else:
            company    = responsible_party.strip().split()[0] if responsible_party else "Team"
            salutation = f"Dear {company} Team,"

        outstanding_text = f"\n\nOutstanding items:\n{outstanding}" if outstanding else ""
        on_behalf_line = f" writing on behalf of {on_behalf_of}, SCOPE Consulting MMC" if on_behalf_of else ""

        prompt = f"""Draft a {tone_map.get(reminder_number, 'formal')} reminder email to an external party.

Salutation: {salutation}
Meeting reference: {meeting_ref}
Action item: {action_item}
Responsible: {responsible_name or responsible_party}
Due date: {due_date}
Reminder number: {reminder_number} of 3{outstanding_text}

Start with exactly: {salutation}
In the opening sentence, mention you are{on_behalf_line}, following up on the referenced meeting.
Write complete formal professional email. No bullet points or symbols.
End with:
Alex Rivera
Construction Expert
SCOPE Consulting MMC
internal@scope-iq.io"""

        response = anthropic_client.messages.create(
            model=MODEL,
            max_tokens=800,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        logger.error(f"Draft reminder error: {e}")
        return None


def get_action_data_from_row(row):
    participants_raw = row[14].strip() if len(row) > 14 else ""
    all_participants = [p.strip() for p in participants_raw.split(",")
                        if p.strip() and "@" in p.strip()]
    client_raw    = row[15].strip() if len(row) > 15 else ""
    client_emails = [c.strip() for c in client_raw.split(",")
                     if c.strip() and "@" in c.strip()]
    return {
        "date":             row[0].strip()  if row[0]        else "",
        "meeting_ref":      row[1].strip()  if len(row) > 1  else "",
        "action":           row[2].strip()  if len(row) > 2  else "",
        "responsible":      row[3].strip()  if len(row) > 3  else "",
        "email":            row[4].strip()  if len(row) > 4  else "",
        "due_date":         row[5].strip()  if len(row) > 5  else "",
        "status":           row[6].strip()  if len(row) > 6  else "",
        "last_reminded":    row[7].strip()  if len(row) > 7  else "",
        "reminder_count":   int(row[8].strip()) if len(row) > 8 and row[8].strip().isdigit() else 0,
        "thread_id":        row[9].strip()  if len(row) > 9  else "",
        "mom_sender":       row[10].strip() if len(row) > 10 else "",
        "all_participants": all_participants,
        "client_emails":    client_emails,
        "responsible_name": row[16].strip() if len(row) > 16 else ""
    }


def build_cc_for_external(action_data):
    resp_email = action_data["email"].lower()
    cc_list    = []
    for addr in action_data["all_participants"]:
        addr_lower = addr.lower()
        if addr_lower == ZOHO_EMAIL.lower():
            continue
        if addr_lower == resp_email:
            continue
        cc_list.append(addr)
    for r in REPORT_RECIPIENTS:
        if r.lower() not in [c.lower() for c in cc_list]:
            cc_list.append(r)
    return list(set(cc_list))


def get_all_domain_emails(domain, all_participants):
    return [p for p in all_participants
            if "@" in p and p.split("@")[-1].lower() == domain.lower()]


def route_external_reply_for_approval(sender, subject, body,
                                      action_data, msg_id_hdr):
    analysis   = analyse_external_reply(action_data["action"], body)
    mom_sender = action_data["mom_sender"]
    cc_list    = [r for r in REPORT_RECIPIENTS
                  if r.lower() != mom_sender.lower()]
    resp_label = action_data["responsible_name"] or action_data["responsible"]

    if analysis.get("satisfied"):
        update_action_row(action_data["row"], status="Closed",
                          external_reply=body[:300],
                          notes=analysis.get("analysis", ""))

        notice  = f"Dear {get_first_name(mom_sender)},\n\n"
        notice += f"A reply has been received from {resp_label} and I have assessed it as satisfactorily addressing the required action.\n\n"
        notice += f"Meeting reference: {action_data['meeting_ref']}\n"
        notice += f"Action: {action_data['action']}\n\n"
        notice += f"My assessment:\n{analysis.get('analysis', '')}\n\n"
        notice += f"This action has been closed in the Action Tracker. No further follow-up is required.\n\n"
        notice += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"

        send_email([mom_sender],
                   f"Action Closed — {action_data['action'][:50]}",
                   notice,
                   html_body=build_reply_html(notice),
                   cc_emails=cc_list)
        logger.info(f"Action closed — notified internal team")

    else:
        outstanding = analysis.get("outstanding_items", "")
        draft = draft_external_reminder(
            action_data["action"],
            action_data["responsible_name"],
            action_data["responsible"],
            "As previously agreed",
            action_data["meeting_ref"], 1, outstanding,
            on_behalf_of=get_first_name(mom_sender)
        )
        if draft:
            approval  = f"Dear {get_first_name(mom_sender)},\n\n"
            approval += f"A reply has been received from {resp_label} regarding the action below. My assessment is that it does not fully satisfy the required action.\n\n"
            approval += f"Meeting reference: {action_data['meeting_ref']}\n"
            approval += f"Action: {action_data['action']}\n\n"
            approval += f"My assessment:\n{analysis.get('analysis', '')}\n\n"
            approval += f"Outstanding items:\n{outstanding}\n\n"
            approval += f"I have prepared a follow-up email for your approval. Please reply with approve or send and I will dispatch it to {resp_label} immediately.\n\n"
            approval += f"{'='*50}\nDRAFT FOLLOW-UP TO {resp_label.upper()}:\n{'='*50}\n\n{draft}\n\n{'='*50}\n\n"
            approval += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"

            update_action_row(action_data["row"],
                              status="Draft Pending",
                              external_reply=body[:300],
                              notes=f"Reply unsatisfactory: {outstanding[:200]}")

            send_email([mom_sender],
                       f"Approval Required — Follow-up to {resp_label}",
                       approval,
                       html_body=build_reply_html(approval),
                       cc_emails=cc_list)
            logger.info(f"Draft follow-up sent to {mom_sender} for approval")


def process_mom_email(sender, subject, body, attachments,
                      all_thread_with_names, msg_id_hdr):
    logger.info(f"Processing MOM: {subject}")

    mom_content = body
    if attachments:
        for att in attachments:
            mom_content += f"\n\n{att['name']}:\n{att['content']}"

    extracted     = extract_mom_actions(mom_content, all_thread_with_names, subject)
    meeting_ref   = extracted.get("meeting_reference", subject)
    actions       = extracted.get("actions", [])
    client_id     = extracted.get("client_identified", "Unknown")
    contractor_id = extracted.get("contractor_identified", "Unknown")

    internal, clients, contractors = classify_thread_participants(all_thread_with_names)
    client_emails   = [p["email"] for p in clients]
    all_participants = [p["email"] for p in all_thread_with_names
                        if p["email"] != ZOHO_EMAIL.lower()]

    if not actions:
        logger.info("No actions extracted from MOM")
        save_to_monitoring(sender, subject,
                           "MOM received — no actions extracted",
                           "Review MOM manually", msg_id_hdr, "Monitoring")
        return

    unknown_count = 0
    for action in actions:
        resp_email = action.get("responsible_email", "UNKNOWN")
        resp_name  = action.get("responsible_name", "")
        party_role = action.get("party_role", "CONTRACTOR")
        status     = "Open" if resp_email != "UNKNOWN" else "Email Unknown"
        if resp_email == "UNKNOWN":
            unknown_count += 1
        save_action_item(
            meeting_ref,
            action.get("action", ""),
            action.get("responsible_party", "Unknown"),
            resp_email, resp_name,
            action.get("due_date", "Not specified"),
            msg_id_hdr, sender,
            all_participants, client_emails, status
        )

    def fmt_p(lst):
        if not lst:
            return "None detected"
        return "\n".join([
            f"  {p['name']} <{p['email']}>" if p['name'] else f"  {p['email']}"
            for p in lst
        ])

    party_analysis  = f"My analysis of the parties in this meeting:\n\n"
    party_analysis += f"Client identified from MOM: {client_id}\n"
    if clients:
        party_analysis += f"Client contacts:\n{fmt_p(clients)}\n"
        party_analysis += f"These parties will be CC'd on all follow-up emails for awareness and will never receive direct action requests.\n\n"
    else:
        party_analysis += f"No client email addresses detected. If the client should be copied please provide their email.\n\n"

    party_analysis += f"Contractor identified from MOM: {contractor_id}\n"
    if contractors:
        party_analysis += f"Contractor contacts:\n{fmt_p(contractors)}\n"
        party_analysis += f"These parties will receive action follow-up emails after your approval.\n\n"
    else:
        party_analysis += f"No contractor email addresses detected. Please provide contact details.\n\n"

    party_analysis += f"SCOPE team:\n{fmt_p(internal)}\n"
    party_analysis += f"These parties will be CC'd on all outgoing emails.\n\n"

    if unknown_count:
        party_analysis += f"Note: {unknown_count} action(s) could not be matched to an email address. Please provide the correct contact details.\n\n"

    action_summary = ""
    for i, action in enumerate(actions, 1):
        role_tag = f"[{action.get('party_role', 'CONTRACTOR')}]"
        name     = action.get("responsible_name", "")
        party    = action.get("responsible_party", "Unknown")
        label    = f"{name} ({party})" if name else party
        action_summary += f"{i}. {role_tag} Action: {action.get('action', '')}\n"
        action_summary += f"   Responsible: {label}\n"
        action_summary += f"   Email: {action.get('responsible_email', 'UNKNOWN')}\n"
        action_summary += f"   Due: {action.get('due_date', 'Not specified')}\n\n"

    notification  = f"Dear {get_first_name(sender)},\n\n"
    notification += f"I have analysed the Minutes of Meeting and extracted {len(actions)} action item(s) from {meeting_ref}.\n\n"
    notification += f"{'='*50}\nPARTY IDENTIFICATION — PLEASE CONFIRM\n{'='*50}\n\n"
    notification += party_analysis
    notification += f"{'='*50}\nEXTRACTED ACTION ITEMS\n{'='*50}\n\n"
    notification += action_summary
    notification += f"{'='*50}\n\n"
    notification += f"All actions have been logged in the Action Tracker. Draft follow-up reminders will be sent to you for approval before any email is dispatched to external parties.\n\n"
    notification += f"Please confirm the party identification above is correct by replying with approve or confirmed. If any party is incorrectly identified please advise and I will update before proceeding.\n\n"
    notification += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"

    cc_approval = [r for r in REPORT_RECIPIENTS if r.lower() != sender.lower()]
    send_email(
        [sender],
        f"MOM Action Items — Confirmation Required — {meeting_ref}",
        notification,
        html_body=build_reply_html(notification),
        cc_emails=cc_approval
    )
    logger.info(f"MOM notification sent to {sender}")


def check_external_action_replies():
    logger.info("Checking external action replies...")
    mail = None
    try:
        sheet = get_action_tracker_sheet()
        if not sheet:
            return

        all_values = sheet.get_all_values()
        if not all_values or len(all_values) < 2:
            return

        open_actions = {}
        for i, row in enumerate(all_values[1:], start=2):
            if len(row) < 7:
                continue
            data = get_action_data_from_row(row)
            if data["status"] in ["Open", "Reminded", "Draft Sent"] and \
               data["thread_id"] and data["email"]:
                open_actions[data["thread_id"]] = {"row": i, **data}

        if not open_actions:
            return

        mail = get_imap_connection()
        mail.select("INBOX")
        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            safe_logout(mail)
            return

        for eid in data[0].split()[-100:]:
            try:
                typ, msg_data = mail.fetch(eid, "(RFC822)")
                if typ != "OK" or not msg_data or not msg_data[0]:
                    continue
                msg         = email.message_from_bytes(msg_data[0][1])
                in_reply_to = safe_decode(msg.get("In-Reply-To", "")).strip()
                references  = safe_decode(msg.get("References",  "")).strip()
                from_addr   = extract_email_address(msg.get("From", ""))

                if ZOHO_EMAIL.lower() in from_addr:
                    continue
                if is_internal_email(from_addr):
                    continue

                for thread_id, action_data in open_actions.items():
                    if thread_id and (thread_id in in_reply_to or
                                      thread_id in references):
                        reply_body = get_email_body(msg)
                        logger.info(f"External reply from {from_addr} — routing for approval")
                        route_external_reply_for_approval(
                            from_addr,
                            safe_decode(msg.get("Subject", "")),
                            reply_body, action_data,
                            safe_decode(msg.get("Message-ID", "")).strip()
                        )
            except (imaplib.IMAP4.abort, OSError, EOFError) as fetch_err:
                logger.warning(f"Fetch error — reconnecting: {fetch_err}")
                safe_logout(mail)
                try:
                    mail = get_imap_connection()
                    mail.select("INBOX")
                except Exception as reconnect_err:
                    logger.error(f"Reconnect failed: {reconnect_err}")
                    break
                continue
            except Exception as e:
                logger.error(f"Reply check error: {e}")
                continue

        safe_logout(mail)
    except Exception as e:
        logger.error(f"External reply check error: {e}")
        if mail:
            safe_logout(mail)


def check_action_approvals():
    logger.info("Checking action approvals...")
    mail = None
    try:
        sheet = get_action_tracker_sheet()
        if not sheet:
            return

        all_values        = sheet.get_all_values()
        pending_approvals = {}

        for i, row in enumerate(all_values[1:], start=2):
            if len(row) < 7:
                continue
            data = get_action_data_from_row(row)
            if data["status"] == "Draft Pending" and data["thread_id"]:
                pending_approvals[data["thread_id"]] = {"row": i, **data}

        if not pending_approvals:
            return

        mail = get_imap_connection()
        mail.select("INBOX")
        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            safe_logout(mail)
            return

        for eid in data[0].split()[-50:]:
            try:
                typ, msg_data = mail.fetch(eid, "(RFC822)")
                if typ != "OK" or not msg_data or not msg_data[0]:
                    continue
                msg         = email.message_from_bytes(msg_data[0][1])
                from_addr   = extract_email_address(msg.get("From", ""))
                in_reply_to = safe_decode(msg.get("In-Reply-To", "")).strip()
                references  = safe_decode(msg.get("References",  "")).strip()
                body        = get_email_body(msg)

                if not is_internal_email(from_addr):
                    continue
                if not is_approval_reply(body):
                    continue

                for thread_id, action_data in pending_approvals.items():
                    if thread_id and (thread_id in in_reply_to or
                                      thread_id in references):
                        reminder_count = action_data["reminder_count"] + 1
                        draft = draft_external_reminder(
                            action_data["action"],
                            action_data["responsible_name"],
                            action_data["responsible"],
                            action_data["due_date"],
                            action_data["meeting_ref"],
                            reminder_count,
                            on_behalf_of=get_first_name(action_data["mom_sender"])
                        )
                        if draft and action_data["email"] not in ["UNKNOWN", ""]:
                            if not action_data["responsible_name"]:
                                domain    = action_data["email"].split("@")[-1] if "@" in action_data["email"] else ""
                                all_to    = get_all_domain_emails(domain, action_data["all_participants"])
                                to_emails = all_to if all_to else [action_data["email"]]
                            else:
                                to_emails = [action_data["email"]]

                            cc_list      = build_cc_for_external(action_data)
                            resp_label   = action_data["responsible_name"] or action_data["responsible"]
                            reminder_tag = {1: "Follow-up", 2: "Second Follow-up",
                                            3: "Escalation Notice"}.get(reminder_count, "Follow-up")

                            html_reminder = build_external_reminder_html(
                                draft,
                                action_data["meeting_ref"],
                                action_data["action"],
                                resp_label,
                                action_data["due_date"],
                                reminder_tag,
                                on_behalf_of=get_first_name(action_data["mom_sender"])
                            )

                            sent = send_email(
                                to_emails,
                                f"Action Item Follow-up — {action_data['meeting_ref']}",
                                draft,
                                html_body=html_reminder,
                                cc_emails=cc_list
                            )
                            if sent:
                                update_action_row(
                                    action_data["row"],
                                    status="Reminded",
                                    last_reminded=datetime.now().strftime("%d.%m.%Y %H:%M"),
                                    reminder_count=reminder_count,
                                    draft_sent=datetime.now().strftime("%d.%m.%Y %H:%M")
                                )
                                logger.info(f"Approved — sent to {to_emails} CC {cc_list}")
            except (imaplib.IMAP4.abort, OSError, EOFError) as fetch_err:
                logger.warning(f"Fetch error — reconnecting: {fetch_err}")
                safe_logout(mail)
                try:
                    mail = get_imap_connection()
                    mail.select("INBOX")
                except Exception as reconnect_err:
                    logger.error(f"Reconnect failed: {reconnect_err}")
                    break
                continue
            except Exception as e:
                logger.error(f"Approval check error: {e}")
                continue

        safe_logout(mail)
    except Exception as e:
        logger.error(f"Action approval check error: {e}")
        if mail:
            safe_logout(mail)


def check_external_action_reminders():
    logger.info("Checking external action reminders...")
    try:
        sheet = get_action_tracker_sheet()
        if not sheet:
            return

        all_values = sheet.get_all_values()
        today      = datetime.now()

        for i, row in enumerate(all_values[1:], start=2):
            try:
                if len(row) < 7:
                    continue
                data = get_action_data_from_row(row)

                if data["status"] not in ["Open", "Reminded"]:
                    continue
                if not data["email"] or data["email"] in ["UNKNOWN", ""]:
                    continue

                try:
                    logged_date = datetime.strptime(data["date"], "%d.%m.%Y %H:%M")
                except:
                    continue

                days_open  = (today - logged_date).days
                resp_label = data["responsible_name"] or data["responsible"]

                if days_open >= AUTO_CLOSE_DAYS:
                    update_action_row(i, status="Closed — No Response",
                                      notes=f"Auto-closed after {days_open} days")
                    notice  = f"Dear {get_first_name(data['mom_sender'])},\n\n"
                    notice += f"The following action item has been automatically closed after {AUTO_CLOSE_DAYS} days with no response from {resp_label}.\n\n"
                    notice += f"Meeting reference: {data['meeting_ref']}\nAction: {data['action']}\n"
                    notice += f"Responsible: {resp_label} ({data['email']})\nDays open: {days_open}\n\n"
                    notice += f"Please advise if further action is required.\n\n"
                    notice += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    cc = [r for r in REPORT_RECIPIENTS if r.lower() != data["mom_sender"].lower()]
                    send_email([data["mom_sender"]],
                               f"External Action Auto-Closed — {data['action'][:50]}",
                               notice, html_body=build_reply_html(notice), cc_emails=cc)
                    continue

                reminder_due = None
                if days_open >= REMINDER_3_DAYS and data["reminder_count"] < 3:
                    reminder_due = 3
                elif days_open >= REMINDER_2_DAYS and data["reminder_count"] < 2:
                    reminder_due = 2
                elif days_open >= REMINDER_1_DAYS and data["reminder_count"] < 1:
                    reminder_due = 1

                if not reminder_due:
                    continue

                if data["last_reminded"]:
                    try:
                        last_date = datetime.strptime(data["last_reminded"], "%d.%m.%Y %H:%M")
                        if (today - last_date).days < 3:
                            continue
                    except:
                        pass

                tone_label = {1: "Follow-up", 2: "Second Follow-up",
                              3: "Escalation Notice"}[reminder_due]

                draft = draft_external_reminder(
                    data["action"], data["responsible_name"],
                    data["responsible"], data["due_date"],
                    data["meeting_ref"], reminder_due,
                    on_behalf_of=get_first_name(data["mom_sender"])
                )

                if draft:
                    if not data["responsible_name"]:
                        domain    = data["email"].split("@")[-1] if "@" in data["email"] else ""
                        all_to    = get_all_domain_emails(domain, data["all_participants"])
                        to_preview = ", ".join(all_to) if all_to else data["email"]
                    else:
                        to_preview = data["email"]

                    cc_preview = build_cc_for_external(data)

                    approval  = f"Dear {get_first_name(data['mom_sender'])},\n\n"
                    approval += f"The action item below from {data['meeting_ref']} has been open for {days_open} days without a response from {resp_label}.\n\n"
                    approval += f"Action: {data['action']}\nResponsible: {resp_label} ({data['email']})\nDue date: {data['due_date']}\n\n"
                    approval += f"I have prepared a {tone_label.lower()} for your approval. Please reply with approve or send to dispatch.\n\n"
                    approval += f"When approved this email will be sent:\nTo: {to_preview}\nCC: {', '.join(cc_preview)}\n\n"
                    approval += f"{'='*50}\nDRAFT — {tone_label.upper()} TO {resp_label.upper()}:\n{'='*50}\n\n{draft}\n\n{'='*50}\n\n"
                    approval += f"Kind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"

                    cc = [r for r in REPORT_RECIPIENTS
                          if r.lower() != data["mom_sender"].lower()]

                    sent = send_email(
                        [data["mom_sender"]],
                        f"Approval Required — {tone_label} to {resp_label}",
                        approval,
                        html_body=build_reply_html(approval),
                        cc_emails=cc
                    )
                    if sent:
                        update_action_row(i, status="Draft Pending",
                                          last_reminded=today.strftime("%d.%m.%Y %H:%M"),
                                          reminder_count=reminder_due)
                        logger.info(f"Draft {reminder_due} sent for approval: {data['action'][:50]}")

            except Exception as e:
                logger.error(f"External reminder row error: {e}")
                continue
    except Exception as e:
        logger.error(f"External reminder check error: {e}")


def check_thread_replies(thread_ids):
    replied_threads = set()
    if not thread_ids:
        return replied_threads
    mail = None
    try:
        mail = get_imap_connection()
        mail.select("INBOX")
        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            safe_logout(mail)
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
            except (imaplib.IMAP4.abort, OSError, EOFError) as fetch_err:
                logger.warning(f"Fetch error — reconnecting: {fetch_err}")
                safe_logout(mail)
                try:
                    mail = get_imap_connection()
                    mail.select("INBOX")
                except Exception:
                    break
                continue
            except:
                continue
        safe_logout(mail)
    except Exception as e:
        logger.error(f"Check replies error: {e}")
        if mail:
            safe_logout(mail)
    return replied_threads


def check_followup_reminders():
    logger.info("Checking internal follow-up reminders...")
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
                    notice = f"Dear Team,\n\nThe following email thread has been automatically closed after {AUTO_CLOSE_DAYS} days.\n\nSubject: {subject}\nFrom: {sender}\nDays open: {days_open}\n\nKind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    send_email(REPORT_RECIPIENTS, f"Auto-Closed — {subject}",
                               notice, html_body=build_reply_html(notice))
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
                    body  = f"Dear {first_name},\n\nI am writing to follow up on the email below, open for {days_open} days.\n\nSubject: {subject}\nDate raised: {date_str}\n\nSummary:\n{summary}\n\nAction required:\n{action}\n\nPlease review and respond at your earliest convenience.\n\nKind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    cc = [ALWAYS_CC] if sender.lower() != ALWAYS_CC.lower() else []
                elif reminder_due == 2:
                    subject_line = f"Second Follow-up — {subject}"
                    body  = f"Dear {first_name},\n\nSecond follow-up. This matter has been open for {days_open} days without a response.\n\nSubject: {subject}\nDate raised: {date_str}\n\nSummary:\n{summary}\n\nAction required:\n{action}\n\nThis requires your urgent attention.\n\nKind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    cc = [ALWAYS_CC] if sender.lower() != ALWAYS_CC.lower() else []
                elif reminder_due == 3:
                    subject_line = f"Escalation Notice — {subject}"
                    body  = f"Dear {first_name},\n\nFormal escalation. This matter has been open for {days_open} days despite two previous reminders.\n\nSubject: {subject}\nDate raised: {date_str}\n\nSummary:\n{summary}\n\nAction required:\n{action}\n\nIf no response within 7 days this will be auto-closed. Copied to management.\n\nKind regards,\n\nAlex Rivera\nConstruction Expert\nSCOPE Consulting MMC\ninternal@scope-iq.io"
                    cc = [r for r in REPORT_RECIPIENTS if r.lower() != sender.lower()]

                sent = send_email([sender], subject_line, body,
                                  cc_emails=cc,
                                  html_body=build_reply_html(body))
                if sent:
                    update_row(i, last_reminded=today.strftime("%d.%m.%Y %H:%M"),
                               reminder_count=reminder_due)

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
    results = extract_all_emails_with_names(header_value)
    return [r["email"] for r in results]


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
                    text = f"Excel: {filename}\nSheets: {', '.join(wb.sheetnames)}\n\n"
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
        logger.info(f"Sent to {to_emails} CC {cc_emails}")
        return True
    except Exception as e:
        logger.error(f"Resend error: {e}")
        return False


def analyse_email(sender, subject, body, attachments=None,
                  memory_files=None, is_cc=False):
    try:
        active_files = attachments if attachments else memory_files
        using_memory = not attachments and bool(memory_files)

        if is_cc:
            full_content = body
            if active_files:
                full_content += "\n\nFILES:\n" + "\n".join(
                    f"{f['name']}:\n{f['content']}" for f in active_files)
            prompt = f"""CC'd email — internal analysis only. Do not reply.
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
                    memory_note = f"Note: Using files previously received ({file_names}, saved {saved_at}).\n\n"
                files_content = "\n".join(
                    f"{f['name']}:\n{f['content']}" for f in active_files)
                prompt = f"""Email from SCOPE team member.
From: {sender}
Subject: {subject}
Email body: {body}

{memory_note}FILE CONTENT:
{files_content}

Analyse every discipline, every sheet, every line item equally.
For each item: description, quantity, unit rate, market assessment, Baku market range, quantity concerns.
Flag missing scope. Discipline risk summary. Total risk exposure at end.
Write complete formal professional reply now."""
            else:
                prompt = f"""Email from SCOPE team member.
From: {sender}
Subject: {subject}
Content: {body}
Write complete formal professional reply. If files needed and never provided, request them."""

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


def build_reply_html(body_text):
    today    = datetime.now().strftime("%d %B %Y")
    time_now = datetime.now().strftime("%H:%M")
    paragraphs = body_text.strip().split("\n\n")
    html_body  = ""

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
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
            lines     = para.split("\n")
            para_html = "<br>".join(line.strip() for line in lines if line.strip())
            html_body += f'<p style="font-size:14px;color:#333;line-height:1.8;margin:0 0 16px;">{para_html}</p>'

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
<div style="max-width:620px;margin:0 auto;padding:20px 0;">
  <div style="background:#1a2942;border-radius:12px 12px 0 0;padding:18px 28px;">
    <div style="display:flex;justify-content:space-between;align-items:center;">
      <div style="color:#fff;font-size:20px;font-weight:600;letter-spacing:1px;">SCOPE <span style="color:#3CB496;">IQ</span></div>
      <div style="font-size:11px;color:#8facc8;">{today} &nbsp;·&nbsp; {time_now} Baku</div>
    </div>
    <div style="font-size:12px;color:#8facc8;margin-top:6px;">Response from Alex Rivera &nbsp;·&nbsp; Construction Expert</div>
  </div>
  <div style="background:#fff;border:1px solid #e8e8e8;border-top:none;border-radius:0 0 12px 12px;padding:28px 28px 24px;">
    {html_body}
    <div style="background:#f8f9fa;border-radius:6px;padding:10px 14px;margin-top:20px;">
      <div style="font-size:11px;color:#888;">This response was prepared by <strong style="color:#1a2942;">Alex Rivera</strong>, Construction Expert at SCOPE Consulting MMC, using <strong style="color:#3CB496;">SCOPE IQ</strong>.</div>
    </div>
  </div>
</div>
</body></html>"""


def build_external_reminder_html(body_text, meeting_ref, action_item,
                                  responsible_label, due_date,
                                  reminder_label="Follow-up", on_behalf_of=None):
    """
    Branded HTML for external follow-up emails — matches the daily report style
    (navy header, stats-style bar, structured action card).
    """
    today    = datetime.now().strftime("%d %B %Y")
    time_now = datetime.now().strftime("%H:%M")

    paragraphs = body_text.strip().split("\n\n")
    html_body  = ""
    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
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
            html_body += f'<div style="margin-top:20px;padding-top:16px;border-top:1px solid #f0f0f0;line-height:1.8;">{sig_html}</div>'
        else:
            lines     = para.split("\n")
            para_html = "<br>".join(line.strip() for line in lines if line.strip())
            html_body += f'<p style="font-size:14px;color:#333;line-height:1.8;margin:0 0 16px;">{para_html}</p>'

    badge_color = "#f0a030" if reminder_label == "Follow-up" else \
                  "#e07030" if reminder_label == "Second Follow-up" else "#c00000"

    from_label = f"Alex Rivera on behalf of {on_behalf_of}" if on_behalf_of else "Alex Rivera"

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
<div style="max-width:620px;margin:0 auto;padding:20px 0;">

  <div style="background:#1a2942;border-radius:12px 12px 0 0;padding:24px 28px;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
      <div style="color:#fff;font-size:22px;font-weight:600;letter-spacing:1px;">SCOPE <span style="color:#3CB496;">IQ</span></div>
      <div style="background:{badge_color};color:#fff;font-size:11px;padding:4px 12px;border-radius:20px;font-weight:500;">{reminder_label}</div>
    </div>
    <table style="width:100%;font-size:12px;border-collapse:collapse;">
      <tr>
        <td style="color:#8facc8;padding:2px 0;width:50%;">Date &nbsp;<strong style="color:#c8ddf0;">{today}</strong></td>
        <td style="color:#8facc8;padding:2px 0;">Time &nbsp;<strong style="color:#c8ddf0;">{time_now} Baku</strong></td>
      </tr>
      <tr>
        <td style="color:#8facc8;padding:2px 0;">From &nbsp;<strong style="color:#c8ddf0;">{from_label}</strong></td>
        <td style="color:#8facc8;padding:2px 0;">Meeting ref &nbsp;<strong style="color:#c8ddf0;">{meeting_ref}</strong></td>
      </tr>
    </table>
  </div>

  <div style="background:#243550;padding:16px 28px;">
    <div style="font-size:11px;color:#8facc8;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Action item under follow-up</div>
    <div style="font-size:14px;color:#fff;font-weight:600;line-height:1.5;">{action_item}</div>
  </div>

  <div style="background:#fff;border:1px solid #e8e8e8;border-top:none;border-radius:0 0 12px 12px;padding:24px 28px;">

    <div style="border:1px solid #e8e8e8;border-radius:8px;margin-bottom:20px;overflow:hidden;">
      <div style="background:#f8f9fa;padding:10px 16px;border-bottom:1px solid #e8e8e8;">
        <span style="font-size:12px;color:#888;font-weight:500;">ACTION DETAILS</span>
      </div>
      <div style="padding:14px 16px;">
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          <tr><td style="color:#888;padding:4px 0;width:130px;">Responsible</td><td style="color:#1a2942;font-weight:600;">{responsible_label}</td></tr>
          <tr><td style="color:#888;padding:4px 0;">Due date</td><td style="color:#333;">{due_date}</td></tr>
          <tr><td style="color:#888;padding:4px 0;">Meeting reference</td><td style="color:#333;">{meeting_ref}</td></tr>
        </table>
      </div>
    </div>

    {html_body}

    <div style="background:#f8f9fa;border-radius:6px;padding:10px 14px;margin-top:16px;">
      <div style="font-size:11px;color:#888;line-height:1.6;">
        This follow-up was prepared by <strong style="color:#1a2942;">{from_label}</strong>,
        Construction Expert at SCOPE Consulting MMC, using
        <strong style="color:#3CB496;">SCOPE IQ</strong> — Intelligent Email Monitoring.
      </div>
    </div>

  </div>
</div>
</body>
</html>"""


def build_report_html(pending, closed, today, day_name, time_now,
                      open_actions=None, flagged=None):
    n_open       = len([r for r in pending if r.get("Status") == "Open"])
    n_monitor    = len([r for r in pending if r.get("Status") == "Monitoring"])
    n_closed     = len(closed)
    open_actions = open_actions or []
    flagged      = flagged or []
    total_open   = len(pending) + len(open_actions)

    items_html = ""
    for i, r in enumerate(pending[-15:], 1):
        subj    = r.get("Subject") or "No subject"
        sender  = r.get("Sender")  or "Unknown"
        date    = r.get("Date")    or "Not recorded"
        status  = r.get("Status")  or "Open"
        summary = r.get("Summary") or "No summary"
        action  = r.get("Action")  or "Review required"
        sbg = "#fff8ee" if status == "Open" else "#e1f5ee"
        stx = "#9a6000" if status == "Open" else "#0f6e56"
        items_html += f"""
        <div style="border:1px solid #e8e8e8;border-radius:8px;margin-bottom:16px;overflow:hidden;">
          <div style="background:#f8f9fa;padding:12px 16px;display:flex;justify-content:space-between;border-bottom:1px solid #e8e8e8;">
            <span style="font-size:12px;color:#888;">ITEM {i} OF {len(pending)}</span>
            <span style="background:{sbg};color:{stx};font-size:11px;padding:3px 10px;border-radius:20px;">{status}</span>
          </div>
          <div style="padding:16px;">
            <table style="width:100%;font-size:13px;border-collapse:collapse;">
              <tr><td style="color:#888;padding:4px 0;width:120px;">Subject</td><td style="color:#1a2942;font-weight:600;">{subj}</td></tr>
              <tr><td style="color:#888;padding:4px 0;">From</td><td style="color:#333;">{sender}</td></tr>
              <tr><td style="color:#888;padding:4px 0;">Date</td><td style="color:#333;">{date}</td></tr>
            </table>
            <div style="border-top:1px solid #f0f0f0;margin:12px 0;"></div>
            <div style="font-size:11px;color:#888;text-transform:uppercase;margin-bottom:4px;">Summary</div>
            <div style="font-size:13px;color:#444;line-height:1.6;margin-bottom:10px;">{summary}</div>
            <div style="background:#fff8ee;border:1px solid #f0c060;border-radius:6px;padding:10px 12px;">
              <div style="font-size:10px;font-weight:600;color:#9a6000;text-transform:uppercase;margin-bottom:4px;">Action required</div>
              <div style="font-size:13px;color:#5a3a00;">{action}</div>
            </div>
          </div>
        </div>"""

    ext_html = ""
    if open_actions:
        ext_html += """<div style="margin-top:24px;border-top:2px solid #e8e8e8;padding-top:20px;">
        <div style="font-size:11px;font-weight:600;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;">External action tracker</div>"""
        for i, r in enumerate(open_actions[-10:], 1):
            a = r.get("Action Item","") or "No description"
            p = r.get("Responsible Party","") or "Unknown"
            n = r.get("Responsible Name","") or ""
            e = r.get("Responsible Email","") or "Unknown"
            d = r.get("Due Date","") or "Not specified"
            s = r.get("Status","") or "Open"
            m = r.get("Meeting Reference","") or "Unknown"
            display = f"{n} ({p})" if n else p
            ext_html += f"""
            <div style="border:1px solid #e8e8e8;border-radius:8px;margin-bottom:12px;overflow:hidden;">
              <div style="background:#f8f9fa;padding:10px 14px;display:flex;justify-content:space-between;border-bottom:1px solid #e8e8e8;">
                <span style="font-size:12px;color:#888;">External Action {i}</span>
                <span style="background:#fff8ee;color:#9a6000;font-size:11px;padding:2px 8px;border-radius:20px;">{s}</span>
              </div>
              <div style="padding:14px;">
                <div style="font-size:13px;color:#1a2942;font-weight:600;margin-bottom:8px;">{a}</div>
                <table style="width:100%;font-size:12px;border-collapse:collapse;">
                  <tr><td style="color:#888;padding:3px 0;width:120px;">Responsible</td><td style="color:#333;">{display}</td></tr>
                  <tr><td style="color:#888;padding:3px 0;">Email</td><td style="color:#333;">{e}</td></tr>
                  <tr><td style="color:#888;padding:3px 0;">Due date</td><td style="color:#333;">{d}</td></tr>
                  <tr><td style="color:#888;padding:3px 0;">Meeting</td><td style="color:#333;">{m}</td></tr>
                </table>
              </div>
            </div>"""
        if flagged:
            ext_html += f"""<div style="background:#fff0f0;border:1px solid #f0c0c0;border-radius:6px;padding:10px 14px;">
              <div style="font-size:11px;font-weight:600;color:#c00000;margin-bottom:4px;">EMAIL UNKNOWN</div>
              <div style="font-size:12px;color:#800000;">{len(flagged)} action(s) unmatched. Please provide contact details.</div>
            </div>"""
        ext_html += "</div>"

    no_items = ""
    if not pending and not open_actions:
        no_items = """<div style="text-align:center;padding:32px;">
          <div style="width:48px;height:48px;border-radius:50%;background:#e1f5ee;margin:0 auto 12px;font-size:22px;color:#3CB496;display:flex;align-items:center;justify-content:center;">&#10003;</div>
          <div style="font-size:15px;color:#333;font-weight:500;">All clear</div>
          <div style="font-size:13px;color:#888;margin-top:4px;">No outstanding items as of today.</div>
        </div>"""

    greeting = f"Good morning. Daily report for <strong>{today}</strong>. <strong>{total_open} open item(s)</strong> require attention." if total_open else f"Good morning. Daily report for <strong>{today}</strong>. All items are clear."

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
<div style="max-width:620px;margin:0 auto;padding:20px 0;">
  <div style="background:#1a2942;border-radius:12px 12px 0 0;padding:24px 28px;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
      <div style="color:#fff;font-size:22px;font-weight:600;">SCOPE <span style="color:#3CB496;">IQ</span></div>
      <div style="background:#3CB496;color:#fff;font-size:11px;padding:4px 12px;border-radius:20px;">Daily Report</div>
    </div>
    <table style="width:100%;font-size:12px;border-collapse:collapse;">
      <tr>
        <td style="color:#8facc8;padding:2px 0;width:50%;">Date &nbsp;<strong style="color:#c8ddf0;">{day_name}, {today}</strong></td>
        <td style="color:#8facc8;padding:2px 0;">Time &nbsp;<strong style="color:#c8ddf0;">{time_now} Baku</strong></td>
      </tr>
      <tr>
        <td style="color:#8facc8;padding:2px 0;">Prepared by &nbsp;<strong style="color:#c8ddf0;">Alex Rivera</strong></td>
        <td style="color:#8facc8;padding:2px 0;">Status &nbsp;<strong style="color:#c8ddf0;">{"All clear" if total_open == 0 else f"{total_open} open"}</strong></td>
      </tr>
    </table>
  </div>
  <div style="background:#243550;padding:12px 28px;display:flex;">
    <div style="flex:1;text-align:center;border-right:1px solid #1a2942;">
      <div style="font-size:22px;font-weight:600;color:#f0a030;">{n_open}</div>
      <div style="font-size:11px;color:#8facc8;margin-top:2px;">Internal Open</div>
    </div>
    <div style="flex:1;text-align:center;border-right:1px solid #1a2942;">
      <div style="font-size:22px;font-weight:600;color:#3CB496;">{n_monitor}</div>
      <div style="font-size:11px;color:#8facc8;margin-top:2px;">Monitoring</div>
    </div>
    <div style="flex:1;text-align:center;border-right:1px solid #1a2942;">
      <div style="font-size:22px;font-weight:600;color:#e07030;">{len(open_actions)}</div>
      <div style="font-size:11px;color:#8facc8;margin-top:2px;">External Actions</div>
    </div>
    <div style="flex:1;text-align:center;">
      <div style="font-size:22px;font-weight:600;color:#6ab87a;">{n_closed}</div>
      <div style="font-size:11px;color:#8facc8;margin-top:2px;">Closed</div>
    </div>
  </div>
  <div style="background:#fff;border:1px solid #e8e8e8;border-top:none;border-radius:0 0 12px 12px;padding:24px 28px;">
    <p style="font-size:14px;color:#444;line-height:1.7;margin:0 0 20px;">{greeting}</p>
    {"<div style='font-size:11px;font-weight:600;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;'>Internal outstanding items</div>" if pending else ""}
    {items_html}
    {ext_html}
    {no_items}
    <div style="border-top:1px solid #f0f0f0;margin-top:24px;padding-top:20px;display:flex;justify-content:space-between;">
      <div style="font-size:12px;color:#666;line-height:1.8;">
        <strong style="color:#1a2942;font-size:13px;">Alex Rivera</strong><br>
        Construction Expert<br>SCOPE Consulting MMC<br>
        <span style="color:#3CB496;">internal@scope-iq.io</span>
      </div>
      <div style="font-size:11px;color:#aaa;text-align:right;line-height:1.7;">
        Generated automatically<br>SCOPE IQ<br>09:00 Baku daily
      </div>
    </div>
    <div style="background:#f8f9fa;border-radius:6px;padding:10px 14px;margin-top:16px;">
      <div style="font-size:11px;color:#888;">
        <strong style="color:#555;">Chase protocol:</strong>
        Draft at day 3, 7, 14 &nbsp;·&nbsp; Auto-close at day 21
      </div>
    </div>
  </div>
</div>
</body></html>"""


def process_emails():
    logger.info("Checking emails via IMAP...")
    load_processed_ids()
    mail = None
    try:
        mail = get_imap_connection()
        logger.info("IMAP login successful")
        mail.select("INBOX")

        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            safe_logout(mail)
            return

        all_ids = data[0].split()
        if not all_ids:
            safe_logout(mail)
            return

        recent    = all_ids[-50:]
        new_count = 0
        logger.info(f"Checking {len(recent)} emails")

        for eid in reversed(recent):
            try:
                eid_str = eid.decode() if isinstance(eid, bytes) else str(eid)

                try:
                    typ, msg_data = mail.fetch(eid, "(RFC822)")
                    if typ != "OK" or not msg_data or not msg_data[0]:
                        continue
                except (imaplib.IMAP4.abort, OSError, EOFError) as fetch_err:
                    logger.warning(f"Fetch error on {eid_str} — reconnecting: {fetch_err}")
                    safe_logout(mail)
                    try:
                        mail = get_imap_connection()
                        mail.select("INBOX")
                        typ, msg_data = mail.fetch(eid, "(RFC822)")
                        if typ != "OK" or not msg_data or not msg_data[0]:
                            continue
                    except Exception as reconnect_err:
                        logger.error(f"Reconnect failed: {reconnect_err}")
                        break

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
                in_reply_to  = safe_decode(msg.get("In-Reply-To", "")).strip()

                to_with_names  = extract_all_emails_with_names(msg.get("To",  ""))
                cc_with_names  = extract_all_emails_with_names(msg.get("CC",  ""))
                from_with_name = extract_email_with_name(msg.get("From", ""))

                to_addresses = [p["email"] for p in to_with_names]
                cc_addresses = [p["email"] for p in cc_with_names]

                if ZOHO_EMAIL.lower() in sender:
                    continue

                is_direct    = ZOHO_EMAIL.lower() in to_field
                is_cc_email  = ZOHO_EMAIL.lower() in cc_field
                is_internal  = any(t in sender for t in SCOPE_TEAM_EMAILS)
                is_external  = not is_internal and not is_internal_email(sender)

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
                    if is_internal and is_mom_email(subject, body, attachments):
                        logger.info(f"MOM in CC from internal — extracting actions")
                        all_with_names = []
                        seen = set()
                        for p in [from_with_name] + to_with_names + cc_with_names:
                            if p["email"] and p["email"] not in seen:
                                seen.add(p["email"])
                                all_with_names.append(p)
                        process_mom_email(sender, subject, body,
                                          attachments, all_with_names, msg_id_hdr)
                    else:
                        logger.info(f"CC — silent log: {sender}")
                        analysis = analyse_email(sender, subject, body,
                                                 attachments=attachments,
                                                 memory_files=memory_files,
                                                 is_cc=True)
                        if analysis:
                            save_to_monitoring(sender, subject,
                                               analysis[:400], analysis[:500],
                                               msg_id_hdr, "Monitoring")

                elif is_direct and is_internal:
                    if is_mom_email(subject, body, attachments):
                        logger.info(f"MOM direct from internal — extracting actions")
                        all_with_names = []
                        seen = set()
                        for p in [from_with_name] + to_with_names + cc_with_names:
                            if p["email"] and p["email"] not in seen:
                                seen.add(p["email"])
                                all_with_names.append(p)
                        process_mom_email(sender, subject, body,
                                          attachments, all_with_names, msg_id_hdr)
                        save_to_memory(sender, subject,
                                       "MOM processed — see Action Tracker",
                                       "Review Action Tracker", "Closed")
                    else:
                        analysis = analyse_email(sender, subject, body,
                                                 attachments=attachments,
                                                 memory_files=memory_files,
                                                 is_cc=False)
                        if analysis:
                            reply_sub = f"Re: {subject}" if not subject.startswith("Re:") else subject
                            all_recipients = list(set(
                                [sender] +
                                [a for a in to_addresses if a != ZOHO_EMAIL.lower()] +
                                [a for a in cc_addresses if a != ZOHO_EMAIL.lower()]
                            ))
                            new_references = f"{references} {msg_id_hdr}".strip() if references else msg_id_hdr
                            sent = send_email(
                                all_recipients, reply_sub, analysis,
                                reply_to_msg_id=msg_id_hdr,
                                references=new_references,
                                html_body=build_reply_html(analysis)
                            )
                            save_to_memory(sender, subject, analysis[:400],
                                           "Replied by Alex",
                                           "Closed" if sent else "Open")

                elif is_direct and is_external:
                    logger.info(f"External direct to Alex: {sender}")
                    matched_action = find_action_by_thread(in_reply_to, references)

                    if matched_action:
                        logger.info(f"Matched action — routing for internal approval")
                        route_external_reply_for_approval(
                            sender, subject, body, matched_action, msg_id_hdr
                        )
                    else:
                        logger.info(f"Unknown external direct — logging silently")
                        save_to_monitoring(
                            sender, subject,
                            f"External email received from {sender}. No matching tracked action found.",
                            "Review if action required",
                            msg_id_hdr, "Monitoring"
                        )

            except Exception as e:
                logger.error(f"Email error: {e}")
                continue

        safe_logout(mail)
        logger.info(f"Done — {new_count} new emails processed")

    except imaplib.IMAP4.error as e:
        logger.error(f"IMAP auth error: {e}")
        if mail:
            safe_logout(mail)
    except Exception as e:
        logger.error(f"IMAP error: {e}")
        if mail:
            safe_logout(mail)


def send_morning_report():
    logger.info("Sending morning report...")
    try:
        pending, closed          = read_memory_for_report()
        open_actions, _, flagged = read_actions_for_report()
        today    = datetime.now().strftime("%d %B %Y")
        day_name = datetime.now().strftime("%A")
        time_now = datetime.now().strftime("%H:%M")

        html_body  = build_report_html(pending, closed, today, day_name, time_now,
                                       open_actions=open_actions, flagged=flagged)
        total_open = len(pending) + len(open_actions)
        plain      = f"SCOPE IQ Daily Report — {today}\nInternal: {len(pending)} | External: {len(open_actions)}\nAlex Rivera | SCOPE Consulting MMC"

        subject_line  = f"SCOPE IQ Daily Report — {today}"
        subject_line += f" — {total_open} Open Item(s)" if total_open else " — All Clear"

        send_email(REPORT_RECIPIENTS, subject_line, plain, html_body=html_body)
        logger.info("Morning report sent")
    except Exception as e:
        logger.error(f"Morning report error: {e}")


def main():
    logger.info("Alex Email Service starting")
    logger.info(f"Monitoring: {ZOHO_EMAIL}")
    logger.info(f"Internal domain: {INTERNAL_DOMAIN}")
    logger.info(f"Client domains: {CLIENT_DOMAINS}")
    logger.info(f"Authorised team: {SCOPE_TEAM_EMAILS}")
    logger.info(f"Chase protocol: Day {REMINDER_1_DAYS}/{REMINDER_2_DAYS}/{REMINDER_3_DAYS}/Close {AUTO_CLOSE_DAYS}")

    load_processed_ids()
    logger.info("Duplicate protection active")

    schedule.every(10).minutes.do(process_emails)
    schedule.every().day.at("05:00").do(send_morning_report)
    schedule.every(6).hours.do(check_followup_reminders)
    schedule.every(6).hours.do(check_external_action_replies)
    schedule.every(6).hours.do(check_action_approvals)
    schedule.every(6).hours.do(check_external_action_reminders)

    process_emails()

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
