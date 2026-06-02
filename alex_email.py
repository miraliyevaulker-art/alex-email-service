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

Civil works: review every excavation, concrete, formwork, reinforcement, and masonry position individually. For each state the item description, quantity submitted, unit rate submitted, whether the rate is within Baku market range, above market or below market, the Baku market range, and any quantity concerns.

Fit-out works: review every partition, ceiling, flooring, tiling, painting, joinery, and door position individually with the same depth.

MEP works — IDENTICAL DEPTH TO CIVIL AND FIT-OUT — NEVER SUMMARISE AS A BLOCK:
Mechanical: review every ductwork item, fan coil unit, chiller, AHU, VAV box, grille, diffuser, insulation run, pipework run, pump, valve, and control item individually.
Electrical: review every cable run, cable tray, distribution board, lighting fixture, emergency light, socket, switch, earthing conductor, and containment item individually.
Fire protection: review every sprinkler head, sprinkler pipe run, fire pump, alarm panel, smoke detector, heat detector, and suppression item individually.
Low voltage and ICT: review every access control point, CCTV camera, structured cabling run, server rack, and equipment item individually.
Plumbing: review every pipework run, fitting, sanitary fixture, tap, valve, and drainage item individually.

For every single line item across all disciplines state: item description, quantity submitted, unit rate submitted, market assessment, Baku market range, and quantity concerns if any.

Flag all missing scope items across all disciplines.
Provide a discipline-by-discipline cost risk summary and a combined total risk exposure at the end.

SIGNATURE — always end every email with exactly this:

Alex Rivera
Construction Expert
SCOPE Consulting MMC
internal@scope-iq.io

BAKU MARKET RATES:
Mechanical excavation 8 to 15 AZN per cubic metre, Manual excavation 60 to 90 AZN per cubic metre, Concrete C25/30 190 to 240 AZN per cubic metre, Reinforcement 1200 to 1500 AZN per tonne, Formwork 18 to 28 AZN per square metre, External brickwork 25 to 40 AZN per square metre, Internal brickwork 20 to 32 AZN per square metre, Plastering 16 to 24 AZN per square metre, Paint 12 to 18 AZN per square metre, Ceramic tiles 25 to 40 AZN per square metre, Premium tiles 45 to 80 AZN per square metre, Gypsum partition 32 to 48 AZN per square metre, Armstrong ceiling 28 to 42 AZN per square metre, Raised access floor 55 to 85 AZN per square metre, Epoxy floor 35 to 55 AZN per square metre, Carpet tiles 40 to 70 AZN per square metre, Aluminium glazing 180 to 280 AZN per square metre, Timber door 350 to 550 AZN each, Fire door 600 to 1200 AZN each, Metal door 4500 to 6500 AZN each, Aluminium door 700 to 900 AZN each, HVAC ductwork 45 to 75 AZN per square metre, Fan coil unit 350 to 600 AZN each, Chiller 120 to 200 AZN per kW, AHU 800 to 2500 AZN each, VAV box 250 to 600 AZN each, Grille and diffuser 35 to 85 AZN each, Duct insulation 15 to 28 AZN per square metre, Plumbing pipework 25 to 55 AZN per metre, Sanitary fixture 180 to 450 AZN each, Pump 800 to 3500 AZN each, Cable tray 35 to 65 AZN per metre, LV cable 8 to 25 AZN per metre, Distribution board 800 to 3500 AZN each, Lighting fixture 45 to 120 AZN each, Emergency lighting 80 to 180 AZN each, Socket and switch 25 to 65 AZN each, Fire alarm panel 1500 to 8000 AZN each, Smoke detector 45 to 120 AZN each, Sprinkler head 25 to 55 AZN each, Sprinkler pipework 18 to 45 AZN per metre, Fire pump 3500 to 12000 AZN each, Access control 800 to 2500 AZN per door, CCTV camera 250 to 600 AZN each, Structured cabling point 85 to 180 AZN each, Concrete paving 35 to 55 AZN per square metre, Natural stone paving 85 to 150 AZN per square metre, Soft landscape 25 to 55 AZN per square metre, Site fencing 45 to 120 AZN per metre, Passenger lift 45000 to 80000 AZN each, Freight lift 60000 to 120000 AZN each."""


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
    global _file_memory_cache
    sender = sender.strip().lower()
    _file_memory_cache[sender] = [
        {
            "name":     att["name"],
            "content":  att["content"],
            "saved_at": datetime.now().strftime("%d.%m.%Y %H:%M")
        }
        for att in attachments
    ]
    logger.info(f"Saved {len(attachments)} files to memory for {sender}")
    try:
        sheet = get_or_create_sheet("File Memory", rows=2000, cols=4)
        if not sheet:
            return
        all_values = sheet.get_all_values()
        if not all_values:
            sheet.append_row(["Sender", "Filename", "Content", "Saved At"])
            all_values = [["Sender", "Filename", "Content", "Saved At"]]
        rows_to_delete = []
        for i, row in enumerate(all_values):
            if i == 0:
                continue
            if row and len(row) > 0 and row[0].strip().lower() == sender:
                rows_to_delete.append(i + 1)
        for row_num in reversed(rows_to_delete):
            sheet.delete_rows(row_num)
        saved_at = datetime.now().strftime("%d.%m.%Y %H:%M")
        for att in attachments:
            sheet.append_row([sender, att["name"], att["content"][:10000], saved_at])
        logger.info(f"Persisted {len(attachments)} files to sheet for {sender}")
    except Exception as e:
        logger.error(f"Save files to sheet error: {e}")


def load_files_from_memory(sender):
    global _file_memory_cache
    sender = sender.strip().lower()
    if sender in _file_memory_cache and _file_memory_cache[sender]:
        files = _file_memory_cache[sender]
        logger.info(f"Loaded {len(files)} files from cache for {sender}")
        return files
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
            _file_memory_cache[sender] = files
            logger.info(f"Loaded {len(files)} files from sheet for {sender}")
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
                    text = f"Excel file: {filename}\n"
                    text += f"Sheets: {', '.join(wb.sheetnames)}\n\n"
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        text += f"\n{'='*40}\nSHEET: {sheet_name}\n{'='*40}\n"
                        for row in ws.iter_rows(max_row=500, values_only=True):
                            row_data = [str(c) for c in row if c is not None]
                            if row_data:
                                text += " | ".join(row_data) + "\n"
                    attachments.append({"name": filename, "content": text[:15000]})
                    logger.info(f"Extracted Excel: {filename} — {len(wb.sheetnames)} sheets")
                except Exception as e:
                    logger.error(f"Excel error: {e}")

            elif ext == "pdf":
                try:
                    import fitz
                    doc  = fitz.open(stream=payload, filetype="pdf")
                    text = f"PDF: {filename}\n"
                    for page in doc:
                        text += page.get_text()
                    doc.close()
                    attachments.append({"name": filename, "content": text[:5000]})
                except Exception as e:
                    logger.error(f"PDF error: {e}")

            elif ext in ["docx", "doc"]:
                try:
                    import docx
                    document = docx.Document(io.BytesIO(payload))
                    text     = f"Word: {filename}\n"
                    text    += "\n".join([p.text for p in document.paragraphs])
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
                            charset = part.get_content_charset() or "utf-8"
                            body   += payload.decode(charset, errors="replace")
                except:
                    continue
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                body    = payload.decode(charset, errors="replace")
    except Exception as e:
        logger.error(f"Body error: {e}")
    return body[:2000]


def send_email(to_emails, subject, body, reply_to_msg_id=None, references=None):
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
        if reply_to_msg_id:
            params["headers"]["In-Reply-To"] = reply_to_msg_id
            params["headers"]["References"]  = references or reply_to_msg_id
        resend.Emails.send(params)
        logger.info(f"Email sent to {to_emails}")
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
                full_content += "\n\nFILES:\n"
                for f in active_files:
                    full_content += f"\n{f['name']}:\n{f['content']}\n"
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
                    saved_at   = active_files[0].get("saved_at", "previously") if active_files else ""
                    file_names = ", ".join([f["name"] for f in active_files])
                    memory_note = f"Note: No new files in this email. Using files previously received from this sender ({file_names}, saved {saved_at}).\n\n"

                files_content = chr(10).join([
                    f"{f['name']}:\n{f['content']}" for f in active_files
                ])

                prompt = f"""Email from SCOPE team member.

From: {sender}
Subject: {subject}
Email body: {body}

{memory_note}FILE CONTENT — ALL SHEETS:
{files_content}

Analyse every discipline, every sheet, and every single line item with equal depth.
Civil, fit-out, and all MEP disciplines — each line item reviewed individually. Never summarise any discipline as a block.
For each line item: description, quantity, unit rate, market assessment, Baku market range, quantity concerns.
MEP receives identical line-by-line depth as civil and fit-out. Every ductwork item, cable run, fixture, pipe, sprinkler head — individually assessed.
Flag all missing scope. Discipline-by-discipline risk summary. Combined total risk exposure at end.
Write complete formal professional reply with full analysis now."""

            else:
                prompt = f"""Email from SCOPE team member. No files attached and no files previously received from this sender.

From: {sender}
Subject: {subject}
Content: {body}

Write complete formal professional reply.
If document analysis is required and no files have ever been provided, politely request them and confirm reply within minutes of receipt."""

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
            logger.info("Inbox empty")
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
                    logger.info(f"New files from {sender}: {[a['name'] for a in attachments]}")
                    save_files_to_memory(sender, attachments)
                    memory_files = None
                else:
                    memory_files = load_files_from_memory(sender)
                    if memory_files:
                        logger.info(f"Using {len(memory_files)} memory files for {sender}")

                logger.info(f"Processing: {sender} | {subject}")

                # CC'd email — silent analysis only, NEVER reply
                if is_cc_email:
                    logger.info(f"CC'd email — silent log only, no reply: {sender}")
                    analysis = analyse_email(
                        sender, subject, body,
                        attachments=attachments,
                        memory_files=memory_files,
                        is_cc=True
                    )
                    if analysis:
                        save_to_memory(
                            sender, subject,
                            analysis[:400], analysis[:500],
                            "Monitoring"
                        )

                # Direct email from internal team — reply
                elif is_direct and is_internal:
                    analysis = analyse_email(
                        sender, subject, body,
                        attachments=attachments,
                        memory_files=memory_files,
                        is_cc=False
                    )
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
                            references=new_references
                        )
                        save_to_memory(
                            sender, subject, analysis[:400],
                            "Replied by Alex",
                            "Closed" if sent else "Open"
                        )

                # Direct email from external — ignore completely
                elif is_direct and not is_internal:
                    logger.info(f"Ignoring external direct: {sender}")

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
        today = datetime.now().strftime("%d %B %Y")

        if pending:
            report  = "Dear Team,\n\n"
            report += f"Good morning. Please find below a detailed summary of outstanding emails and open action items as of {today}.\n\n"
            report += "=" * 60 + "\n\n"
            for i, r in enumerate(pending[-15:], 1):
                subj    = r.get("Subject") or r.get("Topic")   or "No subject"
                sender  = r.get("Sender")  or r.get("Project") or "Unknown"
                date    = r.get("Date")    or "Not recorded"
                status  = r.get("Status")  or "Open"
                summary = r.get("Summary") or "No summary available"
                action  = r.get("Action")  or "Review and action required"
                report += f"Item {i}\n\n"
                report += f"Subject: {subj}\n"
                report += f"Date received: {date}\n"
                report += f"From: {sender}\n"
                report += f"Status: {status}\n\n"
                report += f"Content summary:\n{summary}\n\n"
                report += f"Action required:\n{action}\n\n"
                report += "-" * 40 + "\n\n"
            report += "Please review the above items and ensure the necessary actions are taken at your earliest convenience.\n\n"
        else:
            report  = "Dear Team,\n\n"
            report += f"Good morning. As of {today}, there are no outstanding emails or open action items requiring your attention.\n\n"
            report += "Should you have any queries or wish to submit documents for review, please send them directly to this address.\n\n"

        report += "Kind regards,\n\n"
        report += "Alex Rivera\n"
        report += "Construction Expert\n"
        report += "SCOPE Consulting MMC\n"
        report += "internal@scope-iq.io"

        send_email(
            REPORT_RECIPIENTS,
            f"SCOPE IQ Daily Report — {today}",
            report
        )
        logger.info("Morning report sent to all team")

    except Exception as e:
        logger.error(f"Morning report error: {e}")


def main():
    logger.info("Alex Email Service starting")
    logger.info(f"Monitoring: {ZOHO_EMAIL}")
    logger.info(f"Authorised team: {SCOPE_TEAM_EMAILS}")
    logger.info(f"Report recipients: {REPORT_RECIPIENTS}")

    load_processed_ids()
    logger.info("Duplicate protection active")

    schedule.every(10).minutes.do(process_emails)
    schedule.every().day.at("05:00").do(send_morning_report)

    process_emails()

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
