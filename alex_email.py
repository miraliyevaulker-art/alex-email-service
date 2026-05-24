import os
import json
import time
import imaplib
import smtplib
import email
import schedule
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
from datetime import datetime
from anthropic import Anthropic
import gspread
from google.oauth2.service_account import Credentials

ZOHO_EMAIL        = os.environ.get("ZOHO_EMAIL", "internal@scope-iq.io")
ZOHO_APP_PASSWORD = os.environ.get("ZOHO_APP_PASSWORD")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
GOOGLE_SHEET_ID   = "1i-DZghVlJdLdWUB4jDjCWU_5-1VSzHwgpjZiVUz0-fg"
GOOGLE_CREDS_FILE = "primordial-mile-495807-k9-0217981265dd.json"
MODEL             = "claude-haiku-4-5-20251001"

IMAP_HOST = "imappro.zoho.com"
IMAP_PORT = 993
SMTP_HOST = "smtppro.zoho.com"
SMTP_PORT = 465

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

anthropic_client   = Anthropic(api_key=ANTHROPIC_API_KEY)
processed_ids_file = "/tmp/processed_emails.json"

SYSTEM_PROMPT = """You are Alex Rivera — senior Construction Expert at SCOPE Consulting MMC.

25+ years across commercial, residential, infrastructure, oil & gas, and high-end fit-out projects in Europe, Middle East, and CIS countries including Azerbaijan.

YOUR EXPERTISE:
- Quantity Surveying and Cost Management
- Quality Assurance and Quality Control (QA/QC)
- Contract Administration (FIDIC, NEC4, AzDTN)
- Technical document review — all disciplines
- Materials and specifications
- Handover and commissioning
- Claims and variations
- Eurocodes, British Standards, ISO, GOST, AzDTN, SNiP standards

LANGUAGE RULES:
1. English email → English reply ONLY
2. Azerbaijani email → Azerbaijani reply ONLY
3. Mixed → dominant language
4. NEVER mix languages

AZERBAIJANI: ə ı ö ü ğ ş ç — mandatory. AzDTN/GOST terminology. Formal register.

EMAIL REPLY FORMAT:
Professional, structured, concise.
Reference document numbers where relevant.
Clear action items with deadlines.

Signature:
Alex Rivera
Construction Expert | SCOPE Consulting MMC
internal@scope-iq.io

QA/QC: MAR → ✅ APPROVED / ⚠️ WITH COMMENTS / ❌ REJECTED
BOQ: Every position vs Baku market rates
FIDIC/NEC4: Contractually correct always"""


def get_sheet(tab="Sheet1"):
    try:
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
        client = gspread.authorize(creds)
        return client.open_by_key(GOOGLE_SHEET_ID).worksheet(tab)
    except Exception as e:
        logger.error(f"Sheet error: {e}")
        return None


def save_to_memory(sender, subject, summary, action, status="Open"):
    try:
        sheet = get_sheet("Sheet1")
        if sheet:
            sheet.append_row([
                datetime.now().strftime("%d.%m.%Y %H:%M"),
                sender, subject, summary, action, status
            ])
            logger.info(f"Saved to memory: {subject}")
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


def load_processed_ids():
    try:
        if os.path.exists(processed_ids_file):
            with open(processed_ids_file) as f:
                return set(json.load(f))
    except:
        pass
    return set()


def save_processed_id(msg_id):
    try:
        ids = load_processed_ids()
        ids.add(str(msg_id))
        with open(processed_ids_file, "w") as f:
            json.dump(list(ids)[-500:], f)
    except Exception as e:
        logger.error(f"Save ID error: {e}")


def safe_decode(value, fallback=""):
    """Safely decode email header value"""
    if value is None:
        return fallback
    try:
        if isinstance(value, bytes):
            return value.decode("utf-8", errors="replace")
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


def get_email_body(msg):
    """Extract plain text body from email"""
    body = ""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                try:
                    if part.get_content_type() == "text/plain":
                        payload = part.get_payload(decode=True)
                        if payload:
                            body += payload.decode(
                                part.get_content_charset() or "utf-8",
                                errors="replace"
                            )
                except:
                    continue
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="replace")
    except Exception as e:
        logger.error(f"Body extraction error: {e}")
    return body[:3000]


def send_reply(to_email, subject, body, reply_to_msg_id=None):
    """Send email via SMTP"""
    try:
        msg = MIMEMultipart()
        msg["From"]    = ZOHO_EMAIL
        msg["To"]      = to_email
        msg["Subject"] = subject
        if reply_to_msg_id:
            msg["In-Reply-To"] = str(reply_to_msg_id)
            msg["References"]  = str(reply_to_msg_id)
        msg.attach(MIMEText(body, "plain", "utf-8"))

        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
            server.send_message(msg)
            logger.info(f"✅ Reply sent to {to_email}")
            return True
    except Exception as e:
        logger.error(f"SMTP error: {e}")
        return False


def analyse_email(sender, subject, body, is_cc=False):
    """Analyse email with Claude"""
    try:
        if is_cc:
            prompt = f"""CC'd email — internal analysis only. Do NOT write a reply.

From: {sender}
Subject: {subject}
Content: {body}

Provide:
1. Type: MAR/RFI/BOQ/Variation/Claim/NCR/General
2. Key points (2-3 lines)
3. Action needed from SCOPE team
4. Risk: High/Medium/Low
5. Response deadline"""
        else:
            prompt = f"""Email from SCOPE team member — write professional reply.

From: {sender}
Subject: {subject}
Content: {body}

Write complete professional reply. Match language exactly."""

        response = anthropic_client.messages.create(
            model=MODEL,
            max_tokens=1500,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        logger.error(f"Analysis error: {e}")
        return None


def process_emails():
    logger.info("Checking emails via IMAP...")
    processed_ids = load_processed_ids()

    try:
        # Connect to IMAP
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
        logger.info("✅ IMAP login successful")
        mail.select("INBOX")

        # Search all emails
        typ, data = mail.search(None, "ALL")
        if typ != "OK":
            logger.error("Search failed")
            mail.logout()
            return

        # Handle empty inbox
        if not data or not data[0]:
            logger.info("Inbox empty")
            mail.logout()
            return

        all_ids = data[0].split()
        if not all_ids:
            logger.info("No emails found")
            mail.logout()
            return

        # Process last 25 emails
        recent = all_ids[-25:]
        logger.info(f"Checking {len(recent)} recent emails")
        new_count = 0

        for eid in reversed(recent):
            try:
                eid_str = eid.decode() if isinstance(eid, bytes) else str(eid)

                # Skip already processed
                if eid_str in processed_ids:
                    continue

                # Fetch email
                typ, msg_data = mail.fetch(eid, "(RFC822)")
                if typ != "OK" or not msg_data or not msg_data[0]:
                    save_processed_id(eid_str)
                    continue

                raw = msg_data[0][1]
                if not raw:
                    save_processed_id(eid_str)
                    continue

                msg = email.message_from_bytes(raw)

                # Extract headers safely
                sender   = safe_decode(msg.get("From"),    "").lower()
                subject  = safe_decode(msg.get("Subject"), "No subject")
                to_field = safe_decode(msg.get("To"),      "").lower()
                cc_field = safe_decode(msg.get("CC"),      "").lower()
                msg_hdr  = safe_decode(msg.get("Message-ID"), "")

                # Skip Alex's own sent emails
                if ZOHO_EMAIL.lower() in sender:
                    save_processed_id(eid_str)
                    continue

                is_direct   = ZOHO_EMAIL.lower() in to_field
                is_cc_email = ZOHO_EMAIL.lower() in cc_field
                is_internal = any(t in sender for t in SCOPE_TEAM_EMAILS)

                # Skip if not relevant
                if not is_direct and not is_cc_email:
                    save_processed_id(eid_str)
                    continue

                new_count += 1
                body = get_email_body(msg)
                logger.info(f"Processing: {sender} | {subject}")

                if is_direct and is_internal:
                    # Internal SCOPE team → reply
                    analysis = analyse_email(sender, subject, body, is_cc=False)
                    if analysis:
                        reply_sub = f"Re: {subject}" if not subject.startswith("Re:") else subject
                        sent = send_reply(sender, reply_sub, analysis, msg_hdr)
                        save_to_memory(
                            sender, subject, analysis[:300],
                            "Replied by Alex",
                            "Closed" if sent else "Open"
                        )

                elif is_direct and not is_internal:
                    # External direct → ignore
                    logger.info(f"Ignoring external direct: {sender}")

                elif is_cc_email:
                    # CC'd → silent log
                    analysis = analyse_email(sender, subject, body, is_cc=True)
                    if analysis:
                        save_to_memory(
                            sender, subject, analysis[:300],
                            "Logged — action required",
                            "Monitoring"
                        )

                save_processed_id(eid_str)

            except Exception as e:
                logger.error(f"Email error: {e}")
                try:
                    save_processed_id(eid_str)
                except:
                    pass
                continue

        mail.logout()
        logger.info(f"✅ Done — {new_count} new emails processed")

    except imaplib.IMAP4.error as e:
        logger.error(f"IMAP auth error: {e}")
    except Exception as e:
        logger.error(f"IMAP connection error: {e}")


def send_morning_report():
    logger.info("Sending morning report...")
    try:
        pending, closed = read_memory_for_report()
        today    = datetime.now().strftime("%d.%m.%Y")
        day_name = datetime.now().strftime("%A")
        day_az   = {
            "Monday":    "Bazar ertəsi",
            "Tuesday":   "Çərşənbə axşamı",
            "Wednesday": "Çərşənbə",
            "Thursday":  "Cümə axşamı",
            "Friday":    "Cümə",
            "Saturday":  "Şənbə",
            "Sunday":    "Bazar"
        }.get(day_name, day_name)

        report  = f"SCOPE IQ — Günlük E-poçt Hesabatı\n"
        report += f"{day_az}, {today}\n"
        report += "=" * 45 + "\n\n"

        if pending:
            report += f"CAVAB GÖZLƏYƏN / MONİTORİNQ: {len(pending)}\n\n"
            for r in pending[-15:]:
                subj   = r.get("Subject") or r.get("Topic") or "?"
                sender = r.get("Sender") or r.get("Project") or "?"
                date   = r.get("Date") or ""
                status = r.get("Status") or ""
                action = r.get("Action") or ""
                report += f"— {subj}\n"
                report += f"  {sender} | {date} | {status}\n"
                report += f"  {action}\n\n"
        else:
            report += "Gözləyən e-poçt yoxdur. ✅\n\n"

        report += "=" * 45 + "\n"
        report += "Alex Rivera\n"
        report += "Construction Expert | SCOPE Consulting MMC\n"
        report += "internal@scope-iq.io"

        for recipient in REPORT_RECIPIENTS:
            send_reply(
                recipient,
                f"SCOPE IQ — Günlük Hesabat — {today}",
                report
            )
        logger.info("✅ Morning report sent")

    except Exception as e:
        logger.error(f"Morning report error: {e}")


def main():
    logger.info("=" * 50)
    logger.info("Alex Email Service starting...")
    logger.info(f"Email: {ZOHO_EMAIL}")
    logger.info(f"IMAP: {IMAP_HOST}:{IMAP_PORT}")
    logger.info(f"Team: {SCOPE_TEAM_EMAILS}")
    logger.info(f"Reports to: {REPORT_RECIPIENTS}")
    logger.info("=" * 50)

    # Check every 5 minutes
    schedule.every(5).minutes.do(process_emails)

    # Morning report 9:00 AM Baku = 05:00 UTC
    schedule.every().day.at("05:00").do(send_morning_report)

    # Run immediately on start
    process_emails()

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
