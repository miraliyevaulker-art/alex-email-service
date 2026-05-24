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

# Configuration
ZOHO_EMAIL        = os.environ.get("ZOHO_EMAIL", "internal@scope-iq.io")
ZOHO_APP_PASSWORD = os.environ.get("ZOHO_APP_PASSWORD")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
GOOGLE_SHEET_ID   = "1i-DZghVlJdLdWUB4jDjCWU_5-1VSzHwgpjZiVUz0-fg"
GOOGLE_CREDS_FILE = "primordial-mile-495807-k9-0217981265dd.json"
MODEL             = "claude-haiku-4-5-20251001"

# Zoho IMAP/SMTP settings
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

anthropic_client = Anthropic(api_key=ANTHROPIC_API_KEY)
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

WHO YOU ARE:
A senior construction professional. Real colleague. Not a robot.
You adapt to any project, any contractor, any vendor, any discipline.

LANGUAGE RULES — CRITICAL:
1. English email → respond in English ONLY
2. Azerbaijani email → respond in Azerbaijani ONLY
3. Mixed → dominant language
4. NEVER mix languages

AZERBAIJANI QUALITY:
1. MANDATORY: ə ı ö ü ğ ş ç İ Ə Ö Ü Ğ Ş Ç
2. NEVER: ə→e, ı→i, ğ→g, ş→s, ç→c
3. AzDTN/GOST standard terminology
4. Formal register always

EMAIL REPLY FORMAT:
- Professional, concise, structured
- Reference document numbers and clause numbers
- Clear action items with deadlines
- End with signature:

Alex Rivera
Construction Expert | SCOPE Consulting MMC
internal@scope-iq.io

QA/QC:
MAR → ✅ APPROVED / ⚠️ APPROVED WITH COMMENTS / ❌ REJECTED
Azerbaijani → ✅ TƏSDİQLƏNDİ / ⚠️ ŞƏRHLƏ TƏSDİQLƏNDİ / ❌ RƏDD EDİLDİ

BOQ: Every position vs Baku market rates — flag HIGH/LOW/OK
FIDIC/NEC4: Contractually correct responses always

BAKU MARKET RATES:
TORPAQ: Mexaniki qazıntı 8-15 AZN/m³, Əl qazıntı 60-90 AZN/m³, Doldurma 8-12 AZN/m³
BETON: C25/30 190-240 AZN/m³, Armatur 1200-1500 AZN/ton, Qəlib 18-28 AZN/m²
HÖRGÜ: Kərpic xarici 25-40 AZN/m², Kərpic daxili 20-32 AZN/m², Suvaq 16-24 AZN/m²
BƏZƏYİŞ: Kafel 25-40 AZN/m², Alçipan 32-48 AZN/m², Armstrong 28-42 AZN/m², Boya 12-18 AZN/m²
MEP: Hava kanalları 45-75 AZN/m², Fankoyl 350-600 AZN/ədəd, İşıq 45-120 AZN/ədəd, Sprinkler 45-75 AZN/m²
QAPI: Metal qapı 4500-6500 AZN/ədəd, Alüminium qapı 700-900 AZN/ədəd"""


def get_sheet(tab="Sheet1"):
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        google_creds_json = os.environ.get("GOOGLE_CREDENTIALS")
        if google_creds_json:
            creds = Credentials.from_service_account_info(
                json.loads(google_creds_json), scopes=scopes)
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
            date = datetime.now().strftime("%d.%m.%Y %H:%M")
            sheet.append_row([date, sender, subject, summary, action, status])
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


def load_processed_ids():
    try:
        if os.path.exists(processed_ids_file):
            with open(processed_ids_file) as f:
                return set(json.load(f))
        return set()
    except:
        return set()


def save_processed_id(msg_id):
    try:
        ids = load_processed_ids()
        ids.add(str(msg_id))
        with open(processed_ids_file, "w") as f:
            json.dump(list(ids)[-500:], f)
    except Exception as e:
        logger.error(f"Save ID error: {e}")


def decode_str(s):
    if s is None:
        return ""
    decoded = decode_header(s)
    result = ""
    for part, enc in decoded:
        if isinstance(part, bytes):
            result += part.decode(enc or "utf-8", errors="replace")
        else:
            result += str(part)
    return result


def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    body += part.get_payload(decode=True).decode("utf-8", errors="replace")
                except:
                    pass
    else:
        try:
            body = msg.get_payload(decode=True).decode("utf-8", errors="replace")
        except:
            pass
    return body[:3000]


def send_reply(to_email, subject, body, reply_to_msg_id=None):
    try:
        msg = MIMEMultipart()
        msg["From"]    = ZOHO_EMAIL
        msg["To"]      = to_email
        msg["Subject"] = subject
        if reply_to_msg_id:
            msg["In-Reply-To"] = reply_to_msg_id
            msg["References"]  = reply_to_msg_id
        msg.attach(MIMEText(body, "plain", "utf-8"))

        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
            server.send_message(msg)
            logger.info(f"Reply sent to {to_email} ✅")
            return True
    except Exception as e:
        logger.error(f"Send reply error: {e}")
        return False


def analyse_email_content(sender, subject, body, is_cc=False):
    try:
        if is_cc:
            prompt = f"""You have been CC'd on this email. Analyse for internal SCOPE Consulting purposes ONLY.
Do NOT write a reply. Internal analysis only.

From: {sender}
Subject: {subject}
Content: {body}

Provide:
1. Email type (MAR/RFI/BOQ/Variation/Claim/NCR/General)
2. Key points — 2-3 lines
3. Action required from SCOPE team
4. Risk level: High/Medium/Low
5. Recommended response deadline"""
        else:
            prompt = f"""You received this email directly from a SCOPE Consulting team member.
Write a complete professional reply.

From: {sender}
Subject: {subject}
Content: {body}

Write full professional email reply. Match the language exactly."""

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
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
        mail.select("INBOX")

        # Fetch all emails
        status, messages = mail.search(None, "ALL")
        if status != "OK":
            logger.error("Could not search emails")
            return

        email_ids = messages[0].split()
        # Process last 20 emails only
        recent_ids = email_ids[-20:] if len(email_ids) > 20 else email_ids
        logger.info(f"Found {len(recent_ids)} recent emails")

        new_count = 0
        for eid in reversed(recent_ids):
            try:
                msg_id_str = eid.decode()

                if msg_id_str in processed_ids:
                    continue

                status, msg_data = mail.fetch(eid, "(RFC822)")
                if status != "OK":
                    continue

                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                sender      = decode_str(msg.get("From", "")).lower()
                subject     = decode_str(msg.get("Subject", "No subject"))
                to_field    = decode_str(msg.get("To", "")).lower()
                cc_field    = decode_str(msg.get("CC", "")).lower()
                msg_id_hdr  = msg.get("Message-ID", "")

                # Skip emails sent by Alex himself
                if ZOHO_EMAIL.lower() in sender:
                    save_processed_id(msg_id_str)
                    continue

                is_direct   = ZOHO_EMAIL.lower() in to_field
                is_cc       = ZOHO_EMAIL.lower() in cc_field
                is_internal = any(team_email in sender for team_email in SCOPE_TEAM_EMAILS)

                if not is_direct and not is_cc:
                    save_processed_id(msg_id_str)
                    continue

                new_count += 1
                body = get_email_body(msg)

                if is_direct and is_internal:
                    logger.info(f"Direct internal: {sender} | {subject}")
                    analysis = analyse_email_content(sender, subject, body, is_cc=False)
                    if analysis:
                        reply_subject = f"Re: {subject}" if not subject.startswith("Re:") else subject
                        sent = send_reply(sender, reply_subject, analysis, msg_id_hdr)
                        status_val = "Closed" if sent else "Open"
                        save_to_memory(sender, subject, analysis[:300], "Replied by Alex", status_val)

                elif is_direct and not is_internal:
                    logger.info(f"Ignoring external direct: {sender}")

                elif is_cc:
                    logger.info(f"CC'd: {sender} | {subject}")
                    analysis = analyse_email_content(sender, subject, body, is_cc=True)
                    if analysis:
                        save_to_memory(sender, subject, analysis[:300], "Logged — action required", "Monitoring")

                save_processed_id(msg_id_str)

            except Exception as e:
                logger.error(f"Process email error: {e}")
                continue

        mail.logout()
        logger.info(f"Processed {new_count} new emails")

    except Exception as e:
        logger.error(f"IMAP error: {e}")


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
        report += f"{'='*45}\n\n"

        if pending:
            report += f"CAVAB GÖZLƏYƏN / MONİTORİNQ: {len(pending)}\n\n"
            for r in pending[-15:]:
                report += f"— {r.get('Subject', r.get('Topic','?'))}\n"
                report += f"  Göndərən: {r.get('Sender', r.get('Project','?'))} | {r.get('Date','')} | {r.get('Status','')}\n"
                report += f"  Tədbirlər: {r.get('Action','')}\n\n"
        else:
            report += "Gözləyən e-poçt yoxdur. ✅\n\n"

        report += f"{'='*45}\n"
        report += f"Alex Rivera\n"
        report += f"Construction Expert | SCOPE Consulting MMC\n"
        report += f"internal@scope-iq.io"

        for recipient in REPORT_RECIPIENTS:
            send_reply(recipient, f"SCOPE IQ — Günlük Hesabat — {today}", report)

        logger.info(f"Morning report sent ✅")

    except Exception as e:
        logger.error(f"Morning report error: {e}")


def main():
    logger.info("Alex Email Service starting...")
    logger.info(f"Monitoring: {ZOHO_EMAIL}")
    logger.info(f"IMAP: {IMAP_HOST}:{IMAP_PORT}")
    logger.info(f"Authorised team: {SCOPE_TEAM_EMAILS}")

    # Check emails every 5 minutes
    schedule.every(5).minutes.do(process_emails)

    # Morning report 9:00 AM Baku (UTC+4 = 05:00 UTC)
    schedule.every().day.at("05:00").do(send_morning_report)

    # Run immediately
    process_emails()

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
