import os
import json
import time
import imaplib
import email
import schedule
import logging
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

anthropic_client   = Anthropic(api_key=ANTHROPIC_API_KEY)
processed_ids_file = "/tmp/processed_emails.json"

SYSTEM_PROMPT = """You are Alex Rivera, Construction Expert at SCOPE Consulting MMC.

You have 25 years of experience across commercial, residential, infrastructure, oil and gas, and high-end fit-out projects in Europe, the Middle East, and CIS countries including Azerbaijan.

Your areas of expertise include Quantity Surveying, Cost Management, Quality Assurance and Quality Control, Contract Administration under FIDIC and NEC4, technical document review across all disciplines, materials and specifications, handover and commissioning, claims and variations, and all major construction standards including Eurocodes, British Standards, ISO, GOST, AzDTN and SNiP.

LANGUAGE RULES:
- Default language is English. Always reply in English unless the incoming email is written entirely in Azerbaijani.
- If the email is in Azerbaijani, reply in Azerbaijani using correct formal register and proper special characters: ə ı ö ü ğ ş ç.
- Never mix languages in one reply.

EMAIL WRITING STYLE - CRITICAL:
- Write exactly like a senior construction professional writing a formal business email.
- Use plain professional prose only. No bullet points, no symbols, no emojis, no checkmarks, no arrows, no dashes used as decoration, no hashtags.
- Paragraphs separated by blank lines.
- Begin with a formal salutation such as Dear Alishir, or Dear Team, as appropriate.
- End with a formal closing such as Kind regards or Yours sincerely.
- Be direct, confident and senior in tone. Never start with Certainly or Great question.
- Keep responses concise and actionable.

SIGNATURE — always end every email with exactly this:

Alex Rivera
Construction Expert
SCOPE Consulting MMC
internal@scope-iq.io

QA/QC RESPONSES:
For MAR reviews state Approved, Approved with Comments, or Rejected in plain sentences with clear reasons referencing specification clauses.
For BOQ reviews assess each position against Baku market rates in plain professional language.
For contractual matters apply FIDIC or NEC4 as appropriate with formal contractual language.

BAKU MARKET RATES FOR REFERENCE:
Mechanical excavation 8 to 15 AZN per cubic metre, Manual excavation 60 to 90 AZN per cubic metre, Concrete C25/30 190 to 240 AZN per cubic metre, Reinforcement 1200 to 1500 AZN per tonne, Formwork 18 to 28 AZN per square metre, External brickwork 25 to 40 AZN per square metre, Plastering 16 to 24 AZN per square metre, Ceramic tiles 25 to 40 AZN per square metre, Gypsum partition 32 to 48 AZN per square metre, Paint 12 to 18 AZN per square metre, HVAC ductwork 45 to 75 AZN per square metre, Fan coil unit 350 to 600 AZN each, Lighting fixture 45 to 120 AZN each, Sprinkler system 45 to 75 AZN per square metre, Metal door 4500 to 6500 AZN each, Aluminium door 700 to 900 AZN each."""


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
            logger.info(f"Saved: {subject}")
        else:
            logger.error("Sheet not accessible")
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


def get_email_body(msg):
    body = ""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                try:
                    if part.get_content_type() == "text/plain":
                        payload = part.get_payload(decode=True)
                        if payload:
                            charset = part.get_content_charset() or "utf-8"
                            body += payload.decode(charset, errors="replace")
                except:
                    continue
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="replace")
    except Exception as e:
        logger.error(f"Body error: {e}")
    return body[:3000]


def send_email(to_email, subject, body):
    try:
        resend.api_key = RESEND_API_KEY
        params = {
            "from": f"Alex Rivera <{ZOHO_EMAIL}>",
            "to": [to_email],
            "subject": subject,
            "text": body
        }
        resend.Emails.send(params)
        logger.info(f"Email sent to {to_email}")
        return True
    except Exception as e:
        logger.error(f"Resend error: {e}")
        return False


def analyse_email(sender, subject, body, is_cc=False):
    try:
        if is_cc:
            prompt = f"""You have been copied on the following email. Analyse it for internal SCOPE Consulting purposes only. Do not write a reply to the sender.

From: {sender}
Subject: {subject}
Content: {body}

Provide a brief internal assessment covering the following: the type of email (such as MAR, RFI, BOQ, Variation, Claim, NCR, or General), the key points in two or three sentences, the action required from the SCOPE team, the risk level as High, Medium or Low, and a recommended deadline for response."""

        else:
            prompt = f"""You have received the following email directly from a SCOPE Consulting team member. Write a complete professional reply.

From: {sender}
Subject: {subject}
Content: {body}

Write a formal professional email reply. Default to English unless the email above is written entirely in Azerbaijani. Follow the email writing style in your instructions exactly."""

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
        logger.info("IMAP login successful")
        mail.select("INBOX")

        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            logger.info("Inbox empty")
            mail.logout()
            return

        all_ids = data[0].split()
        if not all_ids:
            logger.info("No emails")
            mail.logout()
            return

        recent = all_ids[-25:]
        logger.info(f"Checking {len(recent)} recent emails")
        new_count = 0

        for eid in reversed(recent):
            try:
                eid_str = eid.decode() if isinstance(eid, bytes) else str(eid)

                if eid_str in processed_ids:
                    continue

                typ, msg_data = mail.fetch(eid, "(RFC822)")
                if typ != "OK" or not msg_data or not msg_data[0]:
                    save_processed_id(eid_str)
                    continue

                raw = msg_data[0][1]
                if not raw:
                    save_processed_id(eid_str)
                    continue

                msg = email.message_from_bytes(raw)

                sender   = safe_decode(msg.get("From"),    "").lower()
                subject  = safe_decode(msg.get("Subject"), "No subject")
                to_field = safe_decode(msg.get("To"),      "").lower()
                cc_field = safe_decode(msg.get("CC"),      "").lower()

                if ZOHO_EMAIL.lower() in sender:
                    save_processed_id(eid_str)
                    continue

                is_direct   = ZOHO_EMAIL.lower() in to_field
                is_cc_email = ZOHO_EMAIL.lower() in cc_field
                is_internal = any(t in sender for t in SCOPE_TEAM_EMAILS)

                if not is_direct and not is_cc_email:
                    save_processed_id(eid_str)
                    continue

                new_count += 1
                body = get_email_body(msg)
                logger.info(f"Processing: {sender} | {subject}")

                if is_direct and is_internal:
                    analysis = analyse_email(sender, subject, body, is_cc=False)
                    if analysis:
                        reply_sub = f"Re: {subject}" if not subject.startswith("Re:") else subject
                        sent = send_email(sender, reply_sub, analysis)
                        save_to_memory(
                            sender, subject, analysis[:300],
                            "Replied by Alex",
                            "Closed" if sent else "Open"
                        )

                elif is_direct and not is_internal:
                    logger.info(f"Ignoring external: {sender}")

                elif is_cc_email:
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

        if pending:
            report  = f"Dear Team,\n\n"
            report += f"Good morning. Please find below a summary of outstanding emails and open action items as of {today}.\n\n"

            for i, r in enumerate(pending[-15:], 1):
                subj   = r.get("Subject") or r.get("Topic") or "No subject"
                sender = r.get("Sender") or r.get("Project") or "Unknown"
                date   = r.get("Date") or ""
                status = r.get("Status") or ""
                action = r.get("Action") or "Action required"
                report += f"{i}. Subject: {subj}\n"
                report += f"   From: {sender}\n"
                report += f"   Date logged: {date}\n"
                report += f"   Status: {status}\n"
                report += f"   Action required: {action}\n\n"

            report += "Please review the above items and take the necessary action at your earliest convenience.\n\n"

        else:
            report  = f"Dear Team,\n\n"
            report += f"Good morning. As of {today}, there are no outstanding emails or open action items requiring your attention.\n\n"
            report += "Should you have any queries or wish to submit documents for review, please send them directly to this address.\n\n"

        report += "Kind regards,\n\n"
        report += "Alex Rivera\n"
        report += "Construction Expert\n"
        report += "SCOPE Consulting MMC\n"
        report += "internal@scope-iq.io"

        for recipient in REPORT_RECIPIENTS:
            send_email(
                recipient,
                f"SCOPE IQ Daily Report — {today}",
                report
            )
        logger.info("Morning report sent")

    except Exception as e:
        logger.error(f"Morning report error: {e}")


def main():
    logger.info("Alex Email Service starting")
    logger.info(f"Monitoring: {ZOHO_EMAIL}")
    logger.info(f"Authorised team: {SCOPE_TEAM_EMAILS}")
    logger.info(f"Report recipients: {REPORT_RECIPIENTS}")

    schedule.every(5).minutes.do(process_emails)
    schedule.every().day.at("05:00").do(send_morning_report)

    process_emails()

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
