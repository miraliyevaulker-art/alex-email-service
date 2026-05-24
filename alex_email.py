import os
import json
import time
import schedule
import requests
import logging
from datetime import datetime
from anthropic import Anthropic
import gspread
from google.oauth2.service_account import Credentials

# Configuration
ZOHO_CLIENT_ID      = os.environ.get("ZOHO_CLIENT_ID")
ZOHO_CLIENT_SECRET  = os.environ.get("ZOHO_CLIENT_SECRET")
ZOHO_AUTH_CODE      = os.environ.get("ZOHO_AUTH_CODE")
ZOHO_EMAIL          = os.environ.get("ZOHO_EMAIL", "internal@scope-iq.io")
ANTHROPIC_API_KEY   = os.environ.get("ANTHROPIC_API_KEY")
GOOGLE_SHEET_ID     = "1_4L63VqFN6etwLRWW2zT0WyJTHUT8LLTWwqiNPJJR4w"
GOOGLE_CREDS_FILE   = "primordial-mile-495807-k9-0217981265dd.json"
MODEL               = "claude-haiku-4-5-20251001"

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

# Token storage
token_file = "/tmp/zoho_tokens.json"

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

LANGUAGE RULES:
1. English email → respond in English ONLY
2. Azerbaijani email → respond in Azerbaijani ONLY
3. Mixed → dominant language
4. NEVER mix languages

EMAIL REPLY FORMAT:
- Professional, concise, structured
- Reference document numbers, clause numbers where relevant
- Clear action items
- Signature: Alex Rivera | Construction Expert | SCOPE Consulting MMC | internal@scope-iq.io

AZERBAIJANI QUALITY:
- MANDATORY: ə ı ö ü ğ ş ç İ Ə Ö Ü Ğ Ş Ç
- AzDTN/GOST standard terminology
- Formal register always

QA/QC: MAR → ✅ TƏSDİQLƏNDİ / ⚠️ ŞƏRHLƏ / ❌ RƏDD EDİLDİ
BOQ: Every position vs Baku market rates
FIDIC/NEC4: Contractually correct responses always"""


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


def save_email_log(sender, subject, summary, action, status="Open"):
    try:
        sheet = get_sheet("Email Log")
        if sheet:
            sheet.append_row([
                datetime.now().strftime("%d.%m.%Y %H:%M"),
                sender, subject, summary, action, status
            ])
    except Exception as e:
        logger.error(f"Email log error: {e}")


def get_tokens():
    """Get or refresh Zoho tokens"""
    # Try to load saved tokens
    if os.path.exists(token_file):
        try:
            with open(token_file) as f:
                tokens = json.load(f)
            # Check if access token still valid
            if tokens.get("expires_at", 0) > time.time() + 60:
                return tokens.get("access_token")
            # Refresh using refresh token
            if tokens.get("refresh_token"):
                return refresh_access_token(tokens["refresh_token"])
        except:
            pass

    # First time — exchange auth code for tokens
    return exchange_auth_code()


def exchange_auth_code():
    """Exchange auth code for access + refresh tokens"""
    try:
        response = requests.post(
            "https://accounts.zoho.com/oauth/v2/token",
            data={
                "code": ZOHO_AUTH_CODE,
                "client_id": ZOHO_CLIENT_ID,
                "client_secret": ZOHO_CLIENT_SECRET,
                "redirect_uri": "https://localhost",
                "grant_type": "authorization_code"
            }
        )
        data = response.json()
        if "access_token" in data:
            tokens = {
                "access_token": data["access_token"],
                "refresh_token": data.get("refresh_token"),
                "expires_at": time.time() + data.get("expires_in", 3600)
            }
            with open(token_file, "w") as f:
                json.dump(tokens, f)
            logger.info("Tokens obtained successfully")
            return data["access_token"]
        else:
            logger.error(f"Token exchange failed: {data}")
            return None
    except Exception as e:
        logger.error(f"Auth code exchange error: {e}")
        return None


def refresh_access_token(refresh_token):
    """Refresh expired access token"""
    try:
        response = requests.post(
            "https://accounts.zoho.com/oauth/v2/token",
            data={
                "refresh_token": refresh_token,
                "client_id": ZOHO_CLIENT_ID,
                "client_secret": ZOHO_CLIENT_SECRET,
                "grant_type": "refresh_token"
            }
        )
        data = response.json()
        if "access_token" in data:
            tokens = {
                "access_token": data["access_token"],
                "refresh_token": refresh_token,
                "expires_at": time.time() + data.get("expires_in", 3600)
            }
            with open(token_file, "w") as f:
                json.dump(tokens, f)
            return data["access_token"]
        return None
    except Exception as e:
        logger.error(f"Token refresh error: {e}")
        return None


def get_account_id(access_token):
    """Get Zoho Mail account ID"""
    try:
        response = requests.get(
            "https://mail.zoho.com/api/accounts",
            headers={"Authorization": f"Zoho-oauthtoken {access_token}"}
        )
        data = response.json()
        accounts = data.get("data", [])
        if accounts:
            return accounts[0].get("accountId")
        return None
    except Exception as e:
        logger.error(f"Account ID error: {e}")
        return None


def fetch_unread_emails(access_token, account_id):
    """Fetch unread emails"""
    try:
        response = requests.get(
            f"https://mail.zoho.com/api/accounts/{account_id}/messages/view",
            headers={"Authorization": f"Zoho-oauthtoken {access_token}"},
            params={"limit": 20, "start": 0, "status": "unread"}
        )
        data = response.json()
        return data.get("data", [])
    except Exception as e:
        logger.error(f"Fetch emails error: {e}")
        return []


def get_email_content(access_token, account_id, message_id):
    """Get full email content"""
    try:
        response = requests.get(
            f"https://mail.zoho.com/api/accounts/{account_id}/messages/{message_id}/content",
            headers={"Authorization": f"Zoho-oauthtoken {access_token}"}
        )
        data = response.json()
        return data.get("data", {}).get("content", "")
    except Exception as e:
        logger.error(f"Get email content error: {e}")
        return ""


def mark_as_read(access_token, account_id, message_id):
    """Mark email as read"""
    try:
        requests.put(
            f"https://mail.zoho.com/api/accounts/{account_id}/updatemessage",
            headers={
                "Authorization": f"Zoho-oauthtoken {access_token}",
                "Content-Type": "application/json"
            },
            json={
                "mode": "markAsRead",
                "messageId": [message_id]
            }
        )
    except Exception as e:
        logger.error(f"Mark read error: {e}")


def send_email(access_token, account_id, to_email, subject, body, reply_to_id=None):
    """Send email via Zoho"""
    try:
        payload = {
            "fromAddress": ZOHO_EMAIL,
            "toAddress": to_email,
            "subject": subject,
            "content": body,
            "mailFormat": "plaintext"
        }
        if reply_to_id:
            payload["inReplyTo"] = reply_to_id

        response = requests.post(
            f"https://mail.zoho.com/api/accounts/{account_id}/messages",
            headers={
                "Authorization": f"Zoho-oauthtoken {access_token}",
                "Content-Type": "application/json"
            },
            json=payload
        )
        result = response.json()
        if response.status_code == 200:
            logger.info(f"Email sent to {to_email}")
            return True
        else:
            logger.error(f"Send email failed: {result}")
            return False
    except Exception as e:
        logger.error(f"Send email error: {e}")
        return False


def analyse_email(sender, subject, body, is_cc=False):
    """Analyse email with Alex AI"""
    try:
        if is_cc:
            prompt = f"""You have been CC'd on this email. Analyse it for internal SCOPE Consulting purposes only.
Do NOT reply to sender. Just provide a brief internal analysis.

From: {sender}
Subject: {subject}
Content: {body[:3000]}

Provide:
1. Email type (MAR/RFI/BOQ/Variation/General)
2. Key points
3. Action required from SCOPE team
4. Risk level (High/Medium/Low)
5. Suggested deadline for response"""
        else:
            prompt = f"""You received this email directly from a SCOPE Consulting team member.
Provide a professional, structured reply.

From: {sender}
Subject: {subject}
Content: {body[:3000]}

Write a complete professional email reply. Match the language of the email."""

        response = anthropic_client.messages.create(
            model=MODEL,
            max_tokens=1500,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        logger.error(f"Email analysis error: {e}")
        return None


def process_emails():
    """Main email processing loop"""
    logger.info("Checking emails...")
    access_token = get_tokens()
    if not access_token:
        logger.error("No access token — skipping")
        return

    account_id = get_account_id(access_token)
    if not account_id:
        logger.error("No account ID — skipping")
        return

    emails = fetch_unread_emails(access_token, account_id)
    logger.info(f"Found {len(emails)} unread emails")

    for email in emails:
        try:
            sender = email.get("sender", "").lower()
            subject = email.get("subject", "No subject")
            message_id = email.get("messageId")
            to_addresses = email.get("toAddress", "").lower()
            cc_addresses = email.get("ccAddress", "").lower()

            # Determine if direct or CC
            is_direct = ZOHO_EMAIL.lower() in to_addresses
            is_cc = ZOHO_EMAIL.lower() in cc_addresses
            is_internal = any(team_email in sender for team_email in SCOPE_TEAM_EMAILS)

            if not is_direct and not is_cc:
                continue

            # Get full content
            body = get_email_content(access_token, account_id, message_id)

            if is_direct and is_internal:
                # Internal team email — analyse and reply
                logger.info(f"Direct from internal: {sender} — {subject}")
                analysis = analyse_email(sender, subject, body, is_cc=False)
                if analysis:
                    reply_subject = f"Re: {subject}" if not subject.startswith("Re:") else subject
                    send_email(access_token, account_id, sender, reply_subject, analysis, message_id)
                    save_email_log(sender, subject, analysis[:200], "Replied", "Closed")

            elif is_direct and not is_internal:
                # External direct email — ignore completely
                logger.info(f"Ignoring external direct email from: {sender}")
                mark_as_read(access_token, account_id, message_id)
                continue

            elif is_cc:
                # CC'd email — analyse silently for internal log
                logger.info(f"CC'd email from: {sender} — {subject}")
                analysis = analyse_email(sender, subject, body, is_cc=True)
                if analysis:
                    save_email_log(sender, subject, analysis[:200], "Logged from CC", "Monitoring")

            mark_as_read(access_token, account_id, message_id)

        except Exception as e:
            logger.error(f"Process email error: {e}")
            continue


def send_morning_report():
    """Send daily 9:00 AM report to SCOPE team"""
    logger.info("Sending morning report...")
    try:
        # Read email log
        sheet = get_sheet("Email Log")
        pending = []
        if sheet:
            records = sheet.get_all_records()
            for r in records:
                if r.get("Status") in ["Open", "Monitoring"]:
                    pending.append(r)

        today = datetime.now().strftime("%d.%m.%Y")
        day_name = datetime.now().strftime("%A")
        day_az = {
            "Monday": "Bazar ertəsi",
            "Tuesday": "Çərşənbə axşamı",
            "Wednesday": "Çərşənbə",
            "Thursday": "Cümə axşamı",
            "Friday": "Cümə",
            "Saturday": "Şənbə",
            "Sunday": "Bazar"
        }.get(day_name, day_name)

        report = f"SCOPE IQ — Günlük E-poçt Hesabatı\n"
        report += f"{day_az}, {today}\n"
        report += f"{'='*40}\n\n"

        if pending:
            report += f"CAVAB GÖZLƏYƏNLƏR: {len(pending)}\n\n"
            for r in pending[-10:]:
                report += f"— {r.get('Subject','?')} | {r.get('Sender','?')} | {r.get('Date','?')} | {r.get('Status','?')}\n"
        else:
            report += "Gözləyən e-poçt yoxdur. ✅\n"

        report += f"\n{'='*40}\n"
        report += "Alex Rivera | SCOPE Consulting MMC\ninternal@scope-iq.io"

        # Send to all team members
        access_token = get_tokens()
        if access_token:
            account_id = get_account_id(access_token)
            if account_id:
                for recipient in REPORT_RECIPIENTS:
                    send_email(
                        access_token,
                        account_id,
                        recipient,
                        f"SCOPE IQ — Günlük Hesabat — {today}",
                        report
                    )
                logger.info("Morning report sent to all team members")

    except Exception as e:
        logger.error(f"Morning report error: {e}")


def main():
    logger.info("Alex Email Service starting...")

    # Process emails every 5 minutes
    schedule.every(5).minutes.do(process_emails)

    # Morning report at 9:00 AM Baku time (UTC+4 = UTC 05:00)
    schedule.every().day.at("05:00").do(send_morning_report)

    # Run immediately on start
    process_emails()

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
