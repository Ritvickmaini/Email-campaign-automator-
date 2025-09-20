import smtplib
import ssl
import imaplib
import time
import urllib.parse
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from concurrent.futures import ThreadPoolExecutor, as_completed

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# ------------------ CONFIG ------------------
SMTP_SERVER = "285235.vps-10.com"
SMTP_PORT = 465
IMAP_SERVER = "285235.vps-10.com"
IMAP_PORT = 993

SENDER_EMAIL = "mike@artificialinteligencesummit.com"
SENDER_PASSWORD = "hydcr3hC~~ks0x8-"

SHEET_NAME = "Email-Campaigns(IOM)"  # Google Sheet name
TAB_NAME = "campaign-2"              # Tab name inside sheet
RANGE = "A:B"                        # A = Name, B = Email (C & D are reserved for global tracker)

EVENT_URL = "https://www.eventbrite.co.uk/e/ai-summit-2025-unlock-the-future-of-business-tickets-1702245194199?aff=emailcampaigns"
SERVICE_ACCOUNT_FILE = "/etc/secrets/service_account.json"


# ------------------ GOOGLE API SETUP ------------------
def get_spreadsheet_id(sheet_name):
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    drive_service = build("drive", "v3", credentials=creds)
    results = drive_service.files().list(
        q=f"name='{sheet_name}' and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)",
    ).execute()
    files = results.get("files", [])
    if not files:
        raise Exception(f"Spreadsheet '{sheet_name}' not found.")
    return files[0]["id"]


def get_sheet_service():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def get_sheet_data():
    service = get_sheet_service()
    spreadsheet_id = get_spreadsheet_id(SHEET_NAME)
    range_name = f"{TAB_NAME}!{RANGE}"

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    rows = result.get("values", [])
    return rows[1:], spreadsheet_id  # Skip header row


def get_campaign_stage(spreadsheet_id):
    service = get_sheet_service()
    range_name = f"{TAB_NAME}!C2:D2"
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    values = result.get("values", [])
    if not values or not values[0]:
        return 0, None
    template_stage = int(values[0][0]) if values[0][0].isdigit() else 0
    last_sent_str = values[0][1] if len(values[0]) > 1 else None
    return template_stage, last_sent_str


def update_campaign_stage(spreadsheet_id, template_num, timestamp):
    service = get_sheet_service()
    range_name = f"{TAB_NAME}!C2:D2"
    body = {"values": [[str(template_num), timestamp]]}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="RAW",
        body=body
    ).execute()


# ------------------ TRACKING LINKS ------------------
def build_tracking_links(recipient_email, subject, event_url):
    encoded_event_url = urllib.parse.quote(event_url, safe='')
    email_for_tracking = recipient_email if recipient_email else "unknown@example.com"
    encoded_subject = urllib.parse.quote(subject or "No Subject", safe='')

    tracking_link = f"https://tracking-enfw.onrender.com/track/click?email={email_for_tracking}&url={encoded_event_url}&subject={encoded_subject}"
    tracking_pixel = f'<img src="https://tracking-enfw.onrender.com/track/open?email={email_for_tracking}&subject={encoded_subject}" width="1" height="1" style="display:block; margin:0 auto;" alt="." />'
    unsubscribe_link = f"https://unsubscribe-uofn.onrender.com/unsubscribe?email={email_for_tracking}"

    return tracking_link, tracking_pixel, unsubscribe_link


# ------------------ EMAIL SENDING ------------------
def send_email(recipient, subject, body):
    try:
        context = ssl.create_default_context()
        msg = MIMEMultipart("alternative")
        msg["From"] = SENDER_EMAIL
        msg["To"] = recipient
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))

        # Send via SMTP
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, recipient, msg.as_string())

        # Save to "Inbox.Sent" folder using IMAP
        imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        imap.login(SENDER_EMAIL, SENDER_PASSWORD)
        imap.append('"Inbox.Sent"', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap.logout()

        print(f"✅ Sent to {recipient}")
        return True
    except Exception as e:
        print(f"❌ Error sending to {recipient}: {e}")
        return False


# ------------------ EMAIL TEMPLATES ------------------
EMAIL_TEMPLATES = [
    {
        "subject": "{name},Unlock your Future with AI",
        "body": """Dear {name},<br><br>
        The future is here, and it’s powered by Artificial Intelligence.<br>
        On 9th October 2025, Villa Marina, Douglas, Isle of Man, will host the AI Summit, 
        part of the British Isles and Isle of Man Business Show — a full-day experience 
        designed for business leaders, innovators, and professionals across the British Isles.<br><br>
        Why Attend:<br>
        • Discover how AI can transform your business operations, marketing, and customer engagement.<br>
        • Gain insights from top AI experts and industry leaders.<br>
        • Network with professionals who are shaping the future.<br><br>
        Seats are limited, and this is your chance to be part of a game-changing experience.<br>
        👉 <a href="{tracking_link}" target="_blank">Reserve Your Spot Today & Get your tickets.</a><br><br>
        Warm regards,<br>Mike
        """
    },
    {
        "subject": "{name},How Will AI Transform You and Your Industry?",
        "body": """Dear {name},<br><br>
        Imagine knowing how AI could accelerate growth, streamline operations, and unlock opportunities in your industry — before your competitors do.<br><br>
        At the AI Summit, you’ll:<br>
        • Learn practical strategies to apply AI in your business immediately.<br>
        • See live demonstrations and case studies from finance, healthcare, retail, and tech.<br>
        • Connect with leaders and innovators driving the AI revolution.<br><br>
        This is your opportunity to step into the future, see AI in action, and gain a competitive edge.<br>
        👉 <a href="{tracking_link}" target="_blank">Secure Your AI Summit Seat Now</a><br><br>
        Don’t miss the chance to experience the power of AI firsthand
        Best regards,<br>Mike
        """
    },
    {
        "subject": "{name},Be in the Room Where the Future is Shaped",
        "body": """Dear {name},<br><br>
        The AI Summit is not just a conference — it’s where ideas become action and connections spark innovation.<br><br>
        When you attend, you’ll:<br>
        • Hear from global AI thought leaders.<br>
        • Discover how businesses are transforming with AI.<br>
        • Network with decision-makers and innovators.<br><br>
        This is your chance to learn, be inspired, and leave with actionable strategies that can shape the future of your business.<br>
        👉 <a href="{tracking_link}" target="_blank">Register Now and Join the AI Summit</a><br><br>
        Seats are filling fast — don’t be left out of this unique experience.
        Warm regards,<br>Mike
        """
    },
    {
        "subject": "{name},Don’t Miss Out – Limited AI Summit Seats Available",
        "body": """Dear {name},<br><br>
        Have you reserved your place at the AI Summit yet?<br>
        The event on 9th October 2025 at Villa Marina is fast approaching, and seats are limited.<br><br>
        This is your chance to:<br>
        • Learn from AI experts and industry leaders.<br>
        • Discover practical AI solutions.<br>
        • Network with innovators from across the British Isles.<br><br>
        This is your moment to stay ahead of the curve and gain insights that could transform your business strategy and career.<br>
        👉 <a href="{tracking_link}" target="_blank">Claim Your AI Summit Seat Now</a><br><br>
        Don’t wait — opportunities like this don’t come often.
        Best regards,<br>Mike
        """
    },
    {
        "subject": "{name},Last Chance: Join the AI Summit and Transform Your Business",
        "body": """Dear {name},<br><br>
        This is it — the AI Summit is almost here, and this is your final chance to reserve your spot.<br><br>
        Join leaders, innovators, and change-makers on 9th October 2025 at Villa Marina, Isle of Man.<br>
        Seats are limited and filling fast. Don’t miss the opportunity to:<br>
        • Gain cutting-edge insights from AI experts.<br>
        • Network with top industry leaders.<br>
        • Discover practical AI strategies.<br><br>
        👉 <a href="{tracking_link}" target="_blank">Reserve Your Seat Before It’s Too Late</a><br><br>
        The future is waiting — make sure you’re part of it.
        Warm regards,<br>Mike
        """
    }
]


# ------------------ MAIN WORKFLOW ------------------
def main():
    rows, spreadsheet_id = get_sheet_data()
    template_stage, last_sent_str = get_campaign_stage(spreadsheet_id)

    if template_stage >= len(EMAIL_TEMPLATES):
        print("✅ Campaign finished. All templates sent.")
        return

    # Check 24h condition
    if last_sent_str:
        last_sent = datetime.strptime(last_sent_str, "%Y-%m-%d %H:%M:%S")
        if datetime.now() - last_sent < timedelta(hours=24):
            print("⏳ 24h not passed. Skipping this run.")
            return

    template = EMAIL_TEMPLATES[template_stage]

    with ThreadPoolExecutor(max_workers=20) as executor:
        futures = {}
        for row in rows:
            if len(row) < 2 or "@" not in row[1]:
                continue
            name = row[0] if row[0] else "Friend"
            recipient = row[1]

            # Personalize subject with name
            subject = template["subject"].format(name=name)

            tracking_link, tracking_pixel, unsubscribe_link = build_tracking_links(
                recipient, subject, EVENT_URL
            )

            body = f"""
            <html><body style="font-family: Arial, sans-serif; line-height:1.6;">
            {template['body'].format(name=name, tracking_link=tracking_link)}
            <p style="font-size:12px; color:gray; margin-top:40px;">
              If you don’t want to receive emails like this, <a href="{unsubscribe_link}">unsubscribe here</a>.
            </p>
            {tracking_pixel}
            </body></html>
            """

            futures[executor.submit(send_email, recipient, subject, body)] = recipient

        for future in as_completed(futures):
            future.result()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    update_campaign_stage(spreadsheet_id, template_stage + 1, timestamp)

    print("🎯 Campaign stage updated and run completed.")

def wait_until_ready(last_sent_str):
    """Check if 24h passed since last timestamp, else wait."""
    if not last_sent_str:  # first run (no timestamp yet)
        return

    last_sent = datetime.strptime(last_sent_str, "%Y-%m-%d %H:%M:%S")
    now = datetime.now()
    next_allowed = last_sent + timedelta(hours=24)

    if now < next_allowed:
        sleep_time = (next_allowed - now).total_seconds()
        print(f"⏳ Last sent at {last_sent}. Waiting {sleep_time/3600:.2f} hours...")
        time.sleep(sleep_time)
    else:
        print("✅ 24h already passed, continuing immediately...")

if __name__ == "__main__":
    while True:
        rows, spreadsheet_id = get_sheet_data()
        template_stage, last_sent_str = get_campaign_stage(spreadsheet_id)

        if template_stage >= len(EMAIL_TEMPLATES):
            print("✅ All templates sent. Campaign finished. Exiting.")
            break

        # --- Wait until 24h passed ---
        wait_until_ready(last_sent_str)

        # --- Run this stage ---
        main()


