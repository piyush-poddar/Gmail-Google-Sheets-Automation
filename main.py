import os
import base64
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Constants
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'
SHEET_NAME = 'Sheet1'

def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

def get_unread_emails(gmail_service, max_results=10):
    results = gmail_service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=max_results).execute()
    messages = results.get('messages', [])

    email_data = []

    for msg in messages:
        msg_content = gmail_service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
        payload = msg_content['payload']
        headers = payload.get("headers")

        write_email = True

        subject = sender = date = ''
        for header in headers:
            if header['name'] == 'Subject':
                subject = header['value']
                if ("Action required" not in subject) or any(keyword in subject for keyword in ["billing", "payment"]):
                    write_email = False
                    break
            if header['name'] == 'From':
                sender = header['value']
            if header['name'] == 'Date':
                date = header['value']

        if not write_email:
            continue

        parts = payload.get('parts')
        body = ''
        if parts:
            for part in parts:
                if part['mimeType'] == 'text/plain':
                    data = part['body']['data']
                    body = base64.urlsafe_b64decode(data).decode('utf-8')
                    break
        else:
            data = payload.get('body', {}).get('data')
            if data:
                body = base64.urlsafe_b64decode(data).decode('utf-8')

        email_data.append([date, sender, subject, body[:98]])

    return email_data

def create_new_sheet(creds, sheet_title="My Gmail Emails"):
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet = {
        'properties': {
            'title': sheet_title
        }
    }
    spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
    sheet_id = spreadsheet.get('spreadsheetId')
    print(f"New Sheet Created: https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
    return sheet_id

def write_to_sheet(email_data, creds):
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(create_new_sheet(creds, "Google Cloud Verification Emails"))
    worksheet = sh.worksheet(SHEET_NAME)

    if worksheet.row_count == 0 or worksheet.cell(1, 1).value != 'Date':
        worksheet.append_row(["Date", "From", "Subject", "Message Snippet"])

    for row in email_data:
        worksheet.append_row(row)

def main():
    creds = get_credentials()
    gmail_service = build('gmail', 'v1', credentials=creds)

    print("Fetching emails and its data...")
    email_data = get_unread_emails(gmail_service, max_results=100)

    if email_data:
        print(f"Writing {len(email_data)} emails to Google Sheet...")
        write_to_sheet(email_data, creds)
        print("Done! Emails saved to Google Sheet.")
    else:
        print("No unread emails found.")

if __name__ == '__main__':
    main()
