# Standard Library Imports
from datetime import date, timedelta
import json
import os.path
import base64

# Third-Party Imports
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Email MIME Imports
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



# GLOBAL VARIABLES 

def load_config(filename):
    with open(filename, "r") as config_file:
        return json.load(config_file)

config = load_config("config.json")
my_firstname = config["my_firstname"]
my_lastname = config["my_lastname"]
my_email = config["my_email"]
to_email = config["recipient_email"]
to_name = config["to_name"]

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def load_credentials():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    return creds

def refresh_credentials(creds):
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return creds

def authenticate():
    flow = InstalledAppFlow.from_client_secrets_file(
        'client_secret_file.json', SCOPES)
    creds = flow.run_local_server(port=0)
    return creds

def save_credentials(creds):
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

def initialize_service():
    creds = load_credentials()
    if not creds or not creds.valid:
        creds = refresh_credentials(creds) if creds else authenticate()
        save_credentials(creds)
    return build('gmail', 'v1', credentials=creds)

def create_message(sender, to, subject, message_text, file_name):
    msg = MIMEMultipart()
    msg['to'] = to
    msg['from'] = sender
    msg['subject'] = subject
    text_part = MIMEText(message_text, 'plain')
    msg.attach(text_part)
    attach_file(msg, file_name)
    return {'raw': base64.urlsafe_b64encode(msg.as_bytes()).decode()}


def attach_file(msg, file_name):
    with open(file_name, "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="xlsx")
        attach.add_header("Content-Disposition", f"attachment; filename= {file_name}")
        msg.attach(attach)

def create_mail(service, sender, recipient):
    time_period = f"{(date.today() - timedelta(days=7)).strftime('%d-%m-%Y')} to {date.today().strftime('%d-%m-%Y')}"
    subject = f"{time_period} Timesheet"
    message_text = f"Hi {to_name},\n\nI've attached my timesheet from {time_period}.\n\nKind regards,\n{my_firstname}."
    file_name = f"{time_period} Timesheet - {my_firstname} {my_lastname}.xlsx"

    message = create_message(sender, recipient, subject, message_text, file_name)
    result = service.users().messages().send(userId='me', body=message).execute()
    print(f"Sent message to {recipient}, message id: {result['id']}")

if __name__ == "__main__":
    service = initialize_service()
    sender_email = my_email
    recipient_email = to_email

    with open("timesheetMaker.py", "r") as f:
        code = f.read()
        exec(code)

    create_mail(service, sender_email, recipient_email)
