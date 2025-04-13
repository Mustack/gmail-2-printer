import email
import time
import os, sys
import win32
import win32print
import win32api
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_printer.log'),
        logging.StreamHandler()
    ]
)

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_credentials():
    """Gets valid user credentials from storage.
    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def Diff(li1, li2):
    li_dif = [i for i in li1 + li2 if i not in li1 or i not in li2]
    return li_dif

def getAttachments(service, message_id):
    message = service.users().messages().get(userId='me', id=message_id, format='raw').execute()
    msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
    mime_msg = email.message_from_bytes(msg_str)
    
    fileName = []
    for part in mime_msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fName = part.get_filename()
        if fName:
            filePath = os.path.join(detach_dir, 'attachments', fName)
            print(fName)
            file = open(fName, 'wb')
            file.write(part.get_payload(decode=True))
            file.close()
            fileName.append(file)
    if bool(fileName):
        return fileName
    else:
        return None

def printFile(fileName):
    try:
        # Verify the file exists
        if not os.path.exists(fileName):
            logging.error(f"File not found: {fileName}")
            return False

        # Get the default printer
        printer_name = win32print.GetDefaultPrinter()
        if not printer_name:
            logging.error("No default printer found")
            return False

        # Check if printer is available
        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            win32print.ClosePrinter(hPrinter)
        except Exception as e:
            logging.error(f"Printer error: {str(e)}")
            return False

        # Try to print the file
        try:
            win32api.ShellExecute(0, "print", fileName, None, ".", 0)
            logging.info(f"Successfully sent {fileName} to printer {printer_name}")
            return True
        except Exception as e:
            logging.error(f"Print error: {str(e)}")
            return False

    except Exception as e:
        logging.error(f"Unexpected error while printing {fileName}: {str(e)}")
        return False

detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

# Get OAuth2 credentials
creds = get_credentials()

# Build the Gmail service
service = build('gmail', 'v1', credentials=creds)

# Get initial list of messages
results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
messages = results.get('messages', [])
id_list = [msg['id'] for msg in messages]

# Print attachments from the last 4 messages
for msg_id in id_list[-4:]:
    attachments = getAttachments(service, msg_id)
    if attachments:
        for attachment in attachments:
            printFile(attachment.name)

# Setup printer
CurrentPrinter = win32print.GetDefaultPrinter()
if not CurrentPrinter:
    logging.error("No default printer found. Please set a default printer in Windows settings.")
    sys.exit(1)

logging.info(f"Using default printer: {CurrentPrinter}")

while True:
    # Check for new messages
    results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
    messages = results.get('messages', [])
    current_ids = [msg['id'] for msg in messages]
    
    # Find new messages
    new_ids = Diff(current_ids, id_list)
    
    # Process new messages
    for msg_id in new_ids:
        attachments = getAttachments(service, msg_id)
        if attachments:
            for attachment in attachments:
                printFile(attachment.name)
        id_list.append(msg_id)
    
    time.sleep(10)  # Wait 10 seconds before checking again

