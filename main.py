import email
import imaplib
import time
import os, sys
import win32
import win32print
import win32api
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import base64

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

def getContent(id):
    result, data = imap.fetch(id, "(RFC822)")
    msg = email.message_from_bytes(data[0][1])
    return msg.get_payload(0).get_payload(decode=True)

def getAttachments(id):
    result, data = imap.fetch(id, "(RFC822)")
    msg = email.message_from_bytes(data[0][1])
    fileName = []
    for part in msg.walk():
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
    win32api.ShellExecute(0, "print", fileName, CurrentPrinter, ".", 0)

detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

#credentials 
#CHANGE THESE TO YOUR CREDENTIALS
username = "mailprintserver@gmail.com"
password = "ypv.jfa7qup5brw9DFZ"

# Get OAuth2 credentials
creds = get_credentials()
access_token = creds.token

# Convert the access token to XOAUTH2 format
auth_string = f'user={username}\1auth=Bearer {access_token}\1\1'
auth_string = base64.b64encode(auth_string.encode()).decode()

# Setup connection to email using OAuth2
imap = imaplib.IMAP4_SSL("imap.gmail.com", 993)
imap.authenticate('XOAUTH2', lambda x: auth_string)

#setup printer stuff
CurrentPrinter = win32print.GetDefaultPrinter()


imap.select("Inbox", readonly=True)
result, ids = imap.search(None, "ALL")
id_list = ids[0].split()

attachments = getAttachments(id_list[len(id_list) - 4])
for att in attachments:
    printFile(att.name)
while True:
    imap.select("Inbox", readonly=True)
    result, ids2 = imap.search(None, "ALL")
    process_ids = Diff(id_list, ids2[0].split())
    for d in process_ids:
        attachments = getAttachments(d)

        for attachment in attachments:
            printFile(attachment.name)
        id_list.append(d)

