import email
import time
import os, sys
import win32
import win32print
import win32api
import win32con
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
import subprocess
import tempfile
import wmi

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
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
]

def get_credentials():
    """Gets valid user credentials from storage.
    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        try:
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        except Exception as e:
            logging.error(f"Error loading token.pickle: {str(e)}")
            logging.info("Will attempt to create new credentials.")
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                logging.error(f"Error refreshing credentials: {str(e)}")
                logging.info("Token has expired or been revoked. Please reauthorize the application.")
                # Delete the invalid token file
                try:
                    os.remove('token.pickle')
                except:
                    pass
                # Start fresh with new credentials
                creds = None
        
        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)
            except Exception as e:
                logging.error(f"Error during OAuth2 flow: {str(e)}")
                logging.error("""
Please ensure:
1. You have a valid credentials.json file
2. You have enabled the Gmail API in your Google Cloud Console
3. You have added your email as a test user in the OAuth consent screen
4. You have the correct redirect URI set (http://localhost)
""")
                sys.exit(1)
    
    return creds

def Diff(li1, li2):
    """Return elements in li2 that are not in li1."""
    return [x for x in li2 if x not in li1]

def getAttachments(service, message_id):
    """Get attachments from a message and return them as in-memory file objects."""
    message = service.users().messages().get(userId='me', id=message_id, format='raw').execute()
    msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
    mime_msg = email.message_from_bytes(msg_str)
    
    attachments = []
    for part in mime_msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fName = part.get_filename()
        if fName:
            # Create an in-memory file object
            file_data = part.get_payload(decode=True)
            attachments.append({
                'name': fName,
                'data': file_data
            })
    
    return attachments if attachments else None

def checkSumatraPDF():
    """Check if SumatraPDF is installed and return the path if found."""
    possible_paths = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        os.path.join(os.environ['LOCALAPPDATA'], "SumatraPDF", "SumatraPDF.exe")
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    return None

def isPrinterOnline(printer_name):
    """Check if the printer is online by checking USB device status."""
    try:
        # Get printer handle
        h_printer = win32print.OpenPrinter(printer_name)
        try:
            # Get printer info
            printer_info = win32print.GetPrinter(h_printer, 2)
            port_name = printer_info['pPortName']
            
            # Check if it's a USB printer
            if port_name.startswith('USB'):
                # Try to get device status
                try:
                    # Use Windows Management Instrumentation to check device status
                    c = wmi.WMI()
                    
                    # First check Win32_Printer for more detailed printer status
                    for printer in c.Win32_Printer(DeviceID=printer_name):
                        logging.info(f"Found printer in Win32_Printer: {printer.DeviceID}")
                        
                        # Map printer status codes to human-readable messages
                        status_codes = {
                            0: "Ready",
                            1: "Paused",
                            2: "Error",
                            3: "Pending Deletion",
                            4: "Paper Jam",
                            5: "Paper Out",
                            6: "Manual Feed",
                            7: "Paper Problem",
                            8: "Offline",
                            9: "IO Active",
                            10: "Busy",
                            11: "Printing",
                            12: "Output Bin Full",
                            13: "Not Available",
                            14: "Waiting",
                            15: "Processing",
                            16: "Initializing",
                            17: "Warming Up",
                            18: "Toner Low",
                            19: "No Toner",
                            20: "Page Punt",
                            21: "User Intervention Required",
                            22: "Out of Memory",
                            23: "Door Open",
                            24: "Server Unknown",
                            25: "Power Save"
                        }
                        
                        status_msg = status_codes.get(printer.PrinterStatus, f"Unknown status ({printer.PrinterStatus})")
                        logging.info(f"Printer status: {status_msg}")
                        logging.info(f"Printer state: {'Offline' if printer.WorkOffline else 'Online'}")
                        
                        if printer.WorkOffline:
                            logging.warning(f"Printer {printer_name} is set to work offline")
                            return False
                        
                        # Check USB device status
                        for usb in c.Win32_USBControllerDevice():
                            try:
                                device = usb.Dependent
                                if port_name in str(device.DeviceID):
                                    logging.info(f"Found USB device for printer: {device.DeviceID}")
                                    if device.Status == "OK":
                                        logging.info(f"USB device is connected and powered on")
                                        return True
                                    else:
                                        logging.warning(f"USB device status: {device.Status}")
                                        return False
                            except:
                                continue
                        
                        # If we get here, we didn't find the USB device
                        logging.warning(f"USB device for printer {printer_name} not found")
                        return False
                    
                    # If we get here, we didn't find the printer in Win32_Printer
                    logging.warning(f"Printer {printer_name} not found in Win32_Printer")
                    return False
                    
                except Exception as e:
                    logging.warning(f"Failed to check printer status: {str(e)}")
                    return False
            else:
                # For non-USB printers, use the standard status check
                status = printer_info['Status']
                if status == 0:
                    logging.info(f"Printer {printer_name} reports ready")
                    return True
                else:
                    logging.warning(f"Printer {printer_name} reports status: {status}")
                    return False
                    
        finally:
            win32print.ClosePrinter(h_printer)
            
    except Exception as e:
        logging.error(f"Error checking printer status: {str(e)}")
        return False

def printFile(file_data, file_name, retry_delay=30):
    """Print a file from memory using SumatraPDF with indefinite retry mechanism."""
    retry_count = 0
    while True:
        try:
            # Get the default printer
            printer_name = win32print.GetDefaultPrinter()
            if not printer_name:
                logging.error("No default printer found")
                return False

            # Check if printer is online
            if not isPrinterOnline(printer_name):
                retry_count += 1
                logging.warning(f"Printer is offline. Retry attempt {retry_count} in {retry_delay} seconds...")
                time.sleep(retry_delay)
                continue

            # Check for SumatraPDF
            sumatra_path = checkSumatraPDF()
            if not sumatra_path:
                logging.error("""SumatraPDF not found. Please install it first:
1. Download SumatraPDF from https://www.sumatrapdfreader.org/download-free-pdf-viewer
2. Install it using one of these methods:
   - Standard installation: Run the installer and use default settings
   - Portable installation: Extract to %LOCALAPPDATA%\SumatraPDF
3. Restart this application after installation""")
                return False

            logging.info(f"Using SumatraPDF at: {sumatra_path}")

            # Create a temporary file for printing
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_name)[1]) as temp_file:
                temp_file.write(file_data)
                temp_path = temp_file.name

            try:
                # Build and execute the print command
                cmd = [sumatra_path, "-print-to", printer_name, "-silent", temp_path]
                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                stdout, stderr = process.communicate()
                
                if process.returncode == 0:
                    logging.info(f"Successfully sent {file_name} to printer {printer_name}")
                    return True
                else:
                    error_msg = stderr.decode()
                    logging.error(f"Print failed with error: {error_msg}")
                    retry_count += 1
                    logging.warning(f"Print failed. Retry attempt {retry_count} in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                    continue

            finally:
                # Clean up the temporary file
                try:
                    os.unlink(temp_path)
                except Exception as e:
                    logging.error(f"Failed to delete temporary file {temp_path}: {str(e)}")

        except Exception as e:
            logging.error(f"Unexpected error while printing {file_name}: {str(e)}")
            return False

def processMessage(service, msg_id):
    """Process a single message: get attachments, print them, and move the message to trash if successful."""
    # First check if the message is already in trash
    try:
        message = service.users().messages().get(userId='me', id=msg_id, format='metadata', metadataHeaders=['labels']).execute()
        labels = message.get('labelIds', [])
        if 'TRASH' in labels:
            logging.info(f"Message {msg_id} is already in trash, skipping")
            return
    except Exception as e:
        logging.error(f"Failed to check message labels: {str(e)}")
        return

    attachments = getAttachments(service, msg_id)
    if attachments:
        all_successful = True
        for attachment in attachments:
            if not printFile(attachment['data'], attachment['name']):
                all_successful = False
                break
        
        if all_successful:
            try:
                # Move the message to trash
                service.users().messages().trash(userId='me', id=msg_id).execute()
                logging.info(f"Successfully moved message {msg_id} to trash after printing")
                
                # Verify the message was moved to trash
                message = service.users().messages().get(userId='me', id=msg_id, format='metadata', metadataHeaders=['labels']).execute()
                labels = message.get('labelIds', [])
                if 'TRASH' in labels:
                    logging.info(f"Verified message {msg_id} is in trash")
                else:
                    logging.error(f"Message {msg_id} was not moved to trash successfully")
            except Exception as e:
                logging.error(f"Failed to move message {msg_id} to trash: {str(e)}")
        else:
            logging.warning(f"Not deleting message {msg_id} due to failed print attempts")

detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

# Get OAuth2 credentials
creds = get_credentials()

# Build the Gmail service
service = build('gmail', 'v1', credentials=creds)

# Get initial list of messages
results = service.users().messages().list(
    userId='me', 
    labelIds=['INBOX'],
    q='-label:trash'  # Exclude messages in trash
).execute()
messages = results.get('messages', [])
processed_messages = set()  # Track all processed messages

logging.info(f"Initial message count: {len(messages)}")

# Setup printer
CurrentPrinter = win32print.GetDefaultPrinter()
if not CurrentPrinter:
    logging.error("No default printer found. Please set a default printer in Windows settings.")
    sys.exit(1)

logging.info(f"Using default printer: {CurrentPrinter}")

while True:
    try:
        # Check for new messages
        results = service.users().messages().list(
            userId='me', 
            labelIds=['INBOX'],
            q='-label:trash'  # Exclude messages in trash
        ).execute()
        messages = results.get('messages', [])
        current_ids = [msg['id'] for msg in messages]
        
        # Find new messages that haven't been processed
        new_ids = [msg_id for msg_id in current_ids if msg_id not in processed_messages]
        
        if new_ids:
            logging.info(f"Found {len(new_ids)} new messages to process")
            for msg_id in new_ids:
                processMessage(service, msg_id)
                processed_messages.add(msg_id)
        else:
            logging.debug("No new messages found")
        
        time.sleep(10)  # Wait 10 seconds before checking again
        
    except Exception as e:
        logging.error(f"Error in main loop: {str(e)}")
        time.sleep(30)  # Wait longer on error before retrying

