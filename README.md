# EmailPrinter

A python script that monitors a Gmail account and automatically prints any attachments from new emails.

My family's home printer was somewhat old and incapable of connecting to our house's wifi network. As such, it being in the basement made printing files difficult and tedious. This simple script was made to be run on a Rasberry Pi that was connected to our printer. It monitors a Gmail account and automatically prints any attachments from new emails.

## Features

- Monitors a Gmail account for new emails
- Automatically prints any attachments found in new emails
- Uses Gmail API for secure authentication
- Runs continuously, checking for new emails every 10 seconds

## Requirements

- Python 3.6 or higher
- A Gmail account
- Windows (for printer functionality)
- SumatraPDF (for reliable PDF printing)

## Setup Instructions

### Installation

1. Clone this repository
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### OAuth2 Authentication Setup

To get the `credentials.json` file for the Email Printer application:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Gmail API:
   - Go to "APIs & Services" > "Library"
   - Search for "Gmail API"
   - Click "Enable"
4. Create OAuth 2.0 credentials:
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - Choose "Desktop app" as the application type
   - Give it a name (e.g., "Email Printer")
   - Click "Create"
5. Download the credentials:
   - Click the download icon (⬇️) next to your newly created OAuth client ID
   - Save the downloaded file as `credentials.json` in the same directory as your script

### Adding Test Users

Since this application is in testing mode, you need to add any email addresses that will be using the application to the test users list:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Navigate to "APIs & Services" > "OAuth consent screen"
3. Click on "Edit App" (pencil icon)
4. Under "Test users", click "ADD USERS"
5. Add the email addresses of all users who will be testing the application
   - For example: `mailprintserver@gmail.com`
6. Click "SAVE"

Note: Each test user will need to:

1. Use the email address you added to the test users list
2. Go through the OAuth consent flow when running the application for the first time
3. Accept the permissions requested by the application

The downloaded file will contain your actual client ID and client secret. The structure will look like this:

```json
{
  "installed": {
    "client_id": "your-actual-client-id.apps.googleusercontent.com",
    "project_id": "your-project-id",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_secret": "your-actual-client-secret",
    "redirect_uris": ["http://localhost"]
  }
}
```

### Security Notes

1. DO NOT share your `credentials.json` file
2. DO NOT commit it to version control
3. Keep your client ID and client secret secure
4. If you accidentally expose these credentials, you can revoke them in the Google Cloud Console and create new ones

### Troubleshooting

If you encounter any issues:

1. Make sure you've enabled the Gmail API
2. Verify that your OAuth consent screen is properly configured
3. Check that you've selected the correct application type (Desktop app)
4. Ensure the redirect URI is set to `http://localhost`
5. Verify that the email address is added to the test users list
6. If you get "This app isn't verified" warning:
   - Click "Advanced"
   - Click "Go to [Your App Name] (unsafe)"
   - Click "Allow" to grant the requested permissions

### Installing SumatraPDF

This application requires SumatraPDF for reliable PDF printing. You can install it in one of two ways:

1. **Standard Installation**:

   - Download SumatraPDF from [https://www.sumatrapdfreader.org/download-free-pdf-viewer](https://www.sumatrapdfreader.org/download-free-pdf-viewer)
   - Run the installer and use the default settings
   - The application will automatically detect the installation

2. **Portable Installation**:
   - Download the portable version of SumatraPDF
   - Extract it to `%LOCALAPPDATA%\SumatraPDF`
   - The application will automatically detect the portable installation

After installation, restart the application for it to detect SumatraPDF.

## Usage

1. Run the script:
   ```bash
   python main.py
   ```
2. The first time you run it, a browser window will open asking you to authorize the application
3. After authorization, the script will start monitoring your Gmail account for new emails with attachments
4. Any attachments found will be automatically printed to your default printer

## Limitations

- Currently only works with Gmail accounts
- Requires a Google Cloud Project and Gmail API setup
- Only works on Windows due to printer functionality
- Only prints attachments, not email content
