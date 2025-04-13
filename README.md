# Email Printer

This script automatically prints PDF attachments from your Gmail inbox using SumatraPDF.

## Features

- Monitors Gmail inbox for new messages
- Automatically prints PDF attachments
- Moves processed messages to trash
- Retries printing when printer is offline
- Runs in the background

## Prerequisites

1. Python 3.x
2. SumatraPDF
3. Gmail account with OAuth2 credentials

## Installation

1. Install Python dependencies:

   ```bash
   pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client pywin32 wmi
   ```

2. Install SumatraPDF:

   - Download from https://www.sumatrapdfreader.org/download-free-pdf-viewer
   - Install using one of these methods:
     - Standard installation: Run the installer and use default settings
     - Portable installation: Extract to %LOCALAPPDATA%\SumatraPDF

3. Set up Gmail OAuth2 credentials:

   - Go to the Google Cloud Console
   - Create a new project
   - Enable the Gmail API
   - Create OAuth 2.0 credentials
   - Download the credentials and save as `credentials.json` in the script directory

4. Run the script once to authorize:
   ```bash
   python main.py
   ```
   - Follow the OAuth2 flow in your browser
   - The script will create a `token.pickle` file to store your credentials

## Automatic Startup

To make the script run automatically when Windows starts:

1. Press `Win + R`, type `shell:startup` and press Enter to open the Startup folder

2. Create a shortcut to `start_email_printer.bat`:
   - Right-click the batch file
   - Select "Create shortcut"
   - Move the shortcut to the Startup folder

The script will run in the background without showing a console window.

## Usage

1. Start the script:

   ```bash
   python main.py
   ```

2. The script will:

   - Monitor your Gmail inbox for new messages
   - Print any PDF attachments it finds
   - Move processed messages to trash
   - Retry printing if the printer is offline

3. To stop the script:
   - Open Task Manager
   - Find the Python process running `main.py`
   - End the process

## Troubleshooting

- If you get a "No default printer found" error:

  - Set a default printer in Windows settings
  - Make sure the printer is connected and powered on

- If you get a "SumatraPDF not found" error:

  - Make sure SumatraPDF is installed
  - Check the installation paths in the script

- If you get authentication errors:
  - Delete the `token.pickle` file
  - Run the script again to reauthorize

## License

This project is licensed under the MIT License - see the LICENSE file for details.
