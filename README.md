# EmailPrinter

A python script that monitors given email address and checks for an email with attachments, which it prints.

My family's home printer was somewhat old and incapable of connecting to our house's wifi network. As such, it being in the basement made printing files difficult and tedious. This simple script was made to be run on a Rasberry Pi that was connected to our printer. It logged into and repeatedly checked an email address I created to which files could be sent. Then, the files were processed and sent to the printer to be printed.

## OAuth2 Setup Instructions

1. Install the required dependencies:

   ```
   pip install -r requirements.txt
   ```

2. Go to the [Google Cloud Console](https://console.cloud.google.com/)

   - Create a new project or select an existing one
   - [Enable the Gmail API for your project](https://console.cloud.google.com/marketplace/product/google/gmail.googleapis.com)
   - Go to "Credentials" and create an OAuth Client ID, following instructions to configure required project settings.
   - Choose "Desktop app" as the application type
   - Download the client configuration file and save it as `credentials.json` in the same directory as the script

3. Run the script for the first time:

   - A browser window will open asking you to authorize the application
   - Sign in with your Google account
   - Grant the requested permissions
   - The script will save the credentials in `token.pickle` for future use

4. The script will now use OAuth2 authentication instead of password-based authentication, which is more secure and follows Google's recommended practices.

Note: The `token.pickle` file contains sensitive information and should be kept secure. Do not share it or commit it to version control.
