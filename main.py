import imaplib
import email
import os
import time
import logging
import win32api
import win32print
import getpass  # To get the current user's username

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

# Outlook IMAP settings
IMAP_SERVER = 'outlook.office365.com'
EMAIL = 'email@outlook.com'
PASSWORD = 'password'
IMAP_FOLDER = 'INBOX'

# Function to fetch and print PDF attachments
def fetch_and_save_pdf_attachments():
    try:
        # Connect to the IMAP server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select(IMAP_FOLDER)

        logging.info("Searching for emails with subjects 'FT' or 'NC'")

        # Search for emails with subjects containing 'FT' or 'NC'
        result, data = mail.search(None, '(OR SUBJECT "FT" SUBJECT "NC")')

        for num in data[0].split():
            result, data = mail.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)

            # Check for attachments
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                filename = part.get_filename()
                if filename and filename.lower().endswith('.pdf'):
                    # Get current user's desktop folder path
                    desktop_path = os.path.join('C:\\Users', getpass.getuser(), 'Desktop')

                    # Create folder for current month if it doesn't exist
                    current_month = time.strftime("%B")
                    folder = os.path.join(desktop_path, f"faturas_mes_{current_month}")
                    if not os.path.exists(folder):
                        os.makedirs(folder)
                        logging.info(f"A folder for the current month {current_month} has been created")

                    # Save the PDF attachment to desktop folder
                    filepath = os.path.join(folder, filename)
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    logging.info(f"Saved PDF attachment: {filepath}")

                    # Print the PDF file
                    print_pdf(filepath)
                    logging.info(f"Printed PDF file: {filepath}")

        mail.close()
        mail.logout()

    except Exception as e:
        logging.error(f"Error: {str(e)}")

# Function to print a PDF file using win32api
def print_pdf(filepath):
    try:
        # Get default printer
        printer_name = win32print.GetDefaultPrinter()

        # Print PDF file
        win32api.ShellExecute(0, "print", filepath, f'/d:"{printer_name}"', ".", 0)
    except Exception as e:
        logging.error(f"Error printing file: {str(e)}")

# Function to run the script periodically
def run_periodically(interval):
    while True:
        fetch_and_save_pdf_attachments()
        time.sleep(interval)  # Wait for the specified interval before running again

# Test the functionality directly
if __name__ == '__main__':
    run_periodically(300)  # Run every 5 minutes (adjust interval as needed)
