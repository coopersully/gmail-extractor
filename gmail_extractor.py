import imaplib
import os
import email
import sys
import json


class GmailExtractor:

    def ask_login(self):
        print("\nPlease enter your Gmail log-in details below.")
        self.usr = input("Email Address: ")
        self.pwd = input("Password: ")

    def attempt_login(self):
        self.mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        attempt: imaplib.IMAP4
        try:
            attempt = self.mail.login(self.usr, self.pwd)
        except Exception as e:
            print()
            print("Login failed! An error occurred.")

            if type(e) == imaplib.IMAP4.error:
                print("This account requires you to log-in with an app-specific password.")
                print("Learn more about this at https://support.google.com/accounts/answer/185833")
            else:
                print(str(e))

            return False

        if attempt:
            print("\nLogin successful!")
            print("Please choose a destination folder in the form of, \"/Users/username/dest/:\"")
            self.destFolder = input("Destination: ")

            # Add trailing slash if it doesn't exist
            if not self.destFolder.endswith("/"):
                self.destFolder += "/"
            return True

        print()
        print("Login failed!")
        return False

    def ask_extract(self):
        print("\nWe have found " + str(self.mailCount) + " emails in the mailbox " + self.mailbox + ".")
        return True if input(
            "Do you wish to continue extracting all the emails into " + self.destFolder + "? (y/N) ").lower().strip()[
                       :1] == "y" else False

    def ask_mailbox(self):
        self.mailbox = input("\nPlease type the name of the mailbox you want to extract, e.g. Inbox: ")
        bin_count = self.mail.select(self.mailbox)[1]
        self.mailCount = int(bin_count[0].decode("utf-8"))
        return True if self.mailCount > 0 else False

    def search_mailbox(self):
        mailbox_type, self.data = self.mail.search(None, "ALL")
        self.ids = self.data[0]
        self.idsList = self.ids.split()

    def parse_emails(self):
        jsonOutput = {}
        for anEmail in self.data[0].split():
            mailbox_type, self.data = self.mail.fetch(anEmail, '(UID RFC822)')
            raw = self.data[0][1]
            try:
                raw_str = raw.decode("utf-8")
            except UnicodeDecodeError:
                try:
                    raw_str = raw.decode("ISO-8859-1")  # ANSI support
                except UnicodeDecodeError:
                    try:
                        raw_str = raw.decode("ascii")  # ASCII ?
                    except UnicodeDecodeError:
                        print("Failed to decode email; please contact the developer.")
                        return

            msg = email.message_from_string(raw_str)

            jsonOutput['subject'] = msg['subject']
            jsonOutput['from'] = msg['from']
            jsonOutput['date'] = msg['date']

            raw = self.data[0][0]
            raw_str = raw.decode("utf-8")
            uid = raw_str.split()[2]
            # Body #
            if msg.is_multipart():
                for part in msg.walk():
                    partType = part.get_content_type()

                    # Get body
                    if partType == "text/plain" and "attachment" not in part:
                        jsonOutput['body'] = part.get_payload()

                    # Get attachments
                    if part.get('Content-Disposition') is None:
                        attachment_name = part.get_filename()
                        if bool(attachment_name):
                            attachment_path = str(self.destFolder) + str(uid) + str("/") + str(attachment_name)
                            os.makedirs(os.path.dirname(attachment_path), exist_ok=True)
                            with open(attachment_path, "wb") as f:
                                f.write(part.get_payload(decode=True))
            else:
                try:
                    jsonOutput['body'] = msg.get_payload(decode=True).decode("utf-8")
                except UnicodeDecodeError:
                    try:
                        jsonOutput['body'] = msg.get_payload(decode=True).decode("ISO-8859-1")  # ANSI support
                    except UnicodeDecodeError:
                        try:
                            jsonOutput['body'] = msg.get_payload(decode=True).decode("ascii")  # ASCII ?
                        except UnicodeDecodeError:
                            pass

            outputDump = json.dumps(jsonOutput)
            emailInfoFilePath = str(self.destFolder) + str(uid) + str("/") + str(uid) + str(".json")
            os.makedirs(os.path.dirname(emailInfoFilePath), exist_ok=True)
            with open(emailInfoFilePath, "w") as f:
                f.write(outputDump)

    def __init__(self):

        # Initialize variables
        self.usr = ""
        self.pwd = ""
        self.mail = object
        self.mailbox = ""
        self.mailCount = 0
        self.destFolder = ""
        self.data = []
        self.ids = []
        self.idsList = []

        # Request user's login details and attempt to login
        self.ask_login()
        if self.attempt_login():
            # Query which mailbox to fetch
            not self.ask_mailbox() and sys.exit()
        else:
            sys.exit() # If login failed

        # Confirm extraction with email count
        not self.ask_extract() and sys.exit()

        '''
        Get content of all emails and extract subject,
        body, and attachments to user's given directory.
        '''
        self.search_mailbox()
        self.parse_emails()


if __name__ == "__main__":
    run = GmailExtractor()
