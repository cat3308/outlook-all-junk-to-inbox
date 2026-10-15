import imaplib
import time
#ok ah here is the config
IMAP_SERVER = "://outlook.com"  # Use ://outlook.com for personal hotmail
EMAIL_USER = "your email"
EMAIL_PASS = "your password"      # You must generate an App Password in Microsoft account security

def force_to_inbox():
    try:
        # Connect to Outlook
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USER, EMAIL_PASS)
        
        # Select the Junk folder
        # Note: If your Outlook is in another language, change "Junk" to your local folder name
        status, messages = mail.select("Junk")
        
        if status == "OK":
            # Get all email IDs in Junk
            type, data = mail.search(None, "ALL")
            mail_ids = data[0].split()
            
            if mail_ids:
                print(f"Found {len(mail_ids)} emails in Junk. Moving to Inbox...")
                for num in mail_ids:
                    # Move to Inbox
                    mail.copy(num, "Inbox")
                    # Delete from Junk
                    mail.store(num, "+FLAGS", "\\Deleted")
                
                # Permanently remove deleted items from Junk
                mail.expunge()
                print("Successfully forced all mail to Inbox!")
        
        mail.close()
        mail.logout()
        
    except Exception as fucked_yea:
       print(f"Error: {fucked_yea}")

# Run the check every 60 seconds
while True:
    force_to_inbox()
    time.sleep(60)
