import win32com.client
import pandas as pd
import sys

def send_bulk_emails(csv_file_path, subject, email_template, sender_email=None):
    try:
        # Initialize Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # List all available accounts for debugging
        print("Available Outlook accounts:")
        accounts = namespace.Accounts
        sender_account = None
        for account in accounts:
            smtp_address = account.SmtpAddress
            print(f" - {smtp_address}")
            if sender_email and smtp_address.lower() == sender_email.lower():
                sender_account = account
                print(f"Selected account: {smtp_address}")
        
        # If sender_email was specified but no matching account was found
        if sender_email and not sender_account:
            print(f"No account found with email: {sender_email}")
            return
        
        # Read CSV file
        df = pd.read_csv(csv_file_path)
        
        # Check if required columns exist
        required_columns = ['email', 'name', 'version', 'qid','title']
        if not all(col in df.columns for col in required_columns):
            print("CSV must contain 'email', 'name', 'date', and 'url' columns")
            return
        
        # Iterate through each row in the CSV
        for index, row in df.iterrows():
            try:
                # Create new email
                mail = outlook.CreateItem(0)  # 0 = MailItem
                
                # Set the sending account if specified
                if sender_account:
                    mail.SentOnBehalfOfName = sender_account
                    #mail.SendUsingAccount  = sender_account
                    print(f"Using account {sender_account.SmtpAddress} for {row['email']}")
                
                # Set email properties
                mail.To = row['email']
                
                
                # Replace placeholders in email template
                email_body = email_template.replace('<name>', row['name'])
                email_body = email_body.replace('<version>', str(row['version']))
                email_body = email_body.replace('<qid>', str(row['qid']))
                email_body = email_body.replace('<title>', row['title'])
                subject = subject.replace('<title>', row['title'])
                mail.Body = email_body
                mail.Subject = subject
                
                # Send the email
                mail.Send()
                print(f"Email sent to {row['email']}")
                
            except Exception as e:
                print(f"Failed to send email to {row['email']}: {str(e)}")
                
        print("Bulk email sending completed.")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    # CSV file path
    csv_file = "recipients.csv"
    
    # Email subject
    email_subject = "Your static or dynamic title goes here-- <title>"
    
    # Email template
    email_template = """Hi <name>,
We have detected <title> with QID <qid> and version <version>."""
    
    # Specify the sender email address (replace with your second mailbox email)
    sender_email = ""  # Replace with the exact SMTP address
    
    # Call the function
    send_bulk_emails(csv_file, email_subject, email_template, sender_email)

