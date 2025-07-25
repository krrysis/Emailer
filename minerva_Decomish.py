import os
import win32com.client
import pandas as pd
from dotenv import load_dotenv
import logging
from datetime import datetime

# Load environment variables from .env file
load_dotenv()

# Setup logging
logging.basicConfig(
    filename="emailer.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

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
        required_columns = ['minerva_url','MinervaID','ApplicationManager','ITSecurityOfficer']

        if not all(col in df.columns for col in required_columns):
            print("CSV must contain the following columns:", required_columns)
            return

        # Group by ApplicationManager and ITSecurityOfficer
        grouped = df.groupby(['ApplicationManager', 'ITSecurityOfficer'])

        for (app_manager, it_officer), group in grouped:
            try:
                # Create new email
                mail = outlook.CreateItem(0)  # 0 = MailItem

                # Set the sending account if specified
                if sender_account:
                    mail.SentOnBehalfOfName = sender_email
                    print(f"Using account {sender_account.SmtpAddress} for {app_manager}")

                # Set email properties
                mail.To = f"{app_manager}; {it_officer}"
                mail.CC = mailCC  # CC to specified email addresses

                # Extract first name from email
                email_address = app_manager
                first_name = email_address.split('@')[0].split('.')[0].capitalize()

                # Build the HTML table rows for all grouped entries
                table_rows = ""
                for _, row in group.iterrows():
                    table_rows += f"""
                    <tr>
                        <td>{row['MinervaID']}</td>
                        <td>{row['minerva_url']}</td>
                        <td>{row['ApplicationManager']}</td>
                        <td>{row['ITSecurityOfficer']}</td>
                    </tr>
                    """

                # Build the full email body
                email_body = f"""<p>Hello {first_name},</p>
<p>We have identified that the following application(s) have been marked as decommissioned in Minerva. However, they are still present in Qualys.<br>
<br>Could you please confirm whether the applications listed below are indeed decommissioned? If so, we will proceed with removing them from the Qualys subscription accordingly.</p>
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
  <tr>
    <th>Minerva ID</th>
    <th>Minerva URL</th>
    <th>Application Manager</th>
    <th>IT Security Officer</th>
  </tr>
  {table_rows}
</table>
<p>Please reply to this email with your confirmation or any additional information.<br>
Thank you for your attention to this matter.</p>
<br>
<p>Best regards,</p>
<p>CyberSOC Protect Integration Team</p>"""

                # Replace <id> in subject with comma-separated MinervaIDs
                minerva_ids = ', '.join(str(mid) for mid in group['MinervaID'])
                mail.Subject = subject.replace('<id>', minerva_ids)
                mail.HTMLBody = email_body  # Use HTML body for table formatting

                # Log To, CC, Subject with time
                logging.info(f"To: {mail.To} | CC: {mail.CC} | Subject: {mail.Subject}")

                # Send the email
                print(f"Sending email to {app_manager}, {it_officer} with subject: {mail.Subject}")
                mail.Send()
                print("email sent successfully.")
            except Exception as e:
                print(f"Failed to send email to {app_manager}, {it_officer} with subject: {subject}: {str(e)}")
                logging.error(f"Failed to send email to {app_manager}, {it_officer} with subject: {subject}: {str(e)}")

        print("Bulk email sending completed.")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        logging.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    # CSV file path
    csv_file = "Latest_Decomissioned_Domains_List.csv"
    
    # Email subject
    email_subject = "Action Needed: Confirm Decommissioned Application Status | MineveraID: <id>"

    # Specify the sender email address (replace with your second mailbox email)
    sender_email = os.getenv("sender_email")
    mailCC = os.getenv("mailCC")

    # Call the function
    send_bulk_emails(csv_file, email_subject, "", sender_email)

