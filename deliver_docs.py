import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_final_docs(to_email):
    # This script requires SMTP credentials. 
    # Since I don't have the user's gmail password, I will prepare the email 
    # and explain that they can trigger it from the dashboard once they enter 
    # their SMTP details in the sidebar.
    
    files = [
        "/home/team/shared/SCD_Dbase_Sorter/TECHNICAL_MANUAL.pdf",
        "/home/team/shared/SCD_Dbase_Sorter/USER_MANUAL.pdf",
        "/home/team/shared/SCD_Dbase_Sorter/CHAT_HISTORY.pdf",
        "/home/team/shared/SCD_Dbase_Sorter/DEPLOYMENT_SECURITY.pdf",
        "/home/team/shared/SCD_Dbase_Sorter/TECHNICAL_MANUAL.md",
        "/home/team/shared/SCD_Dbase_Sorter/USER_MANUAL.md",
        "/home/team/shared/SCD_Dbase_Sorter/CHAT_HISTORY.md"
    ]
    
    print(f"Preparation complete. The following files are ready for delivery to {to_email}:")
    for f in files:
        if os.path.exists(f):
            print(f" - {f} (Found)")
        else:
            print(f" - {f} (NOT FOUND)")

if __name__ == "__main__":
    send_final_docs("diannejackson2023@gmail.com")
