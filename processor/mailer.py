"""
Email automation module for SCD Dbase Sorter.
Provides functions to send validation requests and finalized data emails.
"""

import smtplib
import os
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from pathlib import Path
from logger import audit_logger
from encryption import decrypt_file_to_memory

# Configuration file path
HOSPITAL_EMAILS_PATH = "/home/team/shared/SCD_Dbase_Sorter/data/config/hospital_emails.csv"
HOSPITALS_DIR = "/home/team/shared/SCD_Dbase_Sorter/data/hospitals/"

# SMTP Configuration (dummy defaults, can be overridden)
SMTP_CONFIG = {
    "host": "smtp.gmail.com",
    "port": 587,
    "starttls": True,
    "username": "",  # Fill in your SMTP username
    "password": "",  # Fill in your SMTP password/app-specific password
    "from_email": "scd.database@example.com",
    "from_name": "SCD Database System"
}


def load_hospital_emails():
    """
    Loads the hospital email configuration from CSV.
    Returns a DataFrame with Hospital, Email, Validator_Email columns.
    """
    if not os.path.exists(HOSPITAL_EMAILS_PATH):
        raise FileNotFoundError(f"Hospital emails config not found at {HOSPITAL_EMAILS_PATH}")
    
    df = pd.read_csv(HOSPITAL_EMAILS_PATH)
    return df


def get_hospital_email(hospital_name):
    """Returns the email address for a given hospital."""
    df = load_hospital_emails()
    match = df[df['Hospital'].str.lower() == hospital_name.lower()]
    if match.empty:
        return None
    return match.iloc[0]['Email']


def get_validator_email(hospital_name):
    """Returns the validator email for a given hospital."""
    df = load_hospital_emails()
    match = df[df['Hospital'].str.lower() == hospital_name.lower()]
    if match.empty:
        return None
    return match.iloc[0]['Validator_Email']


def get_all_hospitals():
    """Returns list of all hospitals with their emails."""
    return load_hospital_emails()


def get_hospital_data_path(hospital_name):
    """
    Returns the file path for a hospital's Excel data file.
    Uses the same naming logic as sorter.py.
    """
    safe_filename = "".join([c for c in hospital_name if c.isalnum() or c in (' ', '_')]).strip().replace(' ', '_')
    return os.path.join(HOSPITALS_DIR, f"{safe_filename}.xlsx")


def load_hospital_data(hospital_name):
    """
    Loads the DataFrame for a given hospital from their Excel file.
    
    Args:
        hospital_name: Name of the hospital (e.g., "City General")
    
    Returns:
        DataFrame if file exists, empty DataFrame otherwise
    """
    file_path = get_hospital_data_path(hospital_name)
    if os.path.exists(file_path):
        return pd.read_excel(file_path)
    return pd.DataFrame()


def _create_smtp_connection():
    """Creates and returns an SMTP connection."""
    config = SMTP_CONFIG
    server = smtplib.SMTP(config["host"], config["port"])
    if config.get("starttls"):
        server.starttls()
    if config.get("username") and config.get("password"):
        server.login(config["username"], config["password"])
    return server


def _send_email(to_email, subject, body, attachments=None):
    """
    Core email sending function.
    
    Args:
        to_email: Recipient email address
        subject: Email subject line
        body: Email body text (plaintext or HTML)
        attachments: List of file paths to attach (optional)
    
    Returns:
        bool: True if sent successfully, False otherwise
    """
    config = SMTP_CONFIG
    
    msg = MIMEMultipart()
    msg['From'] = f"{config['from_name']} <{config['from_email']}>"
    msg['To'] = to_email
    msg['Subject'] = subject
    
    # Attach body as plain text
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach files if provided
    if attachments:
        for file_path in attachments:
            if os.path.exists(file_path):
                # Check if it's an encrypted .xlsx in hospitals dir
                filename = os.path.basename(file_path)
                if file_path.startswith(HOSPITALS_DIR) or filename == "Master_Database.xlsx":
                    # Try to decrypt
                    file_content = decrypt_file_to_memory(file_path)
                else:
                    with open(file_path, "rb") as f:
                        file_content = f.read()
                
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file_content)
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                msg.attach(part)
    
    try:
        server = _create_smtp_connection()
        server.send_message(msg)
        server.quit()
        print(f"Email sent successfully to {to_email}")
        audit_logger.log_action("SEND_EMAIL", details={"to": to_email, "subject": subject, "attachments": [os.path.basename(f) for f in attachments] if attachments else []})
        return True
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
        audit_logger.log_action("SEND_EMAIL_ERROR", details={"to": to_email, "error": str(e)})
        return False


def send_validation_request(hospital_name, hospital_data_df=None, attachments=None):
    """
    Sends a validation request email to the validator for a hospital.
    
    Args:
        hospital_name: Name of the hospital
        hospital_data_df: DataFrame containing the hospital's data (optional if file auto-attached)
        attachments: Optional list of attachment file paths. If None, the hospital's
                     Excel file from data/hospitals/ will be attached automatically.
    
    Returns:
        bool: True if sent successfully, False otherwise
    """
    validator_email = get_validator_email(hospital_name)
    if not validator_email:
        print(f"No validator email found for hospital: {hospital_name}")
        return False
    
    # Auto-attach hospital's Excel file if no attachments provided
    if attachments is None:
        file_path = get_hospital_data_path(hospital_name)
        if os.path.exists(file_path):
            attachments = [file_path]
    
    subject = f"Validation Request - {hospital_name} SCD Data"
    record_count = len(hospital_data_df) if hospital_data_df is not None else "N/A"
    body = f"""
Dear Validator,

A new dataset has been submitted for the following hospital and requires validation:

Hospital: {hospital_name}
Records: {record_count}

Please review the attached data and confirm its accuracy by replying to this email.

If you have any questions, please contact the database administrator.

Best regards,
SCD Database System
"""
    return _send_email(validator_email, subject, body, attachments=attachments)


def send_discovery_invitation(recipient_email, discovery_link):
    """
    Sends a discovery invitation email with the secure link.
    """
    subject = "Action Required: SCD Data Discovery Request"
    body = f"""
Dear Recipient,

You have been invited by the SCD Database Administrator to participate in a secure Sickle Cell Disease (SCD) Data Discovery process.

This process will help identify relevant records in your local files and external email accounts to ensure our database is comprehensive and up-to-date.

To begin, please click the secure link below:
{discovery_link}

Security Notice:
- This link is unique to you and will expire in 7 days.
- You will be required to verify your identity via an SMS code sent to your registered mobile number upon clicking the link.
- Data is processed in-memory and remains secure and private.

If you did not expect this request, please ignore this email or contact the administrator.

Best regards,
SCD Database System
"""
    return _send_email(recipient_email, subject, body)

def send_finalized_data(hospital_name, hospital_data_df=None, attachments=None):
    """
    Sends finalized/validated data to the hospital's email address.
    
    Args:
        hospital_name: Name of the hospital
        hospital_data_df: DataFrame containing the finalized hospital data (optional)
        attachments: Optional list of additional attachment file paths. If None, the
                     hospital's Excel file from data/hospitals/ will be attached automatically.
    
    Returns:
        bool: True if sent successfully, False otherwise
    """
    hospital_email = get_hospital_email(hospital_name)
    if not hospital_email:
        print(f"No hospital email found for: {hospital_name}")
        return False
    
    # Auto-attach hospital's Excel file if no attachments provided
    if attachments is None:
        file_path = get_hospital_data_path(hospital_name)
        if os.path.exists(file_path):
            attachments = [file_path]
    
    subject = f"Finalized SCD Data - {hospital_name}"
    record_count = len(hospital_data_df) if hospital_data_df is not None else "N/A"
    body = f"""
Dear {hospital_name} Team,

Please find attached the finalized SCD (Sickle Cell Disease) data for your review and records.

Hospital: {hospital_name}
Total Records: {record_count}
Status: Validated and Confirmed

If you have any questions or require further assistance, please do not hesitate to reach out.

Best regards,
SCD Database System
"""
    return _send_email(hospital_email, subject, body, attachments=attachments)


def send_bulk_validation_requests(hospital_data_dict):
    """
    Sends validation request emails for multiple hospitals.
    
    Args:
        hospital_data_dict: Dict mapping hospital_name -> DataFrame
    
    Returns:
        dict: Summary of send results {"hospital_name": bool}
    """
    results = {}
    for hospital_name, df in hospital_data_dict.items():
        results[hospital_name] = send_validation_request(hospital_name, df)
    return results


def send_bulk_finalized_data(hospital_data_dict):
    """
    Sends finalized data emails for multiple hospitals.
    
    Args:
        hospital_data_dict: Dict mapping hospital_name -> DataFrame
    
    Returns:
        dict: Summary of send results {"hospital_name": bool}
    """
    results = {}
    for hospital_name, df in hospital_data_dict.items():
        results[hospital_name] = send_finalized_data(hospital_name, df)
    return results


def configure_smtp(host=None, port=None, username=None, password=None, from_email=None, from_name=None):
    """
    Update SMTP configuration.
    
    Args:
        host: SMTP server host
        port: SMTP server port
        username: SMTP username
        password: SMTP password
        from_email: Sender email address
        from_name: Sender display name
    """
    global SMTP_CONFIG
    if host:
        SMTP_CONFIG["host"] = host
    if port:
        SMTP_CONFIG["port"] = port
    if username:
        SMTP_CONFIG["username"] = username
    if password:
        SMTP_CONFIG["password"] = password
    if from_email:
        SMTP_CONFIG["from_email"] = from_email
    if from_name:
        SMTP_CONFIG["from_name"] = from_name


# ============================================================
# Discovery Invitation Email
# ============================================================

def send_discovery_invitation(recipient_email, discovery_link):
    """
    Sends a discovery invitation email with a secure link.
    
    Args:
        recipient_email: Email address of the discovery recipient
        discovery_link: Secure link to the discovery landing page
    
    Returns:
        bool: True if sent successfully
    """
    subject = "SCD Database - External Data Discovery Request"
    body = f"""
Dear Recipient,

You have been invited to participate in an external data discovery scan for the SCD Database System.

Purpose: To help identify Sickle Cell Disease (SCD) records that may be missing from the central database by scanning your local files and email attachments.

What happens next:
1. Click the secure link below to verify your identity
2. Enter the one-time code sent to your mobile phone
3. Grant temporary access to your email account (optional)
4. Review and confirm any discovered records

Secure Discovery Link: {discovery_link}

This link will expire in 7 days. Your email access is temporary and will be revoked after the scan.

If you did not expect this invitation, please ignore this email.

Best regards,
SCD Database System
"""
    return _send_email(recipient_email, subject, body, attachments=None)


# ============================================================
# Helper: Export DataFrame to temp Excel for attachment
# ============================================================

def export_df_to_excel(df, output_path):
    """Exports a DataFrame to an Excel file."""
    df.to_excel(output_path, index=False)
    return output_path


if __name__ == "__main__":
    # Simple test
    print("Testing mailer module...")
    hospitals = get_all_hospitals()
    print(f"Loaded {len(hospitals)} hospitals")
    print(hospitals)