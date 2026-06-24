import os
from twilio.rest import Client
from twilio.base.exceptions import TwilioRestException
from logger import audit_logger

# Twilio Configuration (should be in environment variables)
TWILIO_ACCOUNT_SID = os.environ.get("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.environ.get("TWILIO_AUTH_TOKEN")
TWILIO_VERIFY_SERVICE_SID = os.environ.get("TWILIO_VERIFY_SERVICE_SID")

def _get_client():
    if not TWILIO_ACCOUNT_SID or not TWILIO_AUTH_TOKEN:
        raise ValueError("Twilio credentials not configured in environment variables.")
    return Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

def send_otp(phone_number):
    """
    Sends a 6-digit OTP to the specified phone number using Twilio Verify.
    """
    if not TWILIO_VERIFY_SERVICE_SID:
        raise ValueError("Twilio Verify Service SID not configured.")
        
    client = _get_client()
    try:
        verification = client.verify.v2.services(TWILIO_VERIFY_SERVICE_SID) \
            .verifications \
            .create(to=phone_number, channel='sms')
        
        audit_logger.log_action("SEND_OTP", details={"phone": phone_number, "status": verification.status})
        return True
    except TwilioRestException as e:
        audit_logger.log_action("SEND_OTP_ERROR", details={"phone": phone_number, "error": str(e)})
        print(f"Error sending OTP: {e}")
        return False

def verify_otp(phone_number, code):
    """
    Verifies the 6-digit code for the specified phone number.
    """
    if not TWILIO_VERIFY_SERVICE_SID:
        raise ValueError("Twilio Verify Service SID not configured.")
        
    client = _get_client()
    try:
        verification_check = client.verify.v2.services(TWILIO_VERIFY_SERVICE_SID) \
            .verification_checks \
            .create(to=phone_number, code=code)
        
        success = verification_check.status == 'approved'
        audit_logger.log_action("VERIFY_OTP", details={"phone": phone_number, "status": verification_check.status})
        return success
    except TwilioRestException as e:
        audit_logger.log_action("VERIFY_OTP_ERROR", details={"phone": phone_number, "error": str(e)})
        print(f"Error verifying OTP: {e}")
        return False
