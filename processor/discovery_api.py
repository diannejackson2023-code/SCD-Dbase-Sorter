import os
import io
import sys
import json
import datetime
import pandas as pd

# Handle both relative and absolute imports for compatibility
try:
    from .discovery_service import initiate_discovery, get_discovery_request, update_discovery_status, revoke_discovery_token
    from .otp_service import send_otp, verify_otp
    from .oauth_service import get_google_auth_url, get_ms_auth_url, get_google_credentials, get_ms_token
    from .hashing_service import get_master_patient_hashes, compare_hashes
    from .search_bot import SearchBot
    from .mailer import send_discovery_invitation
    from .mapping import load_and_map_data
    from .logger import audit_logger
except ImportError:
    from discovery_service import initiate_discovery, get_discovery_request, update_discovery_status, revoke_discovery_token
    from otp_service import send_otp, verify_otp
    from oauth_service import get_google_auth_url, get_ms_auth_url, get_google_credentials, get_ms_token
    from hashing_service import get_master_patient_hashes, compare_hashes
    from search_bot import SearchBot
    from mailer import send_discovery_invitation
    from mapping import load_and_map_data
    from logger import audit_logger

import threading
from concurrent.futures import ThreadPoolExecutor

def lead_initiate_request(recipient_email, recipient_phone):
    """
    Called by the Lead Dashboard to start a discovery process.
    1. Generates a secure token.
    2. Stores the request.
    3. Sends an invitation email with the discovery link.
    """
    token = initiate_discovery(recipient_email, recipient_phone)
    
    # Base URL should be from a config or environment variable for permanent hosting
    base_url = os.environ.get("BASE_URL", "http://localhost:3000")
    discovery_link = f"{base_url.rstrip('/')}/?token={token}"
    
    # Send invitation email
    success = send_discovery_invitation(recipient_email, discovery_link)
    
    if success:
        update_discovery_status(token, "INVITATION_SENT")
        audit_logger.log_action("DISCOVERY_INVITATION_SENT", details={"email": recipient_email, "token": token})
    else:
        update_discovery_status(token, "INVITATION_FAILED")
        audit_logger.log_action("DISCOVERY_INVITATION_FAILED", details={"email": recipient_email, "token": token})
        
    return token, discovery_link

def lead_bulk_initiate_request(recipients):
    """
    Initiates multiple discovery requests efficiently using a thread pool.
    recipients: List of dicts with 'email' and 'phone'
    Returns a list of dicts with email, token, link
    """
    results = []
    base_url = os.environ.get("BASE_URL", "http://localhost:3000")
    
    # Maximum 10 concurrent email sends to avoid overwhelming the SMTP server or hitting limits
    with ThreadPoolExecutor(max_workers=10) as executor:
        for rec in recipients:
            email = rec.get('email')
            phone = rec.get('phone')
            if not email:
                continue
                
            token = initiate_discovery(email, phone)
            discovery_link = f"{base_url.rstrip('/')}/?token={token}"
            
            # Queue the invitation email
            executor.submit(_send_invitation_worker, token, email, discovery_link)
            
            results.append({
                "email": email,
                "token": token,
                "link": discovery_link
            })
            
    return results

def _send_invitation_worker(token, email, link):
    """Worker for background email sending."""
    success = send_discovery_invitation(email, link)
    if success:
        update_discovery_status(token, "INVITATION_SENT")
        audit_logger.log_action("DISCOVERY_INVITATION_SENT_BULK", details={"email": email, "token": token})
    else:
        update_discovery_status(token, "INVITATION_FAILED")
        audit_logger.log_action("DISCOVERY_INVITATION_FAILED_BULK", details={"email": email, "token": token})

def get_staging_queue(token):
    """
    Retrieves the current staging queue status for a given token.
    Allows real-time UI polling of files coming from the Companion App.
    """
    staging_dir = "/home/team/shared/SCD_Dbase_Sorter/data/staging"
    queue_file = os.path.join(staging_dir, "queue.json")
    if os.path.exists(queue_file):
        try:
            with open(queue_file, 'r') as f:
                queue = json.load(f)
                return queue.get(token, [])
        except:
            return []
    return []

def recipient_get_context(token):
    """
    Validates a discovery token and returns the recipient context.
    Supports the sequential flow (Local Scan -> Identification -> OTP Verification).
    """
    request = get_discovery_request(token)
    if not request:
        return None
    
    email = request.get("recipient_email", "")
    phone = request.get("recipient_phone", "")
    status = request["status"]
    
    # Mask phone for privacy if it exists
    masked_phone = ""
    if phone:
        if len(phone) > 7:
            masked_phone = phone[:3] + "*" * (len(phone) - 7) + phone[-4:]
        else:
            masked_phone = "***" + phone[-2:] if len(phone) > 2 else "***"
    
    # Determine if identity capture is needed
    # If the Lead provided both, we just ask for confirmation.
    # If anything is missing, or if they just finished local scan, we prompt.
    needs_identity = status in ["INITIATED", "INVITATION_SENT", "LOCAL_SCAN_DONE"]
    
    # If they are already verified, they don't need identity capture anymore
    if status in ["OTP_VERIFIED", "SCANNING", "SCAN_COMPLETED"]:
        needs_identity = False

    return {
        "recipient_email": email,
        "recipient_phone_masked": masked_phone,
        "recipient_phone_raw": phone,
        "status": status,
        "needs_identity": needs_identity,
        "has_local_results": len(request.get("local_results", [])) > 0
    }

def recipient_update_identity(token, email, phone):
    """
    Updates the recipient's identity details and moves status if needed.
    """
    request = get_discovery_request(token)
    if not request:
        return False
        
    update_data = {
        "recipient_email": email,
        "recipient_phone": phone
    }
    
    # If they were in INITIATED/INVITATION_SENT and updated identity, 
    # we stay there until local scan or OTP trigger.
    return update_discovery_status(token, request["status"], update_data)

def recipient_trigger_otp(token):
    """
    Triggers an SMS OTP for the given discovery token.
    Recipient must have a phone number associated with the token.
    """
    request = get_discovery_request(token)
    if not request or not request.get("recipient_phone"):
        return False
    
    return send_otp(request["recipient_phone"])

def recipient_verify_phone(token, code):
    """
    Called by the Landing Page to verify the OTP.
    Transitions status to OTP_VERIFIED and re-analyzes any local results.
    """
    request = get_discovery_request(token)
    if not request:
        return False
    
    if verify_otp(request["recipient_phone"], code):
        update_discovery_status(token, "OTP_VERIFIED")
        # Automatically re-analyze results found during local scan now that identity is verified
        recipient_reanalyze_all_results(token)
        return True
    return False

def recipient_process_local_file(token, file_name, file_content, password=None):
    """
    Processes an uploaded file (Excel/Word) found during the local scan.
    Allows processing before OTP verification.
    Results are associated with the token but marked as 'local_results'.
    """
    request = get_discovery_request(token)
    # Allow processing even without token if it's a direct browser-based local scan before verification
    if token and not request:
        return None
    
    # Allow local scan if token is valid and active (or if we are in 'pre-token' phase)
    if request:
        active_statuses = ["INITIATED", "INVITATION_SENT", "LOCAL_SCAN_DONE", "OTP_VERIFIED", "SCANNING", "SCAN_COMPLETED"]
        if request["status"] not in active_statuses:
            return None
        
    ext = os.path.splitext(file_name)[1].lower()
    df = pd.DataFrame()
    
    file_like = io.BytesIO(file_content)
    
    try:
        if ext in ['.xlsx', '.xls']:
            df = load_and_map_data(file_like, password=password)
        elif ext == '.docx':
            try:
                from .word_processor import process_word_file
            except ImportError:
                from word_processor import process_word_file
            df = process_word_file(file_like)
    except Exception as e:
        # Check if it's a password error
        if "password" in str(e).lower() or "encrypted" in str(e).lower():
            return {"error": "PASSWORD_REQUIRED", "filename": file_name}
        raise e
    
    if df.empty:
        return None
        
    # Check if identity is already verified
    is_verified = request["status"] in ["OTP_VERIFIED", "SCANNING", "SCAN_COMPLETED"]
    
    results = {
        "filename": file_name,
        "total_records": len(df),
        "data": df.to_dict(orient='records'),
        "type": "local",
        "is_analyzed": False,
        "date_found": datetime.datetime.now().isoformat()
    }
    
    if is_verified:
        master_hashes = get_master_patient_hashes()
        bot = SearchBot(None, None)
        df_analyzed = bot.identify_missing_records(df, master_hashes)
        results["missing_records"] = int(df_analyzed['Is_Missing'].sum()) if 'Is_Missing' in df_analyzed.columns else 0
        results["data"] = df_analyzed.to_dict(orient='records')
        results["is_analyzed"] = True
        
        # Log finding with comparison result
        status_msg = "NEW_RECORDS" if results["missing_records"] > 0 else "ALL_PRESENT"
        audit_logger.log_action("FILE_DISCOVERED", details={
            "filename": file_name,
            "type": "local",
            "records": len(df),
            "missing": results["missing_records"],
            "status": status_msg
        })
    else:
        results["missing_records"] = 0 # Not analyzed yet
        audit_logger.log_action("FILE_DISCOVERED", details={
            "filename": file_name,
            "type": "local",
            "records": len(df),
            "status": "AWAITING_VERIFICATION"
        })
    
    # Update status but don't overwrite other scan results
    local_results = request.get("local_results", [])
    # Check if we already have results for this filename
    local_results = [r for r in local_results if r.get("filename") != file_name]
    local_results.append(results)
    
    new_status = request["status"]
    if new_status in ["INITIATED", "INVITATION_SENT"]:
        new_status = "LOCAL_SCAN_DONE"
        
    update_discovery_status(token, new_status, {"local_results": local_results})
    
    return results

def recipient_reanalyze_all_results(token):
    """
    Re-analyzes all local results against the Master DB.
    Called after OTP verification to reveal which records are missing.
    """
    request = get_discovery_request(token)
    if not request or request["status"] not in ["OTP_VERIFIED", "SCANNING", "SCAN_COMPLETED"]:
        return False
        
    local_results = request.get("local_results", [])
    if not local_results:
        return True
        
    master_hashes = get_master_patient_hashes()
    bot = SearchBot(None, None)
    
    updated_local_results = []
    for res in local_results:
        if res.get("is_analyzed"):
            updated_local_results.append(res)
            continue
            
        df = pd.DataFrame(res["data"])
        if not df.empty:
            df_analyzed = bot.identify_missing_records(df, master_hashes)
            res["missing_records"] = int(df_analyzed['Is_Missing'].sum()) if 'Is_Missing' in df_analyzed.columns else 0
            res["data"] = df_analyzed.to_dict(orient='records')
            res["is_analyzed"] = True
        
        updated_local_results.append(res)
        
    update_discovery_status(token, request["status"], {"local_results": updated_local_results})
    return True

def recipient_start_scan(token, provider, access_token, password_map=None):
    """
    Executes the OAuth-based email search bot scan.
    Requires OTP verification.
    password_map: filename -> password dictionary for memorization
    """
    if password_map is None:
        password_map = {}
        
    request = get_discovery_request(token)
    if not request or request["status"] != "OTP_VERIFIED":
        return None
        
    update_discovery_status(token, "SCANNING")
    
    bot = SearchBot(provider, access_token)
    found_attachments = bot.scan_emails()
    
    master_hashes = get_master_patient_hashes()
    
    results = []
    for att in found_attachments:
        filename = att['filename']
        password = password_map.get(filename)
        
        df = bot.process_attachment(att['data'], filename, password=password)
        
        if isinstance(df, str) and df == "PASSWORD_REQUIRED":
            results.append({
                "filename": filename,
                "subject": att['subject'],
                "error": "PASSWORD_REQUIRED",
                "type": "email",
                "date_found": datetime.datetime.now().isoformat()
            })
            continue
            
        if isinstance(df, str) and df.startswith("DAMAGED_PDF"):
            results.append({
                "filename": filename,
                "subject": att['subject'],
                "error": "DAMAGED_PDF",
                "details": df,
                "type": "email",
                "date_found": datetime.datetime.now().isoformat()
            })
            audit_logger.log_action("FILE_DISCOVERED", details={
                "filename": filename,
                "type": "email",
                "status": "DAMAGED_PDF"
            })
            continue

        if not df.empty:
            df_analyzed = bot.identify_missing_records(df, master_hashes)
            missing_count = int(df_analyzed['Is_Missing'].sum()) if 'Is_Missing' in df_analyzed.columns else 0
            
            res_item = {
                "filename": filename,
                "subject": att['subject'],
                "total_records": len(df_analyzed),
                "missing_records": missing_count,
                "data": df_analyzed.to_dict(orient='records'),
                "type": "email",
                "is_analyzed": True,
                "date_found": datetime.datetime.now().isoformat()
            }
            results.append(res_item)
            
            # Log finding
            status_msg = "NEW_RECORDS" if missing_count > 0 else "ALL_PRESENT"
            audit_logger.log_action("FILE_DISCOVERED", details={
                "filename": filename,
                "type": "email",
                "records": len(df_analyzed),
                "missing": missing_count,
                "status": status_msg
            })
            
    update_discovery_status(token, "SCAN_COMPLETED", {"email_results": results})
    return results

def get_final_discovered_df(token):
    """
    Aggregates 'local_results' and 'email_results' for a given token.
    Returns a consolidated DataFrame of MISSING records ready for ingestion.
    """
    request = get_discovery_request(token)
    if not request:
        return pd.DataFrame()
    
    all_dfs = []
    
    # Aggregate analyzed local results
    local_results = request.get("local_results", [])
    for res in local_results:
        if res.get("is_analyzed") and res.get("data"):
            df = pd.DataFrame(res["data"])
            if not df.empty and 'Is_Missing' in df.columns:
                # Filter only records missing from Master DB
                missing_df = df[df['Is_Missing'] == True].copy()
                all_dfs.append(missing_df)
                
    # Aggregate analyzed email results
    email_results = request.get("email_results", [])
    for res in email_results:
        if res.get("is_analyzed") and res.get("data"):
            df = pd.DataFrame(res["data"])
            if not df.empty and 'Is_Missing' in df.columns:
                # Filter only records missing from Master DB
                missing_df = df[df['Is_Missing'] == True].copy()
                all_dfs.append(missing_df)
                
    if not all_dfs:
        return pd.DataFrame()
        
    final_df = pd.concat(all_dfs, ignore_index=True)
    
    # De-duplicate by Patient_ID to avoid double-entry if same patient in multiple files
    if not final_df.empty and 'Patient_ID' in final_df.columns:
        final_df = final_df.drop_duplicates(subset=['Patient_ID'])
        
    return final_df
