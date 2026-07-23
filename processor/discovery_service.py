import os
import json
import secrets
import datetime
import sys

# Handle both relative and absolute imports
try:
    from .logger import audit_logger
except ImportError:
    from logger import audit_logger

try:
    from .config import DISCOVERY_REQUESTS_PATH as DISCOVERY_DATA_FILE
except ImportError:
    from config import DISCOVERY_REQUESTS_PATH as DISCOVERY_DATA_FILE

def _load_requests():
    if not os.path.exists(DISCOVERY_DATA_FILE):
        return {}
    with open(DISCOVERY_DATA_FILE, "r") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return {}

def _save_requests(requests):
    os.makedirs(os.path.dirname(DISCOVERY_DATA_FILE), exist_ok=True)
    with open(DISCOVERY_DATA_FILE, "w") as f:
        json.dump(requests, f, indent=4)

def initiate_discovery(recipient_email, recipient_phone):
    """
    Creates a new discovery request and returns a secure token.
    """
    token = secrets.token_urlsafe(32)
    requests = _load_requests()
    
    requests[token] = {
        "recipient_email": recipient_email,
        "recipient_phone": recipient_phone,
        "status": "INITIATED",
        "created_at": datetime.datetime.now().isoformat(),
        "expires_at": (datetime.datetime.now() + datetime.timedelta(days=7)).isoformat()
    }
    
    _save_requests(requests)
    audit_logger.log_action("INITIATE_DISCOVERY", details={"email": recipient_email, "token": token})
    return token

def get_discovery_request(token):
    """
    Retrieves discovery request details if the token is valid, not expired, and not revoked.
    """
    requests = _load_requests()
    if token not in requests:
        return None
    
    request = requests[token]
    
    # Check status
    if request.get("status") == "REVOKED":
        audit_logger.log_action("DISCOVERY_TOKEN_REVOKED_ACCESS", details={"token": token})
        return None

    # Check expiry
    expiry = datetime.datetime.fromisoformat(request["expires_at"])
    if datetime.datetime.now() > expiry:
        audit_logger.log_action("DISCOVERY_TOKEN_EXPIRED", details={"token": token})
        return None
        
    return request

def revoke_discovery_token(token):
    """
    Revokes a discovery token.
    """
    return update_discovery_status(token, "REVOKED")

def update_discovery_status(token, status, additional_data=None):
    """
    Updates the status and data of a discovery request.
    """
    requests = _load_requests()
    if token not in requests:
        return False
        
    requests[token]["status"] = status
    requests[token]["updated_at"] = datetime.datetime.now().isoformat()
    
    if additional_data:
        requests[token].update(additional_data)
        
    _save_requests(requests)
    audit_logger.log_action("UPDATE_DISCOVERY_STATUS", details={"token": token, "status": status})
    return True

def purge_expired_requests():
    """
    Removes expired tokens from the storage.
    """
    requests = _load_requests()
    now = datetime.datetime.now()
    valid_requests = {k: v for k, v in requests.items() if datetime.datetime.fromisoformat(v["expires_at"]) > now}
    
    if len(valid_requests) != len(requests):
        _save_requests(valid_requests)
        audit_logger.log_action("PURGE_DISCOVERY_REQUESTS", details={"count": len(requests) - len(valid_requests)})
