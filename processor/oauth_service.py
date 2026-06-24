import os
import msal
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from logger import audit_logger

# Google Configuration
GOOGLE_CLIENT_CONFIG = {
    "web": {
        "client_id": os.environ.get("GOOGLE_CLIENT_ID"),
        "project_id": os.environ.get("GOOGLE_PROJECT_ID"),
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_secret": os.environ.get("GOOGLE_CLIENT_SECRET"),
        "redirect_uris": [os.environ.get("GOOGLE_REDIRECT_URI", "http://localhost:8501/oauth2callback")]
    }
}

# Microsoft Configuration
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET")
MS_TENANT_ID = os.environ.get("MS_TENANT_ID", "common")
MS_REDIRECT_URI = os.environ.get("MS_REDIRECT_URI", "http://localhost:8501/oauth2callback")
MS_AUTHORITY = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
MS_SCOPES = ["https://graph.microsoft.com/Mail.Read"]

def get_google_auth_url():
    flow = Flow.from_client_config(
        GOOGLE_CLIENT_CONFIG,
        scopes=['https://www.googleapis.com/auth/gmail.readonly']
    )
    flow.redirect_uri = GOOGLE_CLIENT_CONFIG["web"]["redirect_uris"][0]
    auth_url, state = flow.authorization_url(access_type='offline', include_granted_scopes='true')
    return auth_url, state

def get_google_credentials(code, state):
    flow = Flow.from_client_config(
        GOOGLE_CLIENT_CONFIG,
        scopes=['https://www.googleapis.com/auth/gmail.readonly'],
        state=state
    )
    flow.redirect_uri = GOOGLE_CLIENT_CONFIG["web"]["redirect_uris"][0]
    flow.fetch_token(code=code)
    return flow.credentials

def get_ms_auth_url():
    client = msal.ConfidentialClientApplication(
        MS_CLIENT_ID, authority=MS_AUTHORITY,
        client_credential=MS_CLIENT_SECRET
    )
    auth_url = client.get_authorization_request_url(MS_SCOPES, redirect_uri=MS_REDIRECT_URI)
    return auth_url

def get_ms_token(code):
    client = msal.ConfidentialClientApplication(
        MS_CLIENT_ID, authority=MS_AUTHORITY,
        client_credential=MS_CLIENT_SECRET
    )
    result = client.acquire_token_by_authorization_code(code, scopes=MS_SCOPES, redirect_uri=MS_REDIRECT_URI)
    if "access_token" in result:
        return result["access_token"]
    else:
        audit_logger.log_action("MS_OAUTH_ERROR", details={"error": result.get("error"), "desc": result.get("error_description")})
        return None
