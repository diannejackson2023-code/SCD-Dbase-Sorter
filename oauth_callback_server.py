"""
OAuth Callback Server for External SCD Data Discovery.
Handles OAuth redirects from Google and Microsoft.
Runs on port 3000 as the public website.

This server is part of the volatile processing architecture:
- Access tokens are held in memory only
- Scan results are not persisted to disk
- All data is deleted after the session ends
"""

import os
import sys
import io
import json
import threading
import time
from flask import Flask, request, redirect, render_template_string, session, jsonify
from datetime import datetime, timedelta

# Add processor directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "processor"))

from processor.oauth_service import get_google_auth_url, get_google_credentials, get_ms_auth_url, get_ms_token
from processor.search_bot import SearchBot
from processor.hashing_service import get_master_patient_hashes
from processor.logger import audit_logger

# ============================================================================
# Configuration
# ============================================================================

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", os.urandom(32).hex())

# In-memory token storage (volatile - not persisted)
# This is cleared when the server restarts
TOKEN_STORE = {}
SCAN_RESULTS = {}
SERVER_START_TIME = datetime.now()

# HTML Templates
INDEX_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>SCD External Discovery</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; }
        .status { padding: 15px; border-radius: 5px; margin: 15px 0; }
        .status.pending { background: #fff3cd; color: #856404; }
        .status.success { background: #d4edda; color: #155724; }
        .status.error { background: #f8d7da; color: #721c24; }
        .btn { display: inline-block; padding: 12px 24px; background: #007bff; color: white; text-decoration: none; border-radius: 5px; margin: 10px 5px 10px 0; }
        .btn:hover { background: #0056b3; }
        .btn-outline { background: transparent; border: 2px solid #007bff; color: #007bff; }
        .btn-outline:hover { background: #007bff; color: white; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #f8f9fa; }
        .missing { color: #dc3545; font-weight: bold; }
        .found { color: #28a745; }
        .nav { margin-bottom: 30px; }
        .nav a { margin-right: 20px; color: #007bff; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🔍 SCD External Data Discovery</h1>
        
        <div class="nav">
            <a href="/">Home</a>
            <a href="/status">Status</a>
        </div>
        
        <hr>
        
        {% if error %}
        <div class="status error">
            <strong>Error:</strong> {{ error }}
        </div>
        {% endif %}
        
        {% if message %}
        <div class="status success">
            {{ message }}
        </div>
        {% endif %}
        
        {% block content %}{% endblock %}
        
        <hr>
        <small style="color: #666;">
            🔒 This server uses volatile (in-memory) processing only.<br>
            No data is stored on disk. Session data is deleted when you close this page.
        </small>
    </div>
</body>
</html>
"""

HOME_TEMPLATE = """
{% extends "base" %}
{% block content %}
<h2>External Data Discovery</h2>
<p>Find missing SCD records by scanning your email attachments.</p>

<div class="status pending">
    <strong>Step 1:</strong> Choose your email provider
</div>

<div style="margin: 30px 0;">
    <a href="/auth/google" class="btn">🔵 Scan Gmail</a>
    <a href="/auth/microsoft" class="btn" style="background: #0078d4;">📧 Scan Outlook</a>
</div>

<h3>How it works:</h3>
<ol style="line-height: 1.8;">
    <li>Click one of the provider buttons above to authorize read-only access</li>
    <li>We'll search your emails for SCD-related attachments (Excel/Word files)</li>
    <li>Found records are compared against our secure database (using SHA-256 hashes)</li>
    <li>Records <strong>not already in the database</strong> are shown for your review</li>
    <li>You confirm which records to submit</li>
</ol>

<div class="status pending">
    <strong>⚠️ Privacy Note:</strong> We use SHA-256 hashing to check records without seeing your actual data.
    Only new records you choose to submit are sent to the database.
</div>
{% endblock %}
"""

STATUS_TEMPLATE = """
{% extends "base" %}
{% block content %}
<h2>Scan Status</h2>

{% if not scan_results %}
<div class="status pending">
    No scans have been performed yet. <a href="/">Start a new scan</a>.
</div>
{% else %}
    {% for provider, result in scan_results.items() %}
    <div class="status success">
        <strong>{{ provider|upper }} Scan Complete</strong><br>
        Found {{ result.total_files }} files with {{ result.total_records }} total records.
        {{ result.missing_records }} records are NEW to the database.
    </div>
    
    {% if result.files %}
    <h3>Files Scanned:</h3>
    <table>
        <tr>
            <th>Filename</th>
            <th>Subject</th>
            <th>Total Records</th>
            <th>New Records</th>
        </tr>
        {% for file in result.files %}
        <tr>
            <td>{{ file.filename }}</td>
            <td>{{ file.subject }}</td>
            <td>{{ file.total_records }}</td>
            <td class="missing">{{ file.missing_records }}</td>
        </tr>
        {% endfor %}
    </table>
    {% endif %}
    {% endfor %}
{% endif %}
{% endblock %}
"""

OAUTH_START_TEMPLATE = """
{% extends "base" %}
{% block content %}
<h2>Starting OAuth...</h2>
<p>Redirecting to {{ provider }} authorization page...</p>
<p>If you are not redirected automatically, <a href="{{ auth_url }}">click here</a>.</p>
<script>window.location.href = "{{ auth_url }}";</script>
{% endblock %}
"""

def get_token_store_key(provider, session_id):
    return f"{provider}_{session_id}"

# ============================================================================
# Routes
# ============================================================================

@app.route('/')
def index():
    return render_template_string(INDEX_TEMPLATE + HOME_TEMPLATE)

@app.route('/status')
def status():
    return render_template_string(INDEX_TEMPLATE + STATUS_TEMPLATE, scan_results=SCAN_RESULTS)

@app.route('/auth/google')
def auth_google():
    """Start Google OAuth flow."""
    try:
        auth_url, state = get_google_auth_url()
        session['oauth_state'] = state
        session['provider'] = 'google'
        return render_template_string(INDEX_TEMPLATE + OAUTH_START_TEMPLATE, 
                                    auth_url=auth_url, provider='Google')
    except Exception as e:
        audit_logger.log_action("OAUTH_GOOGLE_ERROR", details={"error": str(e)})
        return render_template_string(INDEX_TEMPLATE, error=f"Failed to start Google OAuth: {str(e)}")

@app.route('/auth/microsoft')
def auth_microsoft():
    """Start Microsoft OAuth flow."""
    try:
        auth_url = get_ms_auth_url()
        session['provider'] = 'microsoft'
        return render_template_string(INDEX_TEMPLATE + OAUTH_START_TEMPLATE,
                                    auth_url=auth_url, provider='Microsoft')
    except Exception as e:
        audit_logger.log_action("OAUTH_MS_ERROR", details={"error": str(e)})
        return render_template_string(INDEX_TEMPLATE, error=f"Failed to start Microsoft OAuth: {str(e)}")

@app.route('/oauth2callback')
def oauth2callback():
    """Handle Google OAuth callback."""
    error = request.args.get('error')
    if error:
        return render_template_string(INDEX_TEMPLATE, error=f"Google OAuth error: {error}")
    
    code = request.args.get('code')
    state = request.args.get('state')
    
    if not code:
        return render_template_string(INDEX_TEMPLATE, error="No authorization code received")
    
    try:
        # Exchange code for credentials
        creds = get_google_credentials(code, state)
        access_token = creds.token if hasattr(creds, 'token') else creds.get('access_token')
        
        session_id = session.get('session_id', os.urandom(16).hex())
        session['session_id'] = session_id
        
        # Store token temporarily in memory (volatile)
        key = get_token_store_key('google', session_id)
        TOKEN_STORE[key] = {
            'provider': 'google',
            'access_token': access_token,
            'expires_at': datetime.now() + timedelta(hours=1)
        }
        
        # Start scan in background
        thread = threading.Thread(target=run_email_scan, args=('google', session_id))
        thread.daemon = True
        thread.start()
        
        audit_logger.log_action("OAUTH_GOOGLE_SUCCESS", details={"session": session_id})
        
        return render_template_string(INDEX_TEMPLATE, 
                                    message="✅ Google authorization successful! Scanning emails...")
    except Exception as e:
        audit_logger.log_action("OAUTH_GOOGLE_ERROR", details={"error": str(e)})
        return render_template_string(INDEX_TEMPLATE, error=f"Failed to complete Google OAuth: {str(e)}")

@app.route('/ms_oauth2callback')
def ms_oauth2callback():
    """Handle Microsoft OAuth callback."""
    error = request.args.get('error')
    if error:
        return render_template_string(INDEX_TEMPLATE, error=f"Microsoft OAuth error: {error}")
    
    code = request.args.get('code')
    
    if not code:
        return render_template_string(INDEX_TEMPLATE, error="No authorization code received")
    
    try:
        access_token = get_ms_token(code)
        
        if not access_token:
            return render_template_string(INDEX_TEMPLATE, error="Failed to get Microsoft access token")
        
        session_id = session.get('session_id', os.urandom(16).hex())
        session['session_id'] = session_id
        
        # Store token temporarily in memory (volatile)
        key = get_token_store_key('microsoft', session_id)
        TOKEN_STORE[key] = {
            'provider': 'microsoft',
            'access_token': access_token,
            'expires_at': datetime.now() + timedelta(hours=1)
        }
        
        # Start scan in background
        thread = threading.Thread(target=run_email_scan, args=('microsoft', session_id))
        thread.daemon = True
        thread.start()
        
        audit_logger.log_action("OAUTH_MS_SUCCESS", details={"session": session_id})
        
        return render_template_string(INDEX_TEMPLATE,
                                    message="✅ Microsoft authorization successful! Scanning emails...")
    except Exception as e:
        audit_logger.log_action("OAUTH_MS_ERROR", details={"error": str(e)})
        return render_template_string(INDEX_TEMPLATE, error=f"Failed to complete Microsoft OAuth: {str(e)}")

def run_email_scan(provider, session_id):
    """
    Runs the email scan in a background thread.
    This is part of the volatile processing architecture.
    """
    try:
        key = get_token_store_key(provider, session_id)
        token_data = TOKEN_STORE.get(key)
        
        if not token_data:
            SCAN_RESULTS[provider] = {"error": "No token found", "total_files": 0, "total_records": 0, "missing_records": 0}
            return
        
        access_token = token_data['access_token']
        
        # Initialize search bot
        bot = SearchBot(provider, access_token)
        
        # Get master DB hashes for comparison
        master_hashes = get_master_patient_hashes()
        
        # Scan emails
        found_attachments = bot.scan_emails()
        
        total_records = 0
        missing_records = 0
        files_processed = []
        
        for att in found_attachments:
            try:
                df = bot.process_attachment(att['data'], att['filename'])
                if not df.empty:
                    df_analyzed = bot.identify_missing_records(df, master_hashes)
                    total = len(df_analyzed)
                    missing = int(df_analyzed['Is_Missing'].sum()) if 'Is_Missing' in df_analyzed.columns else 0
                    
                    total_records += total
                    missing_records += missing
                    files_processed.append({
                        'filename': att['filename'],
                        'subject': att.get('subject', 'N/A'),
                        'total_records': total,
                        'missing_records': missing
                    })
            except Exception as e:
                print(f"Error processing attachment {att.get('filename', 'unknown')}: {e}")
                continue
        
        # Store results in memory (volatile)
        SCAN_RESULTS[provider] = {
            "total_files": len(found_attachments),
            "files": files_processed,
            "total_records": total_records,
            "missing_records": missing_records,
            "scanned_at": datetime.now().isoformat()
        }
        
        # Clean up token immediately after scan
        del TOKEN_STORE[key]
        
        audit_logger.log_action("EMAIL_SCAN_COMPLETE", 
                               details={"provider": provider, 
                                       "files": len(found_attachments),
                                       "records": total_records,
                                       "missing": missing_records})
        
    except Exception as e:
        SCAN_RESULTS[provider] = {"error": str(e), "total_files": 0, "total_records": 0, "missing_records": 0}
        audit_logger.log_action("EMAIL_SCAN_ERROR", details={"provider": provider, "error": str(e)})

@app.route('/api/results')
def api_results():
    """API endpoint to get scan results."""
    return jsonify(SCAN_RESULTS)

@app.route('/api/clear', methods=['POST'])
def api_clear():
    """Clear scan results (volatile cleanup)."""
    SCAN_RESULTS.clear()
    return jsonify({"status": "cleared"})

@app.route('/health')
def health():
    """Health check endpoint."""
    return jsonify({
        "status": "ok",
        "uptime_seconds": (datetime.now() - SERVER_START_TIME).total_seconds(),
        "scan_results_count": len(SCAN_RESULTS),
        "tokens_in_memory": len(TOKEN_STORE)
    })

# ============================================================================
# Main
# ============================================================================

if __name__ == '__main__':
    print("=" * 60)
    print("SCD External Discovery - OAuth Callback Server")
    print("=" * 60)
    print("This server handles OAuth callbacks for email scanning.")
    print("Running on http://0.0.0.0:3000")
    print("Press Ctrl+C to stop")
    print("=" * 60)
    
    # Run Flask app
    app.run(host='0.0.0.0', port=3000, debug=False)
