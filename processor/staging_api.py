import os
import json
import datetime
from flask import Flask, request, jsonify

app = Flask(__name__)

# Manual CORS implementation since flask-cors is not available
@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
    return response

STAGING_BASE_DIR = "/home/team/shared/SCD_Dbase_Sorter/data/staging"
QUEUE_FILE = os.path.join(STAGING_BASE_DIR, "queue.json")

def update_queue_status(token, filename, status, details=None):
    os.makedirs(STAGING_BASE_DIR, exist_ok=True)
    queue = {}
    if os.path.exists(QUEUE_FILE):
        try:
            with open(QUEUE_FILE, 'r') as f:
                queue = json.load(f)
        except Exception:
            queue = {}
    
    if token not in queue:
        queue[token] = []
    
    found = False
    for item in queue[token]:
        if item.get('filename') == filename:
            item['status'] = status
            item['updated_at'] = datetime.datetime.now().isoformat()
            if details:
                item['details'] = details
            found = True
            break
    
    if not found:
        queue[token].append({
            "filename": filename,
            "status": status,
            "updated_at": datetime.datetime.now().isoformat(),
            "details": details
        })
    
    with open(QUEUE_FILE, 'w') as f:
        json.dump(queue, f, indent=4)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({
        "status": "healthy", 
        "timestamp": datetime.datetime.now().isoformat(),
        "staging_dir": STAGING_BASE_DIR
    }), 200

@app.route('/api/staging/init', methods=['POST', 'OPTIONS'])
def init_staging():
    if request.method == 'OPTIONS':
        return jsonify({}), 200
        
    data = request.json or {}
    token = data.get('token')
    if not token:
        return jsonify({"error": "Token required"}), 400
    
    token_dir = os.path.join(STAGING_BASE_DIR, token)
    os.makedirs(token_dir, exist_ok=True)
    
    update_queue_status(token, "SESSION", "INITIALIZED", "Staging session started")
    
    return jsonify({"message": "Staging initialized", "token": token}), 200

def scan_malware(file_storage):
    """
    Conceptual hook for malware scanning.
    In a real-world scenario, this would interface with a security gateway
    like ClamAV or a cloud-based scanning service.
    """
    # Example: check file extension or magic bytes
    # if file_storage.filename.endswith('.exe'): return True
    return False

@app.route('/api/staging/upload/<token>', methods=['POST', 'OPTIONS'])
def upload_to_staging(token):
    if request.method == 'OPTIONS':
        return jsonify({}), 200
        
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    # conceptual malware scan hook
    if scan_malware(file):
        update_queue_status(token, file.filename, "REJECTED", "Security policy violation (Potential Malware)")
        return jsonify({"error": "Malware detected"}), 403
    
    token_dir = os.path.join(STAGING_BASE_DIR, token)
    os.makedirs(token_dir, exist_ok=True)
    
    file_path = os.path.join(token_dir, file.filename)
    file.save(file_path)
    
    update_queue_status(token, file.filename, "STAGED", f"Received at {datetime.datetime.now().strftime('%H:%M:%S')}")
    
    return jsonify({
        "message": "File staged successfully",
        "filename": file.filename,
        "token": token
    }), 201

@app.route('/api/staging/status/<token>', methods=['GET'])
def get_staging_status(token):
    if not os.path.exists(QUEUE_FILE):
        return jsonify({"token": token, "queue": []}), 200
    
    try:
        with open(QUEUE_FILE, 'r') as f:
            queue = json.load(f)
    except Exception:
        return jsonify({"error": "Failed to read queue"}), 500
    
    return jsonify({
        "token": token,
        "queue": queue.get(token, [])
    }), 200

if __name__ == '__main__':
    # Default to port 5000, but allow override
    port = int(os.environ.get("STAGING_API_PORT", 5000))
    print(f"Staging API running on port {port}...")
    app.run(host='0.0.0.0', port=port)
