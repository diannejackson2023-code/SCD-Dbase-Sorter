import os
import sys
import time
import json
import requests
import pandas as pd
import html
import tempfile
from datetime import datetime

# --- Sanitization Protocol (Shared Logic) ---
def sanitize_value(val):
    if not isinstance(val, str):
        return val
    # Prevent Excel Formula Injection
    if val.startswith(('=', '+', '-', '@')):
        val = "'" + val
    # Prevent XSS
    val = html.escape(val)
    return val

def sanitize_dataframe(df):
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].apply(sanitize_value)
    return df
# --------------------------------------------

class CompanionScanner:
    def __init__(self, api_url, token):
        self.api_url = api_url.rstrip('/')
        self.token = token
        self.scanned_files = []

    def log(self, message):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def malware_check(self, file_path):
        """
        Conceptual malware scanning hook.
        In a production environment, this would interface with a local AV engine.
        """
        self.log(f"Running safety check on {os.path.basename(file_path)}...")
        # Conceptual placeholder for signature/heuristic analysis
        return False

    def scan_folders(self):
        self.log("Initializing local filesystem scan...")
        target_folders = []
        if sys.platform == "win32":
            user_profile = os.environ.get("USERPROFILE")
            if user_profile:
                target_folders = [
                    os.path.join(user_profile, "Desktop"),
                    os.path.join(user_profile, "Documents"),
                    os.path.join(user_profile, "Downloads")
                ]
        else:
            home = os.path.expanduser("~")
            target_folders = [
                os.path.join(home, "Desktop"),
                os.path.join(home, "Documents"),
                os.path.join(home, "Downloads")
            ]
        
        found_files = []
        for folder in target_folders:
            if not os.path.exists(folder):
                continue
            self.log(f"Scanning {folder}...")
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith(('.xlsx', '.xls', '.docx')):
                        found_files.append(os.path.join(root, file))
        
        return found_files

    def process_and_upload(self, file_path):
        filename = os.path.basename(file_path)
        
        # 1. Malware Check
        if self.malware_check(file_path):
            self.log(f"❌ SECURITY ALERT: {filename} failed malware check. Skipping.")
            return False

        # 2. Sanitization
        self.log(f"Applying sanitization protocol to {filename}...")
        upload_path = file_path
        temp_file = None
        
        if filename.lower().endswith(('.xlsx', '.xls')):
            try:
                df = pd.read_excel(file_path)
                df = sanitize_dataframe(df)
                # Create a proper temp file
                fd, upload_path = tempfile.mkstemp(suffix=f"_{filename}")
                os.close(fd)
                df.to_excel(upload_path, index=False)
                temp_file = upload_path
            except Exception as e:
                self.log(f"⚠️ Error sanitizing {filename}: {e}")
                return False

        # 3. Stream to Staging API
        self.log(f"Streaming {filename} to staging server...")
        try:
            with open(upload_path, 'rb') as f:
                files = {'file': (filename, f)}
                response = requests.post(f"{self.api_url}/api/staging/upload/{self.token}", files=files)
                if response.status_code == 201:
                    self.log(f"✅ Successfully staged {filename}")
                    return True
                else:
                    self.log(f"❌ Server error ({response.status_code}): {response.text}")
                    return False
        except Exception as e:
            self.log(f"❌ Network error: {e}")
            return False
        finally:
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass

def main():
    print("========================================")
    print("   SCD Companion Scanner - Superbot v1  ")
    print("========================================\n")
    
    # In a real build, these would be passed via command line or a config file
    API_URL = os.environ.get("SCD_API_URL", "http://localhost:5000")
    TOKEN = sys.argv[1] if len(sys.argv) > 1 else "DISCOVERY_TOKEN_REQUIRED"
    
    if TOKEN == "DISCOVERY_TOKEN_REQUIRED":
        print("Usage: scanner.py <token>")
        print("\nNote: Please obtain your secure token from the discovery email.")
        return

    scanner = CompanionScanner(API_URL, TOKEN)
    
    # Init session
    try:
        response = requests.post(f"{API_URL}/api/staging/init", json={"token": TOKEN}, timeout=5)
        if response.status_code != 200:
            print(f"Error: Staging API rejected initialization: {response.text}")
            return
    except Exception as e:
        print(f"Error: Could not connect to staging server at {API_URL}")
        print(f"Details: {e}")
        return

    files = scanner.scan_folders()
    print(f"\nFound {len(files)} potential record files.\n")
    
    success_count = 0
    for file in files:
        if scanner.process_and_upload(file):
            success_count += 1
        time.sleep(0.5)

    print(f"\nScan complete. {success_count} files successfully staged.")
    print("You can now return to the dashboard to review and merge the data.")
    
    if sys.platform == "win32":
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
