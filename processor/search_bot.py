"""
Search Bot Logic for External SCD Data Discovery.
Conceptual implementation for Gmail and Outlook scanning.
"""

import os
import io
import pandas as pd
import sys
# Handle both relative and absolute imports
try:
    from .mapping import load_and_map_data, find_master_match, MASTER_HEADINGS
    from .word_processor import process_word_file, scan_text_for_keywords
except ImportError:
    from mapping import load_and_map_data, find_master_match, MASTER_HEADINGS
    from word_processor import process_word_file, scan_text_for_keywords

class SearchBot:
    def __init__(self, provider, access_token):
        self.provider = provider # 'google' or 'microsoft'
        self.access_token = access_token
        self.keywords = ["SCD", "Sickle Cell", "SCD Data", "Patient Records"]
        
    def scan_emails(self):
        """
        Lists and filters emails by keywords.
        Returns a list of matching email metadata.
        """
        if self.provider == 'google':
            return self._scan_gmail()
        elif self.provider == 'microsoft':
            return self._scan_outlook()
        return []

    def _scan_gmail(self):
        from googleapiclient.discovery import build
        from google.oauth2.credentials import Credentials
        import base64

        creds = Credentials(self.access_token)
        service = build('gmail', 'v1', credentials=creds)
        
        query = " OR ".join(self.keywords)
        results = service.users().messages().list(userId='me', q=query).execute()
        messages = results.get('messages', [])
        
        found_attachments = []
        for msg in messages:
            m = service.users().messages().get(userId='me', id=msg['id']).execute()
            payload = m.get('payload', {})
            parts = payload.get('parts', [])
            for part in parts:
                if part.get('filename'):
                    # Check extension
                    if any(part['filename'].lower().endswith(ext) for ext in ['.xlsx', '.xls', '.docx', '.pdf']):
                        att_id = part['body'].get('attachmentId')
                        if att_id:
                            attachment = service.users().messages().attachments().get(
                                userId='me', messageId=msg['id'], id=att_id).execute()
                            data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
                            found_attachments.append({
                                'filename': part['filename'],
                                'data': data,
                                'subject': next((h['value'] for h in payload.get('headers', []) if h['name'] == 'Subject'), 'No Subject')
                            })
        return found_attachments

    def _scan_outlook(self):
        import requests
        
        headers = {'Authorization': f'Bearer {self.access_token}'}
        query = " OR ".join(f"contains(subject,'{k}')" for k in self.keywords)
        endpoint = f"https://graph.microsoft.com/v1.0/me/messages?$filter={query}&$expand=attachments"
        
        response = requests.get(endpoint, headers=headers)
        if response.status_code != 200:
            print(f"Error scanning Outlook: {response.text}")
            return []
            
        messages = response.json().get('value', [])
        found_attachments = []
        for msg in messages:
            for att in msg.get('attachments', []):
                if att.get('@odata.type') == '#microsoft.graph.fileAttachment':
                    if any(att['name'].lower().endswith(ext) for ext in ['.xlsx', '.xls', '.docx', '.pdf']):
                        import base64
                        data = base64.b64decode(att['contentBytes'])
                        found_attachments.append({
                            'filename': att['name'],
                            'data': data,
                            'subject': msg.get('subject', 'No Subject')
                        })
        return found_attachments

    def process_attachment(self, attachment_data, filename, password=None):
        """
        Processes a downloaded attachment based on its extension.
        """
        ext = os.path.splitext(filename)[1].lower()
        
        # Save to temp file for processing if needed, but prefer in-memory
        file_like = io.BytesIO(attachment_data)
        
        try:
            if ext in ['.xlsx', '.xls']:
                return load_and_map_data(file_like, password=password)
            elif ext == '.docx':
                # Temporary save since docx library usually requires a path or file-like
                return process_word_file(file_like)
            elif ext == '.pdf':
                return self.process_pdf(file_like)
        except Exception as e:
            if "password" in str(e).lower() or "encrypted" in str(e).lower():
                return "PASSWORD_REQUIRED"
            return f"ERROR: {str(e)}"
        return pd.DataFrame()

    def process_pdf(self, file_like):
        """
        Extracts tables from a PDF file.
        """
        import pdfplumber
        all_data = []
        try:
            with pdfplumber.open(file_like) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        # Attempt to map columns
                        mapped_cols = {}
                        for i, col in enumerate(df.columns):
                            match = find_master_match(col)
                            if match:
                                mapped_cols[col] = match
                        if mapped_cols:
                            df = df.rename(columns=mapped_cols)
                            # Keep only master columns
                            df = df[[c for c in df.columns if c in MASTER_HEADINGS]]
                            all_data.append(df)
            if all_data:
                return pd.concat(all_data, ignore_index=True)
        except Exception as e:
            return f"DAMAGED_PDF: {str(e)}"
        return pd.DataFrame()

    def identify_missing_records(self, found_df, master_db_hashes):
        """
        Compares found records against the pre-computed hashes of the Master DB.
        Assumes found_df has a 'Patient_ID' column.
        """
        import hashlib
        
        if 'Patient_ID' not in found_df.columns:
            return found_df
            
        def is_missing(pid):
            if pd.isna(pid): return False
            h = hashlib.sha256(str(pid).strip().encode()).hexdigest()
            return h not in master_db_hashes
            
        found_df['Is_Missing'] = found_df['Patient_ID'].apply(is_missing)
        return found_df
