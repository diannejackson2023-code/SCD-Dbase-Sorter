import os
import json

def read_file(path):
    if not os.path.exists(path): return f"# File {path} not found"
    with open(path, 'r') as f:
        return f.read()

files = {
    "app.py": "/home/team/shared/SCD_Dbase_Sorter/app.py",
    "mapping.py": "/home/team/shared/SCD_Dbase_Sorter/processor/mapping.py",
    "sorter.py": "/home/team/shared/SCD_Dbase_Sorter/processor/sorter.py",
    "mailer.py": "/home/team/shared/SCD_Dbase_Sorter/processor/mailer.py",
    "encryption.py": "/home/team/shared/SCD_Dbase_Sorter/processor/encryption.py",
    "logger.py": "/home/team/shared/SCD_Dbase_Sorter/processor/logger.py",
    "sanitization.py": "/home/team/shared/SCD_Dbase_Sorter/processor/sanitization.py"
}

approvals = [
    {"date": "2026-05-19", "id": "23865abb", "member": "agent-architect", "role": "Architect", "msg": "I have created a comprehensive System Architecture Design for the SCD Dbase Sorter. The design includes: System Overview, Excel Structure, Mapping Logic, Data Distribution, Email Workflow, Tech Stack, and Directory Structure."},
    {"date": "2026-05-19", "id": "f3566313", "member": "agent-data-eng", "role": "Data Eng", "msg": "Implemented the core data processing logic. Created mapping.py (header mapping Rows 1/2, alias support, email merging) and sorter.py (master database updates, autosorting, per-hospital file/sheet generation)."},
    {"date": "2026-05-19", "id": "16b1e8a3", "member": "agent-comm-eng", "role": "Comm Eng", "msg": "Implemented email notification logic in processor/mailer.py. Includes email lookups, validation requests, finalized data emails with Excel attachments, and configurable SMTP settings."},
    {"date": "2026-05-19", "id": "cd2e03a1", "member": "agent-architect", "role": "Architect", "msg": "Developed a comprehensive Technical Manual documenting System Architecture, Data Processing, Sorting/Distribution, Email Automation, and UI Guide."},
    {"date": "2026-05-19", "id": "451b2459", "member": "agent-ui-dev", "role": "UI Dev", "msg": "Developed the Streamlit dashboard (app.py) with 4-step workflow: File Upload, Process & Sort, Data Visualization, and Validate & Email."},
    {"date": "2026-05-20", "id": "e6fc5e1a", "member": "agent-data-eng", "role": "Data Eng", "msg": "Implemented secure file handling and PII strategy. Added password-protected Excel support (msoffcrypto) and Patient Name masking (J**n D*e)."},
    {"date": "2026-05-20", "id": "7e40e1be", "member": "agent-ui-dev", "role": "UI Dev", "msg": "Integrated secure file handling into the dashboard. Added password fields, smart detection, PII security sidebar, and masking notices."},
    {"date": "2026-05-28", "id": "ce6e55e9", "member": "agent-backend-developer", "role": "Backend Dev", "msg": "Implemented Top-Tier Cybersecurity: AES-256 at-rest encryption, Audit Logging (JSONL), Input Sanitization (Excel/XSS), and TLS enforcement for emails."}
]

manual_content = f"""# SCD Dbase Sorter - Technical Manual
**Version 1.3**

## 1. Project Overview
The **SCD Dbase Sorter** is a Python-based automated system designed to manage medical records related to Sickle Cell Disease (SCD) across multiple hospitals. It automates the ingestion of inconsistent Excel data, sorts it into a master database, generates hospital-specific reports, and facilitates secure validation and notification via email.

## 2. Core Features
- **Intelligent Header Mapping**: Automatically detects and maps inconsistent Excel headers (Rows 1 or 2).
- **Automated Sorting & Distribution**: Appends data to a central master database and creates individual Excel files/sheets for every hospital, sorted by year.
- **Secure Email Automation**: Built-in SMTP engine for sending validation requests and finalized data.
- **Top-Tier Cybersecurity**:
  - **AES-256 At-Rest Encryption**: All database and hospital files are encrypted on disk.
  - **Audit Logging**: Comprehensive tracking of all system actions in JSONL format.
  - **Input Sanitization**: Protection against Excel Formula Injection and XSS.
  - **PII Masking**: Automatic masking of sensitive patient names.
  - **TLS/SSL Encryption**: Secure email transport.
- **Interactive Dashboard**: Streamlit-based UI for end-to-end workflow management.

## 3. System Architecture
The system follows a modular design:
- `app.py`: Main Streamlit entry point.
- `processor/mapping.py`: Header alias mapping and ingestion logic.
- `processor/sorter.py`: Master database management and file splitting.
- `processor/mailer.py`: SMTP communication engine.
- `processor/encryption.py`: AES-256 encryption/decryption.
- `processor/logger.py`: Audit logging system.
- `processor/sanitization.py`: Input security logic.

## 4. Chronological Approval Log
Detailed record of project milestones and approvals.

| Date | Task ID | Member | Role | Status | Message Body |
| :--- | :--- | :--- | :--- | :--- | :--- |
"""

for a in approvals:
    manual_content += f"| {a['date']} | {a['id']} | {a['member']} | {a['role']} | Approved | {a['msg']} |\n"

manual_content += """
## 5. Deployment Instructions
1. Install dependencies: `pip install -r requirements.txt`
2. Run the application: `streamlit run app.py`
3. Configure SMTP settings in the sidebar.
4. Ensure `data/config/hospital_emails.csv` is populated with correct contact info.

---

## 6. Source Code Appendix

"""

for name, path in files.items():
    manual_content += f"### 6.{list(files.keys()).index(name)+1} {path}\n"
    manual_content += "```python\n"
    manual_content += read_file(path)
    manual_content += "\n```\n\n"

with open("/home/team/shared/SCD_Dbase_Sorter/TECHNICAL_MANUAL.md", "w") as f:
    f.write(manual_content)

print("TECHNICAL_MANUAL.md has been built with full source code and detailed approval logs.")
