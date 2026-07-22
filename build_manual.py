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
    "sanitization.py": "/home/team/shared/SCD_Dbase_Sorter/processor/sanitization.py",
    "discovery_api.py": "/home/team/shared/SCD_Dbase_Sorter/processor/discovery_api.py",
    "discovery_service.py": "/home/team/shared/SCD_Dbase_Sorter/processor/discovery_service.py",
    "staging_api.py": "/home/team/shared/SCD_Dbase_Sorter/processor/staging_api.py",
    "search_bot.py": "/home/team/shared/SCD_Dbase_Sorter/processor/search_bot.py",
    "word_processor.py": "/home/team/shared/SCD_Dbase_Sorter/processor/word_processor.py",
    "payments.py": "/home/team/shared/SCD_Dbase_Sorter/processor/payments.py",
    "scanner.py": "/home/team/shared/SCD_Dbase_Sorter/companion/scanner.py"
}

approvals = [
    {"date": "2026-05-19", "id": "23865abb", "member": "agent-architect", "role": "Architect", "msg": "Created comprehensive System Architecture Design for the SCD Dbase Sorter."},
    {"date": "2026-05-19", "id": "f3566313", "member": "agent-data-eng", "role": "Data Eng", "msg": "Implemented the core data processing logic (mapping.py, sorter.py)."},
    {"date": "2026-05-19", "id": "16b1e8a3", "member": "agent-comm-eng", "role": "Comm Eng", "msg": "Implemented email notification logic in processor/mailer.py."},
    {"date": "2026-05-19", "id": "cd2e03a1", "member": "agent-architect", "role": "Architect", "msg": "Developed a comprehensive Technical Manual documenting System Architecture."},
    {"date": "2026-05-19", "id": "451b2459", "member": "agent-ui-dev", "role": "UI Dev", "msg": "Developed the Streamlit dashboard (app.py) with 4-step workflow."},
    {"date": "2026-05-20", "id": "e6fc5e1a", "member": "agent-data-eng", "role": "Data Eng", "msg": "Implemented secure file handling and PII strategy (msoffcrypto support)."},
    {"date": "2026-05-20", "id": "7e40e1be", "member": "agent-ui-dev", "role": "UI Dev", "msg": "Integrated secure file handling into the dashboard."},
    {"date": "2026-05-28", "id": "ce6e55e9", "member": "agent-backend-developer", "role": "Backend Dev", "msg": "Implemented Top-Tier Cybersecurity: AES-256, Audit Logging, Input Sanitization."},
    {"date": "2026-06-14", "id": "8f93f13b", "member": "agent-architect", "role": "Architect", "msg": "Updated Discovery architecture for Local-First approach and Lead initiation."},
    {"date": "2026-07-17", "id": "ac90b651", "member": "agent-architect", "role": "Architect", "msg": "Defined Sanitization & Healing Protocol (Formula defense, Macro stripping)."},
    {"date": "2026-07-19", "id": "b5db6ad0", "member": "agent-ui-dev", "role": "UI Dev", "msg": "Built the Visual Sync Box with animated file processing."},
    {"date": "2026-07-19", "id": "b11703fa", "member": "agent-data-eng", "role": "Data Eng", "msg": "Implemented Atomic Export and Header Healing logic."},
    {"date": "2026-07-19", "id": "ceb8cc36", "member": "agent-architect", "role": "Architect", "msg": "Finalized manuals for Milestone 4: Superbot & Security Gateway."}
]

manual_content = f"""# SCD Dbase Sorter - Technical Manual
**Version 1.4 (Milestone 4 Final)**

## 1. Project Overview
The **SCD Dbase Sorter** is a Python-based automated system designed to manage medical records related to Sickle Cell Disease (SCD) across multiple hospitals. It automates the ingestion of inconsistent Excel data, sorts it into a master database, generates hospital-specific reports, and facilitates secure validation and notification via email. Milestone 4 introduces the **Superbot & Security Gateway** for proactive external discovery and structural data healing.

## 2. Core Features
- **Manual Mapping & Teaching System**: Provides a user interface to manually verify and correct column mappings. The system "learns" from manual corrections by saving new aliases to `aliases.json`.
- **Intelligent Header Mapping**: Automatically detects and maps inconsistent Excel headers (Rows 1 or 2).
- **Automated Sorting & Distribution**: Appends data to a central master database and creates individual Excel files/sheets for every hospital, sorted by year.
- **Milestone 4: Superbot Discovery**: Proactive search of local drives and email accounts using a secure, OTP-verified sequential workflow.
- **Security Gateway & Sanitization**: Automated formula injection defense, VBA macro stripping, and structural healing of broken headers.
- **Interactive Dashboard**: Streamlit-based UI for end-to-end workflow management, including a real-time **Visual Sync Box** for discovery progress.
- **Direct PayPal Billing**: Built-in payment integration for license management.

## 3. System Architecture
The system follows a modular design:
- `app.py`: Main Streamlit entry point.
- `processor/mapping.py`: Header alias mapping and ingestion logic.
- `processor/sorter.py`: Master database management and file splitting.
- `processor/mailer.py`: SMTP communication engine.
- `processor/encryption.py`: AES-256 encryption/decryption.
- `processor/logger.py`: Audit logging system.
- `processor/sanitization.py`: Input security logic (Formula/XSS protection).
- `processor/discovery_api.py`: Backend API for discovery requests and tracking.
- `processor/discovery_service.py`: Orchestrates the search bot and staging queue.
- `processor/staging_api.py`: Receives and sanitizes real-time file streams from the companion app.
- `processor/search_bot.py`: Handles OAuth email scanning (Gmail/Outlook).
- `processor/word_processor.py`: Scans Word (.docx) files for SCD records.
- `processor/payments.py`: PayPal payment integration.
- `companion/scanner.py`: Downloadable agent for secure local filesystem scanning.

## 4. Milestone 4: Superbot & Security Gateway
Milestone 4 expands the discovery capabilities with a dedicated security layer and real-time visualization.

### 4.1 Companion Scanner App
- **Function**: A downloadable, single-file executable for recipients to securely scan local drives (Desktop, Documents, Downloads) without requiring full system access for the main web app.
- **Security**: Performs in-memory scanning. Strips macros and neutralizes formulas before upload using the team's sanitization protocol.
- **Protocol**: Communicates with the Staging API via secure HTTPS POST requests.

### 4.2 Staging API & Live Queue
- **Architecture**: A secure, intermediate storage area (Staging directory) where discovered records are held before final merge.
- **Workflow**: 1. Discovery -> 2. Sanitization -> 3. Staging Queue -> 4. Lead Review -> 5. Atomic Export -> 6. Master DB.
- **Live Queue**: Maintains a status file (`queue.json`) for real-time tracking, enabling the Visual Sync Box.

### 4.3 Structural Healing & Security Protocol
- **Formula Neutralization**: Fixes malicious formulas by prepending a single quote (`'`) to strings starting with `=`, `+`, `-`, or `@`.
- **VBA Macro Stripping**: Converts `.xlsm` and `.docm` files to macro-free versions during ingestion.
- **Header Healing**: Uses fuzzy matching (Levenshtein distance <= 2) and deep searching (up to Row 10) to repair misspelled or shifted headers.
- **Atomic Export**: Ensures files are deleted from staging ONLY after successful database write and verification.

### 4.4 Visual Sync Box
- **Real-Time Progress**: Shows files being staged by the companion app and their security status (Healthy, Warning, Blocked).
- **Disappearing Animation**: Files are visually removed from the box one by one as they are successfully merged into the Master Database.

## 5. Chronological Approval Log
Detailed record of project milestones and approvals.

| Date | Task ID | Member | Role | Status | Message Body |
| :--- | :--- | :--- | :--- | :--- | :--- |
"""

for a in approvals:
    manual_content += f"| {a['date']} | {a['id']} | {a['member']} | {a['role']} | Approved | {a['msg']} |\n"

manual_content += """
## 6. Deployment Instructions
1. Install dependencies: `pip install -r requirements.txt`
2. Run the application: `streamlit run app.py`
3. Access the Dashboard to initiate Discovery requests or upload files.
4. For production, use the provided `Dockerfile` for containerized deployment.
5. Configure `PAYPAL_CLIENT_ID` and SMTP settings in the environment/sidebar.

---

## 7. Source Code Appendix

"""

for name, path in files.items():
    manual_content += f"### 7.{list(files.keys()).index(name)+1} {path}\n"
    manual_content += "```python\n"
    manual_content += read_file(path)
    manual_content += "\n```\n\n"

with open("/home/team/shared/SCD_Dbase_Sorter/TECHNICAL_MANUAL.md", "w") as f:
    f.write(manual_content)

print("TECHNICAL_MANUAL.md has been rebuilt with full source code and Milestone 4 details.")
