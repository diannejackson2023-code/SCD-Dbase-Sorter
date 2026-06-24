# SCD Dbase Sorter

Build a high-security, automated medical database system for Sickle Cell Disease (SCD) records using Excel-based source data. The system streamlines data ingestion, sorting, validation, and proactive data discovery.

## 1. Core Objectives
- **Automated Sorting**: Ingest static Excel sheets and automatically sort data under matching headings using a persistent Knowledge Base of aliases.
- **Hierarchical Organization**: Sort and categorize data by hospital and year.
- **Hospital Isolation**: Automatically create and maintain separate sheets/files for each hospital.
- **Secure Communication**: Integrated "push-button" email system for validation requests (to validators) and finalized reports (to hospitals).
- **Proactive Discovery**: A secure mechanism to find missing SCD records in external local files and email accounts via a sequential search bot.

## 2. Key Features
- **Intelligent Mapping**: A `mapping.py` engine that "learns" column aliases from manual user corrections, stored in `aliases.json`.
- **Top-Tier Security**:
    - AES-256 At-Rest Encryption for all data files.
    - JSONL Audit Logging for accountability.
    - Input Sanitization (Formula Injection & XSS defense).
    - Session-based Master Password gating.
- **Bot-Driven Data Discovery**:
    - **Memorization**: Bot reads and memorizes encrypted master records for privacy-preserving comparison.
    - **Sequential Search**: Secure flow (OTP -> Computer Search -> Email Search) for Lead and external recipients.
    - **Duplicate Detection**: Instant feedback on whether files are already in the database.
    - **Encrypted File Support**: Automatic password prompting and memorization for protected Excels.

## 3. Tech Stack
- **Platform**: Streamlit-based dashboard for centralized management.
- **Processing Engine**: Python (Pandas, Openpyxl, python-docx) for robust data manipulation.
- **Communication**: Integrated SMTP engine with TLS/SSL enforcement.
- **Identity & Verification**: Twilio Verify for SMS OTP; Google/Microsoft OAuth for email searching.

## 4. Team
- `architect`: Oversees design and technology choices.
- `backend-developer`: Hardens infrastructure and implements core security.
- `data_eng`: Implements Excel processing, sorting, and mapping logic.
- `comm_eng`: Implements email automation and OAuth integrations.
- `ui_dev`: Builds the Streamlit dashboard and Recipient Portal.

---
---
## 5. Documentation & Delivery
- **TECHNICAL_MANUAL.pdf**: Complete developer and architecture guide.
- **USER_MANUAL.pdf**: End-user guide for the Streamlit dashboard and search bot.
- **CHAT_HISTORY.pdf**: Full transcript of the project development and logic discussions.
- **DEPLOYMENT_SECURITY.pdf**: Hardening and production deployment guide.
- (Markdown versions are also available in the root directory).

**SCD Dbase Sorter Team**
