# SCD Dbase Sorter: Hardening & Deployment Handoff Document

**Project Status:** Production-Ready (Core Logic & Security)
**Target Platform:** Hercules / Enterprise Environment
**Data Sensitivity:** High (Medical PII)

## 1. Overview
The SCD Dbase Sorter is a high-security Streamlit application designed for automated ingestion, mapping, and discovery of Sickle Cell Disease (SCD) records. It uses a "Master-First" memorization approach to allow privacy-preserving comparisons between external data and the Master Database.

## 2. Technical Stack
- **Dashboard:** Streamlit
- **Data Engine:** Python (Pandas, Openpyxl, msoffcrypto-py)
- **Word/PDF Processing:** python-docx, pdfplumber
- **Security:** AES-256 (pycryptodome), SHA-256 Hashing, Twilio SMS OTP
- **Intelligence:** Persistent Alias Engine (`aliases.json`)

## 3. Core Security Features (Implemented)
- **At-Rest Encryption:** The Master Database and all individual hospital files are AES-256 encrypted.
- **Identity Verification:** All sensitive Discovery operations are gated by 6-digit SMS OTP (Twilio).
- **Master Password Gating:** Write operations and PII viewing require session-based Master Password authorization.
- **Audit Logging:** Every file detection, comparison, and database change is logged in `data/logs/audit_log.jsonl`.
- **Privacy-First Search:** Comparisons use SHA-256 hashes. The Search Bot identifies "new" records without exposing existing PII to the recipient.
- **Docker Ready:** The project includes a `Dockerfile` for one-click deployment in a sealed environment.
- **System Health Monitor:** A built-in check in `app.py` ensures the system is initialized and secure before operation.

## 4. Hardening Requirements for Deployment Team
The following infrastructure-level hardening steps are required for production deployment:

### A. Secret Management
- **Master Password:** Currently handled via session state. Recommend moving to an Enterprise Vault (e.g., HashiCorp Vault or AWS Secrets Manager).
- **API Keys:** Move Twilio credentials and SMTP passwords from `secrets.toml` to a managed secret environment.

### B. Network Security
- **SSL/TLS:** The application must be served exclusively over HTTPS.
- **OAuth Callbacks:** Ensure the `oauth_callback_server.py` is configured with a static, secured redirect URI for Gmail/Outlook integrations.

### C. Containerization
- Recommend deploying via **Docker**. A base image of `python:3.11-slim` is sufficient.
- **Persistent Storage:** Ensure the `/data/` directory is mapped to a secure, encrypted persistent volume.

### D. Input Sanitization
- While the app implements basic Formula Injection defense, the deployment environment should implement WAF (Web Application Firewall) rules to filter malicious payloads at the edge.

## 5. Key File References
- `/app.py`: Main entry point.
- `/processor/encryption.py`: Core AES implementation.
- `/processor/search_bot.py`: Discovery and PDF/Word logic.
- `/TECHNICAL_MANUAL.pdf`: Full architecture documentation.
- `/DEPLOYMENT_SECURITY.pdf`: Specific hardening guide.

## 6. Handoff Checklist
- [ ] Connect repository to GitHub for code transfer.
- [ ] Provision Twilio Account (SID/Token) for SMS OTP.
- [ ] Provision SMTP Server for email delivery.
- [ ] Setup persistent, encrypted volume for `/data/database/`.

**SCD Dbase Sorter Team**
