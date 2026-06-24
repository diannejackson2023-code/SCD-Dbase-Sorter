# SCD Dbase Sorter - Security Hardening Documentation

This document outlines the security measures implemented to protect the SCD Dbase Sorter application, its data, and its users.

## 1. Data Protection (At-Rest)
All sensitive data files, including the Master Database and individual hospital Excel sheets, are encrypted at rest.
- **Mechanism**: AES-256 encryption using the `cryptography` Fernet implementation.
- **Key Management**: A master key is generated and stored securely in `data/config/.master.key`.
- **Implementation**: Files are decrypted in-memory only when needed for processing or display, and re-encrypted immediately upon being written back to disk.

## 2. PII Security & Masking
Patient Identifiable Information (PII) is protected through automated masking.
- **Masking Logic**: Patient Names are partially obfuscated (e.g., "John Doe" becomes "J**n D*e").
- **Ingestion**: Masking is applied at the point of ingestion in `processor/mapping.py`, ensuring that raw names are never stored unmasked in the master database.

## 3. Input Sanitization
To prevent malicious data from compromising the system or its users, all incoming data is sanitized.
- **Excel Formula Injection**: Values starting with triggers like `=`, `+`, `-`, or `@` are prefixed with a single quote to prevent execution in spreadsheet software.
- **XSS Protection**: All string data is HTML-escaped before being rendered in the Streamlit dashboard.
- **Implementation**: Located in `processor/sanitization.py`.

## 4. Audit Logging
Comprehensive audit logs are maintained to track all significant actions within the system.
- **Logged Events**: File uploads, data processing, email transmissions, document exports, and system errors.
- **Format**: JSONL (JSON Lines) for easy machine parsing and integration with SIEM tools.
- **Path**: `data/logs/audit_log.jsonl`.

## 5. Transport Security (TLS/HTTPS)
- **Email**: All email communications (validation requests and finalized data) are sent via SMTP with mandatory TLS (`STARTTLS`).
- **Web (HTTPS)**: It is recommended to deploy the Streamlit dashboard behind a reverse proxy (like Nginx) or a CDN (like Cloudflare) to enforce HTTPS.

## 6. Security Headers Strategy
While Streamlit doesn't natively support setting arbitrary HTTP headers, the following headers should be configured at the proxy level (Nginx/Cloudflare):
- `Content-Security-Policy`: Restricts where scripts and resources can be loaded from.
- `Strict-Transport-Security`: Enforces HTTPS for a specified duration.
- `X-Content-Type-Options: nosniff`: Prevents the browser from MIME-sniffing the response.
- `X-Frame-Options: DENY`: Prevents clickjacking attacks.
- `X-XSS-Protection: 1; mode=block`: Enables the browser's XSS filter.

## 7. DDoS Protection & Rate Limiting
- **Cloudflare Integration**: We recommend routing the application traffic through Cloudflare for:
    - **DDoS Protection**: Mitigation of volumetric attacks.
    - **WAF (Web Application Firewall)**: Protection against common web exploits.
    - **Rate Limiting**: Throttling requests from single IP addresses to prevent abuse.
- **Internal Rate Limiting**: For small-scale deployments, Nginx `limit_req` module can be used to rate-limit the dashboard and processing endpoints.

## 8. Secure Session Management
- **Streamlit Sessions**: The application uses Streamlit's built-in session state for managing user data during a browser session.
- **Recommendations**: For multi-user environments, integrate with an OIDC/SAML provider or use Streamlit's authentication features to restrict access.

## 9. External Discovery Security
The proactive discovery of SCD records from external sources follows a high-security, privacy-preserving workflow.
- **Token-Based Access**: Recipients access the discovery portal via unique, cryptographically secure (32-byte `token_urlsafe`), time-limited tokens.
- **Sequential Scan Workflow**:
    1. **Local-First Scan**: Recipients can scan their local files without needing to provide identity details initially, lowering the barrier for participation.
    2. **Gated Data Comparison**: Local records are processed in-memory but are **not** compared against the Master Database (hashing/lookup) until the recipient's identity is verified.
    3. **Identity Verification (Twilio Verify)**: Mandatory SMS OTP verification is required to unlock comparison results and initiate the Email Scan (OAuth).
- **Privacy-Preserving Comparison**: Discovered records are hashed (SHA-256) and compared against a pre-computed index of Master Database hashes. This ensures the contents of the Master Database are never exposed to the recipient or the discovery portal.
- **OAuth Security**: Email scans use read-only OAuth scopes. Access tokens are kept in-session and never stored on disk.
- **Volatile Processing**: All scan results are handled in-memory and are subject to deletion once the session ends or the token expires.
- **Bulk Initiation**: Leads can initiate discovery for multiple recipients efficiently. Dispatch is throttled and handled in background threads to ensure system stability.

---
*Last Updated: 2026-06-15*
