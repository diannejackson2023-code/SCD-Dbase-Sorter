# SCD Dbase Sorter - Owner Handover Document
**Status: Milestone 6 (Global Launch Readiness)**
**Date: 2026-07-23**

## 1. System Delivery Summary
The SCD Dbase Sorter has transitioned from a local Excel sorting tool to a robust, enterprise-grade medical data management platform. The system is hardened for global clinical use, with a primary focus on data accuracy and security.

### Core Capabilities:
- **Intelligent Ingestion**: Automated mapping of inconsistent hospital headers with "Suggested Healing" human-in-the-loop verification.
- **Superbot Discovery**: Proactive search of local drives and email accounts to find missing SCD records.
- **Cybersecurity Gateway**: AES-256 encryption at rest, formula injection defense, and macro-stripping.
- **Global Scaling**: Multi-tenant architecture with regional tagging (e.g., SERHA, West Africa).

## 2. Milestone 5: Strict Accuracy Guardrails
As requested, "Automatic Healing" has been replaced with **Suggested Healing**.
- **The Verification Queue**: All high-risk automated corrections are diverted to a Lead queue.
- **Human-in-the-loop**: A human Lead must review the system's mapping logic before data is merged.
- **Learning Loop**: The system only "learns" new aliases after they have been explicitly verified by the Lead.

## 3. Operational Setup (Action Required)
To go live on Streamlit Community Cloud, the owner must provide the following credentials:
1.  **SMTP Credentials**: For email notifications and validation requests.
2.  **Twilio API SID/Token**: For the SMS OTP identity verification system.
3.  **PayPal Client ID**: For license billing (if applicable).
4.  **Admin Password Hash (`ADMIN_PASSWORD_HASH`)**: SHA-256 hash of the portal admin password. Default is the hash of "Protect#1$".
5.  **National Viewer Password Hash (`VIEWER_PASSWORD_HASH`)**: SHA-256 hash of the view-only password. Default is the hash of "national-view-2026".

### Portal Password Rotation
The system uses SHA-256 hashes stored in environment variables for portal access. To rotate these:
1. Generate a new SHA-256 hash for your desired password (e.g., using an online tool or `hashlib` in Python).
2. Update the `ADMIN_PASSWORD_HASH` or `VIEWER_PASSWORD_HASH` secret in the Streamlit Cloud dashboard.
3. Restart the Streamlit app to apply changes.

## 4. Role-Based Access Control
The system now supports two distinct access levels:
- **Admin**: Full access to all operations including data append, merge, and email distribution.
- **National Viewer (View-Only)**: Authorized access for national oversight. Users can view all metrics, charts, and record details across all regions but cannot modify the database or initiate outgoing communications.

## 5. Documentation References
- **[TECHNICAL_MANUAL.md (Version 1.7)](./TECHNICAL_MANUAL.md)**: Deep dive into the architecture and code modules.
- **[USER_MANUAL.md](./USER_MANUAL.md)**: Guide for clinical administrators and Leads.
- **[SANITIZATION_PROTOCOL.md](./SANITIZATION_PROTOCOL.md)**: Security rules for data ingestion.
- **[EXTERNAL_DISCOVERY_DESIGN.md](./EXTERNAL_DISCOVERY_DESIGN.md)**: Details of the Superbot and Companion App workflow.

## 6. Marketplace Readiness
The system is ready for the **cto.new Marketplace** as a "Gold Standard" template for medical record sorting. The codebase isproduction-locked with all dependencies resolved.

---
**SCD Dbase Sorter Team | 2026**
