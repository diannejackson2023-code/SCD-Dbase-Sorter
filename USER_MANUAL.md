# SCD Dbase Sorter - User Manual
**Version 1.4 (Milestone 4 Update)**

## 1. Overview
Welcome to the **SCD Dbase Sorter**. This tool is designed to manage Sickle Cell Disease (SCD) data across multiple hospitals. It automates importing, sorting, and discovering new records while maintaining the highest levels of medical data security.

## 2. Core Dashboard Workflow

### Step 1: Upload Data
- Upload your Excel file. If it's encrypted, check the password box.
- **Teaching**: If the system misidentifies a column, correct it and click **"Confirm & Teach"**. The system "learns" your specific hospital's terminology.

### Step 2: Process & Sort
- Click **"Process & Sort"**. This appends data to the encrypted Master Database and generates hospital-specific files in the `data/hospitals` folder.

### Step 3: Review & Email
- Review statistics in the **Overview** tab.
- Use the **Validate & Email** section to send reports to hospitals or validation requests to experts with one click.

---

## 3. SCD Data Discovery (Milestone 4 Superbot)

The Superbot proactively finds missing records in external sources.

### 3.1 Initiating Discovery (Lead)
- Use the **Discovery Initiation** form to send a secure link to a recipient.
- Supports **Bulk Initiation**: You can upload a list of contacts to search multiple sources at once.

### 3.2 Recipient Workflow (The Companion App)
When a recipient opens the link:
1. **Verification**: They must verify their identity via a 6-digit SMS OTP code.
2. **The Companion App**: Recipients are prompted to download a small "Companion Scanner".
3. **Local Scan**: The app scans their computer (Desktop, Documents, Downloads) for SCD-related files.
4. **Security Check**: The app automatically strips dangerous macros and neutralizes malicious formulas before sending any data.
5. **Email Search**: Recipients can optionally authorize a search of their Gmail or Outlook attachments.

### 3.3 Visual Sync Box (Real-Time Monitoring)
On your dashboard, you will see the **Visual Sync Box**:
- **Live Queue**: Records appearing in real-time as the bot finds them.
- **Security Shields**: Green icons confirm that a record has been malware-scanned and "healed" (headers fixed).
- **Atomic Merge**: Once you review the findings, click "Approve" to safely merge them into the Master Database.

---

## 4. Intelligence & Structural Healing
One of the most powerful features of the system is **Structural Healing**:
- **Misspelled Headers**: If a hospital sends a file with "Ptnt ID" instead of "Patient_ID", the system recognizes it and fixes it automatically based on its learned Knowledge Base.
- **Damaged Files**: If a header row is missing, the system scans the data types to "guess" the column meaning, ensuring no data is lost.

---

## 5. Security & Privacy
- **AES-256 At-Rest Encryption**: Your data is never stored in plain text.
- **PII Masking**: Patient names are masked (e.g., "John Doe" -> "J**n D*e") to protect privacy.
- **Audit Logging**: Every access and change is recorded in a secure log file.

---

## 6. Final Project Status
The SCD Dbase Sorter is now a fully matured system.
- **Sorting Engine**: Fully automated and "teachable".
- **Security**: Enterprise-grade encryption and sanitization.
- **Discovery**: Sequential, OTP-verified search bot with companion app support.
- **Healing**: Advanced structural repair of inconsistent data.

---
**SCD Dbase Sorter Team | 2026**
