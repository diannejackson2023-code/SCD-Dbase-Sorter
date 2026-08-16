# SCD Dbase Sorter - User Manual
**Version 1.8 (Multi-Entity Portal Update)**

## 1. Overview
Welcome to the **SCD Dbase Sorter**. This tool is designed to manage Sickle Cell Disease (SCD) data across multiple hospitals. It automates importing, sorting, and discovering new records while maintaining the highest levels of medical data security.

## 2. Multi-Entity Portal Access
The system now features a professional **Multi-Entity Portal** landing page to accommodate different health authorities and units.

### 2.1 Accessing the System
1. **Portal Entry**: Upon launching the application, you will see a list of authorized entities (SERHA, SRHA, NERHA, WRHA, UWI, MoHW).
2. **Sign In**: Click the **"Sign In"** button next to your organization.
3. **Authorization**: Enter your secure password to unlock the dashboard.

### 2.2 Access Roles
The system supports two access levels for different organizational needs:

#### 2.2.1 Admin Role (RHAs & UWI)
The **Admin Role** has full control over the system. This includes:
- Uploading and processing new data files.
- Teaching the system new header mappings.
- Sending validation and finalized data emails to hospitals.
- Initiating and approving Data Discovery requests.
- Managing SMTP and System configuration.

#### 2.2.2 National Oversight Mode (MoHW)
The **National Oversight Mode** is designed for view-only monitoring by the Ministry of Health & Wellness.
- **Access**: Log in as MoHW using the National Viewer password.
- **Capabilities**: View all dashboards, metrics, hospital breakdowns, and discovery tracking status.
- **Restrictions**: Cannot upload data, send emails, change settings, or modify the database. This ensures data integrity while providing full transparency for national health authorities.

## 3. Core Dashboard Workflow

### Step 1: Upload Data
- Upload your Excel file. If it's encrypted, check the password box.
- **Teaching**: If the system misidentifies a column, correct it and click **"Confirm & Teach"**. The system "learns" your specific hospital's terminology.

### Step 2: Process & Sort
- Click **"Process & Sort"**. This appends data to the encrypted Master Database and generates hospital-specific files in the `data/hospitals` folder.

### Step 3: Review & Email
- Review statistics in the **Overview** tab.
- Use the **Validate & Email** section to send reports to hospitals or validation requests to experts with one click.

---

## 4. SCD Data Discovery (Milestone 4 Superbot)

The Superbot proactively finds missing records in external sources.

### 4.1 Initiating Discovery (Lead)
- Use the **Discovery Initiation** form to send a secure link to a recipient.
- Supports **Bulk Initiation**: You can upload a list of contacts to search multiple sources at once.

### 4.2 Recipient Workflow (The Companion App)
When a recipient opens the link:
1. **Verification**: They must verify their identity via a 6-digit SMS OTP code.
2. **The Companion App**: Recipients are prompted to download a small "Companion Scanner".
3. **Local Scan**: The app scans their computer (Desktop, Documents, Downloads) for SCD-related files.
4. **Security Check**: The app automatically strips dangerous macros and neutralizes malicious formulas before sending any data.
5. **Email Search**: Recipients can optionally authorize a search of their Gmail or Outlook attachments.

### 4.3 Visual Sync Box (Real-Time Monitoring)
On your dashboard, you will see the **Visual Sync Box**:
- **Live Queue**: Records appearing in real-time as the bot finds them.
- **Security Shields**: Green icons confirm that a record has been malware-scanned and "healed" (headers fixed).
- **Atomic Merge**: Once you review the findings, click "Approve" to safely merge them into the Master Database.

---

## 5. Suggested Healing & Intelligence
One of the most powerful features of the system is **Structural Healing**:
**Suggested Healing (Milestone 5)**: To ensure 100% data accuracy, the system has moved from "Automatic Healing" to "Suggested Healing." If the system repairs a header or infers a column meaning, it will flag the record for Lead approval. These records appear in your **Verification Queue**, where you must click "Verify" or "Correct" before they are merged into the Master Database.
- **Misspelled Headers**: If a hospital sends a file with "Ptnt ID" instead of "Patient_ID", the system recognizes it and fixes it automatically based on its learned Knowledge Base.
- **Damaged Files**: If a header row is missing, the system scans the data types to "guess" the column meaning, ensuring no data is lost.

---

## 6. Security & Privacy
- **AES-256 At-Rest Encryption**: Your data is never stored in plain text.
- **PII Masking**: Patient names are masked (e.g., "John Doe" -> "J**n D*e") to protect privacy.
- **Audit Logging**: Every access and change is recorded in a secure log file.

---

## 7. Final Project Status
The SCD Dbase Sorter is now a fully matured system.
- **Sorting Engine**: Fully automated and "teachable".
- **Security**: Enterprise-grade encryption and sanitization.
- **Discovery**: Sequential, OTP-verified search bot with companion app support.
- **Healing**: Advanced structural repair of inconsistent data.

---
**SCD Dbase Sorter Team | 2026**
