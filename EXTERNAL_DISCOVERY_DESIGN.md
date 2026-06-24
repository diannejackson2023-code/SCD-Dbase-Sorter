# Architecture Design: External SCD Data Discovery

## 1. Feature Overview
The "External SCD Data Discovery" feature allows the system to proactively find Sickle Cell Disease (SCD) records that are not yet in the Master Database by searching a recipient's local computer and external email accounts (Gmail/Outlook). This is triggered by the Lead initiating a request via a secure, OTP-verified process.

## 2. High-Level Architecture

### 2.1 Component Diagram
- **Lead Dashboard**: Interface for initiating discovery requests (supports single or **Bulk Initiation** via multi-line/CSV).
- **Discovery Service (Backend)**: Orchestrates OTP, Local Scan coordination, OAuth, and Search Bot logic.
- **Recipient Landing Page**: Secure web interface for the recipient.
- **Local Scan Module**: Browser-based logic utilizing File System Access API.
- **SMS Gateway**: Twilio API for OTP verification.
- **Email APIs**: Google Gmail API and Microsoft Graph API.
- **Parsing Engine**: Extension of existing `mapping.py` with Word (`python-docx`) support.
- **Master DB Index**: A secure, hashed lookup table for comparing records.

## 3. Workflow Details

### 3.1 Initiation (Lead Action)
1. **Bulk Initiation**: The Lead can initiate requests to multiple recipients simultaneously.
    - **Dashboard UI**: A text area for multi-line input (Email, Phone) or a CSV upload.
2. **Secure Link Generation**: For each recipient, a unique, time-limited token is generated and sent via email.
3. **Tracking**: The Lead Dashboard shows the status of each request (Sent, Opened, Local Scan Done, Verified, Email Scan Done, Complete).

### 3.2 Step 1: Local Filesystem Scan (Recipient Action)
Prior to identification, the recipient is guided to scan their local files to find potential SCD records.
1. **Initial Access**: Recipient clicks the link and arrives at the Landing Page. No login required yet.
2. **Local Scan Strategy**:
    - **File System Access API**: Request user permission to access 'Desktop', 'Documents', 'Downloads'.
    - **Execution**: Recursive scan for `.xlsx`, `.xls`, `.docx` files.
3. **UI Feedback**: Real-time progress indicators (e.g., "Scanning Desktop...").
4. **Outcome**: The system identifies the *number* of potential records found locally.

### 3.3 Step 2: Contact Verification & Identity
After the local scan, the recipient must verify their identity to process the records and unlock the Email Scan.
1. **Contact Confirmation**: The recipient confirms or provides their **Email Address and Mobile Number**.
2. **SMS OTP Verification**:
    - The system triggers the **Twilio Verify API** to send a 6-digit code to the mobile number provided.
    - Recipient enters the code to verify their identity.
3. **Unified OTP**: This same verified identity is used to authorize the subsequent Email Scan.

### 3.4 Step 3: Conditional Email Scan (OAuth)
1. **Trigger**: Prompted only after successful identity verification.
2. **OAuth Flow**: Gmail/Outlook read-only access.
3. **Token Management**: In-session access tokens only.

### 3.5 Search Bot & Parsing
1. **Search Bot (Email)**: Searches for keywords like `SCD`, `Sickle Cell`.
2. **Parsing (Local & Email)**:
    - **Excel**: `openpyxl`/`pandas`.
    - **Word**: **`python-docx`**.
3. **Record Extraction**: Extracts `Patient_ID` or anonymized identifier.

### 3.6 Comparison & Reporting
1. **Comparison**: Extracted IDs are hashed (SHA-256) and compared against the **Master DB Index**.
2. **Reporting**: Recipient reviews "Already in DB" vs "Missing from DB".
3. **Push-to-Database**: Finalized files for missing records are sent to the Lead.

## 4. Technical Specifications

### 4.1 SMS OTP Provider: Twilio
- **Service**: **Twilio Verify**. Handles code generation, delivery, and expiry.
- **Workflow**: `Verify Code` endpoint is called browser-side to confirm recipient identity.

### 4.2 Browser File Access
- **API**: `window.showDirectoryPicker()`.
- **Permissions**: User must explicitly click "Allow" for each folder.

### 4.3 Security & Privacy
- **Verification Gating**: Sensitive operations (hashing comparison, email scanning, pushing to DB) are gated behind SMS OTP verification.
- **Volatile Processing**: No data from local or email scans is stored on the Discovery Service disk; it is all handled in-memory and deleted post-session.
- **Anonymization**: Comparison uses SHA-256 hashes to protect Master DB contents.

## 5. UI/UX Flow for Recipient
1. **Local Scan Start**: Immediate access to "Grant Access to Local Folders" upon landing.
2. **Scan Progress**: Dynamic folder-by-folder status.
3. **Identity Capture**: Form to provide/confirm Email and Phone Number.
4. **Verification**: SMS OTP entry to unlock results and Email Scan.
5. **Email Option**: Optional search of Gmail/Outlook.
6. **Review**: Consolidated list of "Discovered Records" with DB status.
7. **Submit**: "Push missing records to Master Database".

## 6. Implementation Roadmap
1. **Phase 1**: Lead Dashboard for **Bulk Initiation** and request tracking.
2. **Phase 2**: Recipient Landing Page with **Local Filesystem Scan** (Step 1).
3. **Phase 3**: Identity confirmation and **Twilio Verify integration** (Step 2).
4. **Phase 4**: OAuth and Email Search Bot (Step 3).
5. **Phase 5**: Sequential workflow state management and Consolidated Reporting.
