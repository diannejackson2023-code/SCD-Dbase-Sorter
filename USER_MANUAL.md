# SCD Dbase Sorter - User Manual

## 1. Overview
Welcome to the **SCD Dbase Sorter**. This tool is designed to help you manage Sickle Cell Disease (SCD) data across multiple hospitals. It automates the process of importing Excel files, sorting data, and proactively discovering new records from local and email sources.

## 2. Getting Started
To start the application, navigate to the project folder and run:
```bash
streamlit run app.py
```
Open your web browser to the address shown (usually `http://localhost:8501`).

## 3. Core Dashboard Workflow

### Step 1: Upload Your Data
- Click the **"Browse files"** button or drag and drop your Excel file.
- The system supports files where headers are in the first or second row.
- **Teaching the System**: In Step 1.5, verify the column mapping. If the bot misidentifies a column, select the correct master heading and click **"Confirm & Teach"**. The bot will remember this alias for all future uploads.

### Step 2: Process & Sort
- Click the **"Process & Sort"** button.
- The system will append records to the encrypted Master Database and generate separate reports for each hospital, sorted by Year.

### Step 3: Review Results
- Use the **Overview** tab for statistics and charts.
- Use the **By Hospital** tab to filter and view specific hospital records.
- Use the **Raw Data** tab to view the full encrypted database.

### Step 4: Send Emails
- Go to the **"Validate & Email"** section.
- Send **Validation Requests** to validators or **Finalized Data** to hospital staff with the correct Excel files attached automatically.

## 4. SCD Data Discovery (The Search Bot)

The system includes a proactive search bot to find missing records.

### 4.1 Initiating Discovery (Lead)
- Go to **"Discovery Initiation"**.
- Enter a recipient's Email and Phone Number (Single or Bulk via CSV).
- The recipient receives a secure link via email.

### 4.2 Recipient Portal Flow
When you or a recipient uses the discovery link:
1.  **Identity Verification**: You must enter your details and verify via a **6-digit SMS code**.
2.  **Computer Search**: The bot requests to scan your local folders (Desktop, Documents, Downloads) for SCD-related Excel or Word files.
3.  **Duplicate Detection**: If a file is already in the database, the bot states: **"✅ File present in Master Database."**
4.  **Email Search**: After the local scan, the bot can search your Gmail or Outlook attachments for SCD results.
5.  **Password Memorization**: If a file is password-protected, the bot prompts for the password once and **memorizes it** for the session.
6.  **Submission**: Any new records found are submitted to the Lead for inclusion in the Master Database.

### 4.3 Reviewing Discoveries
- In the dashboard, use the **"Discovery Review"** section.
- Select a discovery session to see the **"Missing Records"** found by the bot.
- **Authorize with Master Password** to append these new records to the database.
- Check the **"📜 Discovery Log"** for a full history of all files found and their status.

## 5. Security & Privacy
- **AES-256 Encryption**: All data is stored encrypted. Only authorized users with the Master Password can perform write operations.
- **PII Masking**: Patient names are automatically masked (e.g., "John Doe" -> "J**n D*e") to protect privacy while maintaining record identity.
- **Privacy-First Search**: The search bot uses SHA-256 hashes to compare records, meaning it never "sees" your existing patient names during the search process.

---
---
## 6. Project Documentation
The complete project documentation is available in both Markdown (.md) and PDF (.pdf) formats:
- **TECHNICAL_MANUAL.pdf**: Deep dive into the code and architecture.
- **USER_MANUAL.pdf**: This guide in PDF format.
- **CHAT_HISTORY.pdf**: Full project development transcript.
- **DEPLOYMENT_SECURITY.pdf**: Security hardening guide.

**SCD Dbase Sorter Team**
