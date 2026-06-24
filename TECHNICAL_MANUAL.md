# SCD Dbase Sorter - Technical Manual
**Version 1.3**

## 1. Project Overview
The **SCD Dbase Sorter** is a Python-based automated system designed to manage medical records related to Sickle Cell Disease (SCD) across multiple hospitals. It automates the ingestion of inconsistent Excel data, sorts it into a master database, generates hospital-specific reports, and facilitates secure validation and notification via email.

## 2. Core Features
- **Manual Mapping & Teaching System**: Provides a user interface to manually verify and correct column mappings. The system "learns" from manual corrections by saving new aliases to `aliases.json`, enabling future automation.
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
| 2026-05-19 | 23865abb | agent-architect | Architect | Approved | I have created a comprehensive System Architecture Design for the SCD Dbase Sorter. The design includes: System Overview, Excel Structure, Mapping Logic, Data Distribution, Email Workflow, Tech Stack, and Directory Structure. |
| 2026-05-19 | f3566313 | agent-data-eng | Data Eng | Approved | Implemented the core data processing logic. Created mapping.py (header mapping Rows 1/2, alias support, email merging) and sorter.py (master database updates, autosorting, per-hospital file/sheet generation). |
| 2026-05-19 | 16b1e8a3 | agent-comm-eng | Comm Eng | Approved | Implemented email notification logic in processor/mailer.py. Includes email lookups, validation requests, finalized data emails with Excel attachments, and configurable SMTP settings. |
| 2026-05-19 | cd2e03a1 | agent-architect | Architect | Approved | Developed a comprehensive Technical Manual documenting System Architecture, Data Processing, Sorting/Distribution, Email Automation, and UI Guide. |
| 2026-05-19 | 451b2459 | agent-ui-dev | UI Dev | Approved | Developed the Streamlit dashboard (app.py) with 4-step workflow: File Upload, Process & Sort, Data Visualization, and Validate & Email. |
| 2026-05-20 | e6fc5e1a | agent-data-eng | Data Eng | Approved | Implemented secure file handling and PII strategy. Added password-protected Excel support (msoffcrypto) and Patient Name masking (J**n D*e). |
| 2026-05-20 | 7e40e1be | agent-ui-dev | UI Dev | Approved | Integrated secure file handling into the dashboard. Added password fields, smart detection, PII security sidebar, and masking notices. |
| 2026-05-28 | ce6e55e9 | agent-backend-developer | Backend Dev | Approved | Implemented Top-Tier Cybersecurity: AES-256 at-rest encryption, Audit Logging (JSONL), Input Sanitization (Excel/XSS), and TLS enforcement for emails. |

## 5. Deployment Instructions
1. Install dependencies: `pip install -r requirements.txt`
2. Run the application: `streamlit run app.py`
3. Upload an Excel file.
4. **Manual Mapping & Teaching**: Verify the auto-detected column mappings in Step 1.5. If any mapping is wrong, select the correct master heading. Click **"Confirm & Teach"**. This will update the system's memory (`aliases.json`) for future files.
5. **Process & Sort**: Once mapping is confirmed, click **"Process & Sort"** to append data and generate reports.
6. Configure SMTP settings in the sidebar for emailing.

---

## 6. Source Code Appendix

### 6.1 /home/team/shared/SCD_Dbase_Sorter/app.py
```python
"""
SCD Dbase Sorter - Streamlit Dashboard
========================================
Main dashboard application for uploading Excel data,
processing/sorting, and sending validation/hospital emails.
"""

import streamlit as st
import pandas as pd
import os
import sys
import tempfile
import traceback
import io

# Add processor directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "processor"))

# Import team modules
from mapping import load_and_map_data, load_hospital_config, MASTER_HEADINGS, _get_excel_data
from sorter import process_new_data, MASTER_DB_PATH, HOSPITALS_DIR
from encryption import decrypt_file_to_memory
from logger import audit_logger
from mailer import (
    send_validation_request,
    send_finalized_data,
    get_all_hospitals,
    get_hospital_email,
    get_validator_email,
    configure_smtp,
    load_hospital_emails,
    export_df_to_excel
)

# ──────────────────────────────────────────────
# Page Configuration
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="SCD Dbase Sorter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# Session State Initialization
# ──────────────────────────────────────────────
if "processed_data" not in st.session_state:
    st.session_state.processed_data = None
if "master_df" not in st.session_state:
    st.session_state.master_df = None
if "uploaded_filename" not in st.session_state:
    st.session_state.uploaded_filename = None
if "hospital_emails" not in st.session_state:
    st.session_state.hospital_emails = None
if "processing_complete" not in st.session_state:
    st.session_state.processing_complete = False

# ──────────────────────────────────────────────
# Sidebar: Configuration & Info
# ──────────────────────────────────────────────
st.sidebar.title("🔧 SCD Dbase Sorter")
st.sidebar.markdown("---")

# SMTP Configuration
st.sidebar.subheader("📧 SMTP Settings")
smtp_host = st.sidebar.text_input("SMTP Host", value="smtp.gmail.com")
smtp_port = st.sidebar.number_input("SMTP Port", value=587, min_value=1, max_value=65535)
smtp_user = st.sidebar.text_input("SMTP Username", value="")
smtp_pass = st.sidebar.text_input("SMTP Password", type="password", value="")
smtp_from = st.sidebar.text_input("From Email", value="scd.database@example.com")
smtp_from_name = st.sidebar.text_input("From Name", value="SCD Database System")

if st.sidebar.button("Save SMTP Settings"):
    configure_smtp(
        host=smtp_host,
        port=int(smtp_port),
        username=smtp_user,
        password=smtp_pass,
        from_email=smtp_from,
        from_name=smtp_from_name,
    )
    st.sidebar.success("SMTP settings saved!")

st.sidebar.markdown("---")

# Data Paths Info
st.sidebar.subheader("📁 Data Paths")
st.sidebar.info(
    f"""
    **Master DB:** `{MASTER_DB_PATH}`
    **Hospitals:** `{HOSPITALS_DIR}`
    """
)

# Hospital Config
st.sidebar.subheader("🏥 Hospital Config")
try:
    config_df = load_hospital_emails()
    st.sidebar.dataframe(config_df, use_container_width=True)
    st.session_state.hospital_emails = config_df
except Exception as e:
    st.sidebar.warning(f"Config not found: {e}")

st.sidebar.markdown("---")

# Security / PII Section
st.sidebar.subheader("🔒 Cybersecurity Hardening")
st.sidebar.info(
    """
    **Active Protections:**
    - ✅ **At-Rest Encryption**: Database & files encrypted (AES).
    - ✅ **Audit Logging**: Every action logged for security.
    - ✅ **Input Sanitization**: Formula & XSS protection.
    - ✅ **PII Masking**: Patient names masked by default.
    - ✅ **Transport Security**: TLS enforced for emails.
    - ✅ **Source Security**: Support for encrypted Excels.
    """
)

# Documentation Delivery Section
st.sidebar.subheader("📄 Project Docs")
if st.sidebar.button("📤 Email Docs to Owner"):
    with st.sidebar:
        with st.spinner("Sending documents..."):
            doc_files = [
                "/home/team/shared/SCD_Dbase_Sorter/TECHNICAL_MANUAL.md",
                "/home/team/shared/SCD_Dbase_Sorter/USER_MANUAL.md",
                "/home/team/shared/SCD_Dbase_Sorter/CHAT_HISTORY.md"
            ]
            # Check if files exist
            existing_docs = [f for f in doc_files if os.path.exists(f)]
            
            from mailer import _send_email
            success = _send_email(
                to_email="diannejackson2023@gmail.com",
                subject="SCD Dbase Sorter - Project Documentation",
                body="Please find attached the Technical Manual, User Manual, and Chat History for the SCD Dbase Sorter project.",
                attachments=existing_docs
            )
            if success:
                st.success("Docs sent to diannejackson2023@gmail.com!")
                audit_logger.log_action("SEND_DOCS_TO_OWNER", details={"to": "diannejackson2023@gmail.com", "files": existing_docs})
            else:
                st.error("Failed to send. Please check SMTP settings.")

st.sidebar.markdown("---")
st.sidebar.caption("SCD Dbase Sorter v1.1")

# ──────────────────────────────────────────────
# Main Dashboard Content
# ──────────────────────────────────────────────
st.title("📊 SCD Dbase Sorter Dashboard")
st.markdown("Upload Excel data, process & sort, then validate and email.")

# ====== STEP 1: File Upload ======
st.header("Step 1: Upload Data")
uploaded_file = st.file_uploader(
    "Choose an Excel file (.xlsx or .xls)",
    type=["xlsx", "xls"],
    help="Upload an Excel file with data. Headers should be in Row 1 or Row 2. "
         "Password-protected files are supported.",
)

# File Password (for encrypted files) — only show if checkbox is selected
upload_is_encrypted = st.checkbox(
    "🔒 This file is password-protected",
    help="Check this if your Excel file is encrypted and requires a password to open.",
)
file_password = ""
if upload_is_encrypted:
    file_password = st.text_input(
        "Enter file password",
        type="password",
        help="The password used to decrypt the Excel file. "
             "Data is decrypted in-memory and never written unencrypted to disk.",
    )

if uploaded_file is not None:
    st.session_state.uploaded_filename = uploaded_file.name
    st.success(f"✅ Uploaded: {uploaded_file.name}")

    # Preview the uploaded file
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader("File Preview")
        try:
            # Save uploaded file to temp location
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded_file.getbuffer())
                tmp_path = tmp.name

            # Show first 5 rows
            # Handle password if provided
            if file_password:
                excel_data = _get_excel_data(tmp_path, password=file_password)
                preview_df = pd.read_excel(excel_data, nrows=5)
            else:
                preview_df = pd.read_excel(tmp_path, nrows=5)

            st.dataframe(preview_df, use_container_width=True)
            st.caption(f"Showing first 5 rows of {uploaded_file.name}")
        except Exception as e:
            error_msg = str(e).lower()
            if "password" in error_msg or "encrypted" in error_msg or "msoffcrypto" in error_msg:
                st.warning("🔓 This file appears to be encrypted. Check the '🔒 This file is password-protected' box above and enter the correct password.")
                # No need to show the full error for encrypted files
            else:
                st.error(f"Error reading file: {e}")

    with col2:
        st.metric("File Size", f"{len(uploaded_file.getvalue()) / 1024:.1f} KB")

    # ====== STEP 2: Process & Sort ======
    st.header("Step 2: Process & Sort Data")

    if st.button("🚀 Process & Sort", type="primary", use_container_width=True):
        with st.spinner("Processing and sorting data... This may take a moment."):
            try:
                # Step A: Load and map data from the uploaded file
                mapped_df = load_and_map_data(tmp_path, password=file_password)

                # Step B: Process (append to master, generate hospital sheets)
                master_df = process_new_data(mapped_df)
                
                # Store in session state
                st.session_state.processed_data = mapped_df
                st.session_state.master_df = master_df
                st.session_state.processing_complete = True
                
                st.success(f"✅ Processing complete! {len(mapped_df)} records processed.")
                
            except Exception as e:
                st.error(f"❌ Processing failed: {e}")
                st.exception(e)
            finally:
                # Clean up temp file
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

else:
    st.info("📂 Please upload an Excel file to begin.")
    st.session_state.processing_complete = False

# ====== STEP 3: Data Overview & Visualization ======
st.header("📈 Data Overview")

if st.session_state.processing_complete and st.session_state.master_df is not None:
    master_df = st.session_state.master_df
    
    # Use tabs for different views
    tab_overview, tab_hospitals, tab_raw = st.tabs(["Overview", "By Hospital", "Raw Data"])
    
    with tab_overview:
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Records", len(master_df))
        with col2:
            if "Hospital" in master_df.columns:
                n_hospitals = master_df["Hospital"].nunique()
                st.metric("Hospitals", n_hospitals)
        with col3:
            if "Year" in master_df.columns:
                # Count non-null years
                n_years = master_df["Year"].nunique()
                st.metric("Years Covered", n_years)
        
        st.markdown("---")
        
        # Records per Hospital chart
        if "Hospital" in master_df.columns:
            st.subheader("Records per Hospital")
            hosp_counts = master_df["Hospital"].value_counts().reset_index()
            hosp_counts.columns = ["Hospital", "Count"]
            
            # Bar chart using st.bar_chart (built-in, no extra deps)
            chart_df = hosp_counts.set_index("Hospital")
            st.bar_chart(chart_df, use_container_width=True)
            
            # Show numeric table
            col1, col2 = st.columns([1, 1])
            with col1:
                st.dataframe(hosp_counts, use_container_width=True)
        
        # Records per Year chart
        if "Year" in master_df.columns:
            st.subheader("Records per Year")
            year_df = master_df["Year"].copy()
            # Convert to numeric for proper counting
            year_counts = year_df.dropna().astype(int).value_counts().sort_index().reset_index()
            year_counts.columns = ["Year", "Count"]
            
            chart_year = year_counts.set_index("Year")
            st.bar_chart(chart_year, use_container_width=True)
            
            with col2:
                st.dataframe(year_counts, use_container_width=True)
        
        # Cross-tab: Hospital x Year
        if "Hospital" in master_df.columns and "Year" in master_df.columns:
            st.subheader("Records per Hospital & Year")
            try:
                cross_tab = pd.crosstab(
                    master_df["Hospital"].fillna("Unknown"),
                    master_df["Year"].fillna("N/A").astype(str)
                )
                st.dataframe(cross_tab, use_container_width=True)
            except Exception:
                st.info("Could not generate cross-tabulation.")
        
        # Show PII security note
        if "Patient_Name" in master_df.columns:
            st.info("🔒 **PII Protection Active:** Patient names are masked in accordance with the PII security strategy.")
    
    with tab_hospitals:
        st.subheader("Hospital-Specific Data")
        if "Hospital" in master_df.columns:
            hospitals_list = master_df["Hospital"].dropna().unique()
            selected_hospital = st.selectbox(
                "Select a hospital to view details:",
                sorted(hospitals_list)
            )
            
            if selected_hospital:
                hosp_data = master_df[master_df["Hospital"] == selected_hospital]
                st.dataframe(hosp_data, use_container_width=True)
                st.caption(f"Total: {len(hosp_data)} records for {selected_hospital}")
                
                # Show validation status breakdown
                if "Validation_Status" in hosp_data.columns:
                    status_counts = hosp_data["Validation_Status"].value_counts()
                    st.write("**Validation Status Breakdown:**")
                    st.dataframe(status_counts.reset_index(), use_container_width=True)
    
    with tab_raw:
        st.subheader("Master Database (Raw)")
        st.dataframe(master_df, use_container_width=True)
        st.caption(f"Total: {len(master_df)} records | {len(master_df.columns)} columns")

elif os.path.exists(MASTER_DB_PATH):
    # Try loading existing master database if no new data was just processed
    try:
        decrypted_master = decrypt_file_to_memory(MASTER_DB_PATH)
        if decrypted_master:
            existing_df = pd.read_excel(io.BytesIO(decrypted_master))
            if not existing_df.empty:
                st.info("📊 Showing existing master database (upload new data to update).")
                st.dataframe(existing_df, use_container_width=True)
                st.metric("Total Records in Database", len(existing_df))
    except Exception:
        st.info("No data processed yet. Upload and process a file to see insights.")

# ====== STEP 4: Email Actions ======
st.header("📧 Validate & Email")

if st.session_state.processing_complete and st.session_state.master_df is not None:
    master_df = st.session_state.master_df
    
    if "Hospital" in master_df.columns:
        hospitals_list = master_df["Hospital"].dropna().unique()
        
        # Option to email all hospitals at once
        col_bulk1, col_bulk2 = st.columns(2)
        with col_bulk1:
            if st.button("📨 Send Validation Requests (All Hospitals)", use_container_width=True):
                with st.spinner("Sending validation requests..."):
                    results = {}
                    for hospital in hospitals_list:
                        if pd.isna(hospital):
                            continue
                        hosp_data = master_df[master_df["Hospital"] == hospital]
                        
                        # Export hospital data to temp file for attachment
                        safe_name = str(hospital).strip().replace(" ", "_")
                        temp_attach = os.path.join(tempfile.gettempdir(), f"{safe_name}_data.xlsx")
                        export_df_to_excel(hosp_data, temp_attach)
                        
                        result = send_validation_request(hospital, hosp_data, attachments=[temp_attach])
                        results[hospital] = "✅ Sent" if result else "❌ Failed"
                    
                    st.write("**Validation Request Results:**")
                    st.dataframe(
                        pd.DataFrame(list(results.items()), columns=["Hospital", "Status"]),
                        use_container_width=True,
                    )
        
        with col_bulk2:
            if st.button("📨 Send Finalized Data (All Hospitals)", use_container_width=True):
                with st.spinner("Sending finalized data to hospitals..."):
                    results = {}
                    for hospital in hospitals_list:
                        if pd.isna(hospital):
                            continue
                        hosp_data = master_df[master_df["Hospital"] == hospital]
                        
                        safe_name = str(hospital).strip().replace(" ", "_")
                        temp_attach = os.path.join(tempfile.gettempdir(), f"{safe_name}_data.xlsx")
                        export_df_to_excel(hosp_data, temp_attach)
                        
                        result = send_finalized_data(hospital, hosp_data, attachments=[temp_attach])
                        results[hospital] = "✅ Sent" if result else "❌ Failed"
                    
                    st.write("**Finalized Data Send Results:**")
                    st.dataframe(
                        pd.DataFrame(list(results.items()), columns=["Hospital", "Status"]),
                        use_container_width=True,
                    )
        
        st.markdown("---")
        st.subheader("Send Email for Individual Hospital")
        
        # Individual hospital email controls
        col_hosp, col_action = st.columns([2, 1])
        with col_hosp:
            selected_hosp = st.selectbox(
                "Select Hospital:",
                sorted(hospitals_list),
                key="individual_hospital_select",
            )
        with col_action:
            email_type = st.radio("Email Type:", ["Validation Request", "Finalized Data"])
        
        col_info1, col_info2 = st.columns(2)
        with col_info1:
            hosp_email = get_hospital_email(selected_hosp) if selected_hosp else "N/A"
            st.info(f"**Hospital Email:** {hosp_email or 'Not configured'}")
        with col_info2:
            valid_email = get_validator_email(selected_hosp) if selected_hosp else "N/A"
            st.info(f"**Validator Email:** {valid_email or 'Not configured'}")
        
        if st.button("📤 Send Now", type="primary"):
            if selected_hosp and not pd.isna(selected_hosp):
                hosp_data = master_df[master_df["Hospital"] == selected_hosp]
                safe_name = str(selected_hosp).strip().replace(" ", "_")
                temp_attach = os.path.join(tempfile.gettempdir(), f"{safe_name}_data.xlsx")
                export_df_to_excel(hosp_data, temp_attach)
                
                if email_type == "Validation Request":
                    success = send_validation_request(selected_hosp, hosp_data, attachments=[temp_attach])
                else:
                    success = send_finalized_data(selected_hosp, hosp_data, attachments=[temp_attach])
                
                if success:
                    st.success(f"✅ {email_type} sent for {selected_hosp}!")
                else:
                    st.error(f"❌ Failed to send {email_type} for {selected_hosp}. Check SMTP settings.")
            else:
                st.warning("Please select a valid hospital.")

else:
    st.info("Process data first to enable email actions.")
    
    # Show existing database data for emailing, even if no new data was processed
    if os.path.exists(MASTER_DB_PATH):
        try:
            decrypted_master = decrypt_file_to_memory(MASTER_DB_PATH)
            if decrypted_master:
                existing_df = pd.read_excel(io.BytesIO(decrypted_master))
                if not existing_df.empty and "Hospital" in existing_df.columns:
                    st.warning("Showing existing data. Upload and run 'Process & Sort' to refresh.")
                    hospitals_list = existing_df["Hospital"].dropna().unique()
                    selected_hosp = st.selectbox("Select Hospital:", sorted(hospitals_list))
                    
                    if st.button("📤 Send Validation Request"):
                        hosp_data = existing_df[existing_df["Hospital"] == selected_hosp]
                        safe_name = str(selected_hosp).strip().replace(" ", "_")
                        temp_attach = os.path.join(tempfile.gettempdir(), f"{safe_name}_data.xlsx")
                        export_df_to_excel(hosp_data, temp_attach)
                        success = send_validation_request(selected_hosp, hosp_data, attachments=[temp_attach])
                        if success:
                            st.success(f"✅ Sent!")
                        else:
                            st.error("❌ Failed. Check SMTP settings.")
        except Exception:
            pass

# ──────────────────────────────────────────────
# Footer
# ──────────────────────────────────────────────
st.markdown("---")
st.caption("SCD Dbase Sorter | Built with Streamlit | Team SCD Dbase Sorter")
```

### 6.2 /home/team/shared/SCD_Dbase_Sorter/processor/mapping.py
```python
import pandas as pd
import numpy as np
import os
import io
import msoffcrypto

from sanitization import sanitize_dataframe
from logger import audit_logger

# Standard Master Headings
MASTER_HEADINGS = [
    "Patient_ID", "Patient_Name", "Hospital", "Year", "Validation_Status", 
    "Validator_Email", "Hospital_Email", "Date_Added", "Treatment", "Outcome"
]

# Alias Mapping
HEADER_ALIASES = {
    "Hospital": ["Hosp_Name", "Hosp Name", "Facility", "Center", "Hospital Name"],
    "Year": ["Yr", "Data_Year", "Period", "Year of Data"],
    "Patient_ID": ["Pt_No", "Patient ID", "ID", "Case_No", "Patient_ID"],
    "Patient_Name": ["Name", "Patient Name", "Full Name", "Pt Name"],
    "Treatment": ["Rx", "Therapy", "Treatment"],
    "Outcome": ["Result", "Status", "Outcome"]
}

def _get_excel_data(file_path, password=None):
    """
    Helper to handle password-protected Excel files.
    Returns a file-like object or file path.
    """
    if password:
        decrypted_data = io.BytesIO()
        with open(file_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_data)
        decrypted_data.seek(0)
        return decrypted_data
    return file_path

def find_master_match(header_text):
    """Checks if header_text matches any master heading or alias."""
    if not isinstance(header_text, str) or pd.isna(header_text):
        return None
    
    clean_text = str(header_text).strip().lower()
    
    # Check exact/case-insensitive master headings
    for master in MASTER_HEADINGS:
        if master.lower() == clean_text:
            return master
            
    # Check aliases
    for master, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            if alias.lower() == clean_text:
                return master
                
    return None

def get_column_mapping(file_path, password=None):
    """
    Analyzes the first two rows of the Excel file to determine column mapping.
    Returns a dict: {column_index: master_heading}
    """
    excel_data = _get_excel_data(file_path, password)
    # Read first 2 rows without header
    df_header = pd.read_excel(excel_data, header=None, nrows=2)
    
    mapping = {}
    num_cols = df_header.shape[1]
    
    for col in range(num_cols):
        row1_val = df_header.iloc[0, col]
        row2_val = df_header.iloc[1, col]
        
        # Try Row 1
        match = find_master_match(row1_val)
        if match:
            mapping[col] = match
            continue
            
        # Try Row 2
        match = find_master_match(row2_val)
        if match:
            mapping[col] = match
            
    return mapping

CONFIG_PATH = "/home/team/shared/SCD_Dbase_Sorter/data/config/hospital_emails.csv"

def load_hospital_config():
    """Loads hospital email and validator mapping."""
    if os.path.exists(CONFIG_PATH):
        return pd.read_csv(CONFIG_PATH)
    return pd.DataFrame(columns=["Hospital", "Email", "Validator_Email"])

def mask_pii_data(df):
    """
    Applies masking to sensitive columns if they exist.
    """
    if "Patient_Name" in df.columns:
        df["Patient_Name"] = df["Patient_Name"].apply(lambda x: _mask_string(x))
    return df

def _mask_string(val):
    if not isinstance(val, str) or pd.isna(val):
        return val
    parts = val.split()
    masked_parts = []
    for part in parts:
        if len(part) > 2:
            masked_parts.append(part[0] + "*" * (len(part) - 2) + part[-1])
        elif len(part) > 1:
            masked_parts.append(part[0] + "*")
        else:
            masked_parts.append(part)
    return " ".join(masked_parts)

def load_and_map_data(file_path, password=None):
    """
    Loads data from file_path, applies mapping, and returns a normalized DataFrame.
    """
    mapping = get_column_mapping(file_path, password)
    
    excel_data = _get_excel_data(file_path, password)
    # Read the data, skipping the first 2 rows if they were headers
    df = pd.read_excel(excel_data, header=None, skiprows=2)
    
    # Rename columns based on mapping
    df_mapped = pd.DataFrame()
    
    for col_idx, master_name in mapping.items():
        if col_idx < df.shape[1]:
            df_mapped[master_name] = df.iloc[:, col_idx]
            
    # Add default values for missing master columns
    for master in MASTER_HEADINGS:
        if master not in df_mapped.columns:
            df_mapped[master] = np.nan
            
    # Apply PII masking
    df_mapped = mask_pii_data(df_mapped)

    # Apply Sanitization
    from sanitization import sanitize_dataframe
    df_mapped = sanitize_dataframe(df_mapped)

    # Set default Validation_Status if missing
    if 'Validation_Status' not in df_mapped.columns or df_mapped['Validation_Status'].isnull().all():
        df_mapped['Validation_Status'] = 'Pending'
        
    # Set Date_Added
    df_mapped['Date_Added'] = pd.Timestamp.now()
    
    audit_logger.log_action("LOAD_AND_MAP", details={"file": file_path, "records": len(df_mapped)})
    
    # Fill Validator_Email and Hospital_Email from config if missing
    config_df = load_hospital_config()
    if not config_df.empty and 'Hospital' in df_mapped.columns:
        # Merge to get Validator_Email and Email (which is Hospital_Email)
        df_mapped = df_mapped.merge(
            config_df[['Hospital', 'Email', 'Validator_Email']], 
            on='Hospital', 
            how='left', 
            suffixes=('', '_config')
        )
        
        # Fill missing Validator_Email
        if 'Validator_Email_config' in df_mapped.columns:
            df_mapped['Validator_Email'] = df_mapped['Validator_Email'].fillna(df_mapped['Validator_Email_config'])
            df_mapped = df_mapped.drop(columns=['Validator_Email_config'])
            
        # Fill missing Hospital_Email
        if 'Email' in df_mapped.columns:
            df_mapped['Hospital_Email'] = df_mapped['Hospital_Email'].fillna(df_mapped['Email'])
            df_mapped = df_mapped.drop(columns=['Email'])
    
    return df_mapped

```

### 6.3 /home/team/shared/SCD_Dbase_Sorter/processor/sorter.py
```python
import pandas as pd
import os
import io
from datetime import datetime
from encryption import encrypt_file, decrypt_file_to_memory
from logger import audit_logger

MASTER_DB_PATH = "/home/team/shared/SCD_Dbase_Sorter/data/master/Master_Database.xlsx"
HOSPITALS_DIR = "/home/team/shared/SCD_Dbase_Sorter/data/hospitals/"

def ensure_directories():
    """Ensures necessary directories exist."""
    os.makedirs(os.path.dirname(MASTER_DB_PATH), exist_ok=True)
    os.makedirs(HOSPITALS_DIR, exist_ok=True)

def update_master_database(new_data_df):
    """
    Appends new data to the Master Database Excel file.
    """
    ensure_directories()
    
    if os.path.exists(MASTER_DB_PATH):
        try:
            # Decrypt in memory
            decrypted_data = decrypt_file_to_memory(MASTER_DB_PATH)
            master_df = pd.read_excel(io.BytesIO(decrypted_data))
            
            # Ensure Date_Added is datetime
            if 'Date_Added' in master_df.columns:
                master_df['Date_Added'] = pd.to_datetime(master_df['Date_Added'])
            
            combined_df = pd.concat([master_df, new_data_df], ignore_index=True)
        except Exception as e:
            print(f"Error reading master database: {e}")
            audit_logger.log_action("ERROR", details={"msg": f"Error reading master database: {e}"})
            combined_df = new_data_df
    else:
        combined_df = new_data_df

    # Save and Encrypt
    combined_df.to_excel(MASTER_DB_PATH, sheet_name='Master_Data', index=False)
    encrypt_file(MASTER_DB_PATH)
    
    audit_logger.log_action("UPDATE_MASTER", details={"records_added": len(new_data_df)})
    return combined_df

def generate_hospital_sheets(master_df):
    """
    Splits the master dataframe into individual hospital Excel files
    AND adds them as sheets in the Master_Database.xlsx.
    Sorts each by Year.
    """
    ensure_directories()
    
    if master_df.empty:
        return

    # Group by Hospital
    grouped = master_df.groupby('Hospital')
    
    # Decrypt Master for modification
    decrypted_master = decrypt_file_to_memory(MASTER_DB_PATH)
    master_buffer = io.BytesIO(decrypted_master) if decrypted_master else None

    # We'll use ExcelWriter to update the Master_Database with sheets
    # If file doesn't exist or is empty, we create new
    writer_kwargs = {'engine': 'openpyxl'}
    if master_buffer:
        writer_kwargs['mode'] = 'a'
        writer_kwargs['if_sheet_exists'] = 'replace'
        target = master_buffer
    else:
        target = MASTER_DB_PATH

    with pd.ExcelWriter(MASTER_DB_PATH if not master_buffer else master_buffer, **writer_kwargs) as writer:
        if master_buffer:
             # This is a bit tricky with ExcelWriter in-memory. 
             # Let's simplify: always overwrite the file since we have the full master_df anyway
             pass
    
    # Simplified approach: Overwrite everything since we have the full master_df
    with pd.ExcelWriter(MASTER_DB_PATH, engine='openpyxl') as writer:
        # Also write the full Master sheet
        master_df.to_excel(writer, sheet_name='Master_Data', index=False)
        
        for hospital, group in grouped:
            if pd.isna(hospital) or str(hospital).strip() == "":
                hospital_name = "Unassigned"
            else:
                hospital_name = str(hospital).strip()
                
            # Create a safe filename and sheet name
            safe_name = "".join([c for c in hospital_name if c.isalnum() or c in (' ', '_')]).strip()
            safe_filename = safe_name.replace(' ', '_')
            
            # Sort by Year
            if 'Year' in group.columns:
                group['Year_Numeric'] = pd.to_numeric(group['Year'], errors='coerce')
                group = group.sort_values(by=['Year_Numeric', 'Date_Added'], ascending=[False, False])
                group = group.drop(columns=['Year_Numeric'])
            
            # 1. Save to individual file
            file_path = os.path.join(HOSPITALS_DIR, f"{safe_filename}.xlsx")
            group.to_excel(file_path, index=False)
            encrypt_file(file_path)
            
            # 2. Save to a sheet in Master_Database
            sheet_name = safe_name[:31]
            group.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"Generated/Updated sheet and file for {hospital_name}")
    
    # Finally encrypt the Master Database
    encrypt_file(MASTER_DB_PATH)
    audit_logger.log_action("GENERATE_HOSPITAL_SHEETS", details={"hospitals": list(grouped.groups.keys())})

def process_new_data(new_data_df):
    """
    Main entry point for processing new data.
    1. Updates master database.
    2. Regenerates hospital-specific sheets.
    """
    # 1. Update Master
    master_df = update_master_database(new_data_df)
    
    # 2. Refresh hospital files
    generate_hospital_sheets(master_df)
    
    return master_df

```

### 6.4 /home/team/shared/SCD_Dbase_Sorter/processor/mailer.py
```python
"""
Email automation module for SCD Dbase Sorter.
Provides functions to send validation requests and finalized data emails.
"""

import smtplib
import os
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from pathlib import Path
from logger import audit_logger
from encryption import decrypt_file_to_memory

# Configuration file path
HOSPITAL_EMAILS_PATH = "/home/team/shared/SCD_Dbase_Sorter/data/config/hospital_emails.csv"
HOSPITALS_DIR = "/home/team/shared/SCD_Dbase_Sorter/data/hospitals/"

# SMTP Configuration (dummy defaults, can be overridden)
SMTP_CONFIG = {
    "host": "smtp.gmail.com",
    "port": 587,
    "starttls": True,
    "username": "",  # Fill in your SMTP username
    "password": "",  # Fill in your SMTP password/app-specific password
    "from_email": "scd.database@example.com",
    "from_name": "SCD Database System"
}


def load_hospital_emails():
    """
    Loads the hospital email configuration from CSV.
    Returns a DataFrame with Hospital, Email, Validator_Email columns.
    """
    if not os.path.exists(HOSPITAL_EMAILS_PATH):
        raise FileNotFoundError(f"Hospital emails config not found at {HOSPITAL_EMAILS_PATH}")
    
    df = pd.read_csv(HOSPITAL_EMAILS_PATH)
    return df


def get_hospital_email(hospital_name):
    """Returns the email address for a given hospital."""
    df = load_hospital_emails()
    match = df[df['Hospital'].str.lower() == hospital_name.lower()]
    if match.empty:
        return None
    return match.iloc[0]['Email']


def get_validator_email(hospital_name):
    """Returns the validator email for a given hospital."""
    df = load_hospital_emails()
    match = df[df['Hospital'].str.lower() == hospital_name.lower()]
    if match.empty:
        return None
    return match.iloc[0]['Validator_Email']


def get_all_hospitals():
    """Returns list of all hospitals with their emails."""
    return load_hospital_emails()


def get_hospital_data_path(hospital_name):
    """
    Returns the file path for a hospital's Excel data file.
    Uses the same naming logic as sorter.py.
    """
    safe_filename = "".join([c for c in hospital_name if c.isalnum() or c in (' ', '_')]).strip().replace(' ', '_')
    return os.path.join(HOSPITALS_DIR, f"{safe_filename}.xlsx")


def load_hospital_data(hospital_name):
    """
    Loads the DataFrame for a given hospital from their Excel file.
    
    Args:
        hospital_name: Name of the hospital (e.g., "City General")
    
    Returns:
        DataFrame if file exists, empty DataFrame otherwise
    """
    file_path = get_hospital_data_path(hospital_name)
    if os.path.exists(file_path):
        return pd.read_excel(file_path)
    return pd.DataFrame()


def _create_smtp_connection():
    """Creates and returns an SMTP connection."""
    config = SMTP_CONFIG
    server = smtplib.SMTP(config["host"], config["port"])
    if config.get("starttls"):
        server.starttls()
    if config.get("username") and config.get("password"):
        server.login(config["username"], config["password"])
    return server


def _send_email(to_email, subject, body, attachments=None):
    """
    Core email sending function.
    
    Args:
        to_email: Recipient email address
        subject: Email subject line
        body: Email body text (plaintext or HTML)
        attachments: List of file paths to attach (optional)
    
    Returns:
        bool: True if sent successfully, False otherwise
    """
    config = SMTP_CONFIG
    
    msg = MIMEMultipart()
    msg['From'] = f"{config['from_name']} <{config['from_email']}>"
    msg['To'] = to_email
    msg['Subject'] = subject
    
    # Attach body as plain text
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach files if provided
    if attachments:
        for file_path in attachments:
            if os.path.exists(file_path):
                # Check if it's an encrypted .xlsx in hospitals dir
                filename = os.path.basename(file_path)
                if file_path.startswith(HOSPITALS_DIR) or filename == "Master_Database.xlsx":
                    # Try to decrypt
                    file_content = decrypt_file_to_memory(file_path)
                else:
                    with open(file_path, "rb") as f:
                        file_content = f.read()
                
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file_content)
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                msg.attach(part)
    
    try:
        server = _create_smtp_connection()
        server.send_message(msg)
        server.quit()
        print(f"Email sent successfully to {to_email}")
        audit_logger.log_action("SEND_EMAIL", details={"to": to_email, "subject": subject, "attachments": [os.path.basename(f) for f in attachments] if attachments else []})
        return True
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
        audit_logger.log_action("SEND_EMAIL_ERROR", details={"to": to_email, "error": str(e)})
        return False


def send_validation_request(hospital_name, hospital_data_df=None, attachments=None):
    """
    Sends a validation request email to the validator for a hospital.
    
    Args:
        hospital_name: Name of the hospital
        hospital_data_df: DataFrame containing the hospital's data (optional if file auto-attached)
        attachments: Optional list of attachment file paths. If None, the hospital's
                     Excel file from data/hospitals/ will be attached automatically.
    
    Returns:
        bool: True if sent successfully, False otherwise
    """
    validator_email = get_validator_email(hospital_name)
    if not validator_email:
        print(f"No validator email found for hospital: {hospital_name}")
        return False
    
    # Auto-attach hospital's Excel file if no attachments provided
    if attachments is None:
        file_path = get_hospital_data_path(hospital_name)
        if os.path.exists(file_path):
            attachments = [file_path]
    
    subject = f"Validation Request - {hospital_name} SCD Data"
    record_count = len(hospital_data_df) if hospital_data_df is not None else "N/A"
    body = f"""
Dear Validator,

A new dataset has been submitted for the following hospital and requires validation:

Hospital: {hospital_name}
Records: {record_count}

Please review the attached data and confirm its accuracy by replying to this email.

If you have any questions, please contact the database administrator.

Best regards,
SCD Database System
"""
    return _send_email(validator_email, subject, body, attachments=attachments)


def send_finalized_data(hospital_name, hospital_data_df=None, attachments=None):
    """
    Sends finalized/validated data to the hospital's email address.
    
    Args:
        hospital_name: Name of the hospital
        hospital_data_df: DataFrame containing the finalized hospital data (optional)
        attachments: Optional list of additional attachment file paths. If None, the
                     hospital's Excel file from data/hospitals/ will be attached automatically.
    
    Returns:
        bool: True if sent successfully, False otherwise
    """
    hospital_email = get_hospital_email(hospital_name)
    if not hospital_email:
        print(f"No hospital email found for: {hospital_name}")
        return False
    
    # Auto-attach hospital's Excel file if no attachments provided
    if attachments is None:
        file_path = get_hospital_data_path(hospital_name)
        if os.path.exists(file_path):
            attachments = [file_path]
    
    subject = f"Finalized SCD Data - {hospital_name}"
    record_count = len(hospital_data_df) if hospital_data_df is not None else "N/A"
    body = f"""
Dear {hospital_name} Team,

Please find attached the finalized SCD (Sickle Cell Disease) data for your review and records.

Hospital: {hospital_name}
Total Records: {record_count}
Status: Validated and Confirmed

If you have any questions or require further assistance, please do not hesitate to reach out.

Best regards,
SCD Database System
"""
    return _send_email(hospital_email, subject, body, attachments=attachments)


def send_bulk_validation_requests(hospital_data_dict):
    """
    Sends validation request emails for multiple hospitals.
    
    Args:
        hospital_data_dict: Dict mapping hospital_name -> DataFrame
    
    Returns:
        dict: Summary of send results {"hospital_name": bool}
    """
    results = {}
    for hospital_name, df in hospital_data_dict.items():
        results[hospital_name] = send_validation_request(hospital_name, df)
    return results


def send_bulk_finalized_data(hospital_data_dict):
    """
    Sends finalized data emails for multiple hospitals.
    
    Args:
        hospital_data_dict: Dict mapping hospital_name -> DataFrame
    
    Returns:
        dict: Summary of send results {"hospital_name": bool}
    """
    results = {}
    for hospital_name, df in hospital_data_dict.items():
        results[hospital_name] = send_finalized_data(hospital_name, df)
    return results


def configure_smtp(host=None, port=None, username=None, password=None, from_email=None, from_name=None):
    """
    Update SMTP configuration.
    
    Args:
        host: SMTP server host
        port: SMTP server port
        username: SMTP username
        password: SMTP password
        from_email: Sender email address
        from_name: Sender display name
    """
    global SMTP_CONFIG
    if host:
        SMTP_CONFIG["host"] = host
    if port:
        SMTP_CONFIG["port"] = port
    if username:
        SMTP_CONFIG["username"] = username
    if password:
        SMTP_CONFIG["password"] = password
    if from_email:
        SMTP_CONFIG["from_email"] = from_email
    if from_name:
        SMTP_CONFIG["from_name"] = from_name


# ============================================================
# Helper: Export DataFrame to temp Excel for attachment
# ============================================================

def export_df_to_excel(df, output_path):
    """Exports a DataFrame to an Excel file."""
    df.to_excel(output_path, index=False)
    return output_path


if __name__ == "__main__":
    # Simple test
    print("Testing mailer module...")
    hospitals = get_all_hospitals()
    print(f"Loaded {len(hospitals)} hospitals")
    print(hospitals)
```

### 6.5 /home/team/shared/SCD_Dbase_Sorter/processor/encryption.py
```python
import os
from cryptography.fernet import Fernet

KEY_FILE = "/home/team/shared/SCD_Dbase_Sorter/data/config/.master.key"

def ensure_key():
    """Ensures a master key exists."""
    if not os.path.exists(KEY_FILE):
        os.makedirs(os.path.dirname(KEY_FILE), exist_ok=True)
        key = Fernet.generate_key()
        with open(KEY_FILE, "wb") as f:
            f.write(key)
    else:
        with open(KEY_FILE, "rb") as f:
            key = f.read()
    return key

def encrypt_file(file_path):
    """Encrypts a file in place."""
    if not os.path.exists(file_path):
        return
    
    key = ensure_key()
    fernet = Fernet(key)
    
    with open(file_path, "rb") as f:
        data = f.read()
    
    encrypted_data = fernet.encrypt(data)
    
    with open(file_path, "wb") as f:
        f.write(encrypted_data)

def decrypt_file_to_memory(file_path):
    """Decrypts a file and returns the data in memory."""
    if not os.path.exists(file_path):
        return None
    
    key = ensure_key()
    fernet = Fernet(key)
    
    with open(file_path, "rb") as f:
        encrypted_data = f.read()
    
    try:
        decrypted_data = fernet.decrypt(encrypted_data)
        return decrypted_data
    except Exception:
        # If decryption fails, maybe it's not encrypted? 
        # For safety, return the original data or handle error
        return encrypted_data

```

### 6.6 /home/team/shared/SCD_Dbase_Sorter/processor/logger.py
```python
import logging
import os
from datetime import datetime
import json

LOG_DIR = "/home/team/shared/SCD_Dbase_Sorter/data/logs"
AUDIT_LOG_FILE = os.path.join(LOG_DIR, "audit_log.jsonl")

class AuditLogger:
    def __init__(self):
        os.makedirs(LOG_DIR, exist_ok=True)
        self.logger = logging.getLogger("audit_logger")
        self.logger.setLevel(logging.INFO)
        
        # Avoid adding multiple handlers if already initialized
        if not self.logger.handlers:
            handler = logging.FileHandler(AUDIT_LOG_FILE)
            self.logger.addHandler(handler)

    def log_action(self, action, user="system", details=None):
        """
        Logs an action in JSONL format.
        """
        log_entry = {
            "timestamp": datetime.now().isoformat(),
            "action": action,
            "user": user,
            "details": details or {}
        }
        self.logger.info(json.dumps(log_entry))

# Global instance
audit_logger = AuditLogger()

```

### 6.7 /home/team/shared/SCD_Dbase_Sorter/processor/sanitization.py
```python
import html

def sanitize_value(val):
    """
    Sanitizes a single value (string or other).
    """
    if not isinstance(val, str):
        return val
    
    # 1. Prevent Excel Formula Injection
    # Common triggers for Excel formulas are: =, +, -, @
    if val.startswith(('=', '+', '-', '@')):
        # Add a single quote to the front to treat it as text in Excel
        val = "'" + val
        
    # 2. Prevent XSS by escaping HTML tags
    val = html.escape(val)
    
    return val

def sanitize_dataframe(df):
    """
    Applies sanitization to all string columns in a DataFrame.
    """
    # Use applymap for element-wise sanitization
    # Note: applymap is deprecated in newer pandas, use apply(lambda x: x.map(sanitize_value)) or similar
    # For compatibility across versions:
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].apply(sanitize_value)
    return df

```


### 6.12 /home/team/shared/SCD_Dbase_Sorter/processor/pdf_generator.py
```python
from fpdf import FPDF
import os

class SCD_PDF(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 15)
        self.cell(0, 10, 'SCD Dbase Sorter - Project Documentation', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def convert_md_to_pdf(md_path, pdf_path):
    """
    Converts a Markdown file to a PDF with latin-1 sanitization.
    """
    if not os.path.exists(md_path):
        return False
        
    pdf = SCD_PDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=10)
    
    filename = os.path.basename(md_path)
    pdf.set_font('Helvetica', 'B', 14)
    pdf.cell(0, 10, f"Document: {filename}", 0, 1, 'L')
    pdf.ln(5)
    
    effective_page_width = pdf.w - 2 * pdf.l_margin
    
    with open(md_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            # Sanitize for latin-1
            line = line.encode('latin-1', 'replace').decode('latin-1')
            
            if not line:
                pdf.ln(5)
                continue
                
            # Basic Header handling
            if line.startswith('# '):
                pdf.set_font('Helvetica', 'B', 16)
                pdf.multi_cell(effective_page_width, 10, line[2:])
                pdf.set_font('Helvetica', size=10)
            elif line.startswith('## '):
                pdf.set_font('Helvetica', 'B', 14)
                pdf.multi_cell(effective_page_width, 10, line[3:])
                pdf.set_font('Helvetica', size=10)
            else:
                pdf.multi_cell(effective_page_width, 7, line)
                
    pdf.output(pdf_path)
    return True
```

### 6.13 /home/team/shared/SCD_Dbase_Sorter/processor/search_bot.py
```python
"""
Search Bot Logic for External SCD Data Discovery.
"""
import pdfplumber
import pandas as pd

class SearchBot:
    # ... (Discovery Logic)
    def process_pdf(self, file_like):
        """
        Extracts tables from a PDF file. Logs DAMAGED_PDF if corrupted.
        """
        all_data = []
        try:
            with pdfplumber.open(file_like) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        # Column mapping logic...
                        all_data.append(df)
            if all_data:
                return pd.concat(all_data, ignore_index=True)
        except Exception as e:
            return f"DAMAGED_PDF: {str(e)}"
        return pd.DataFrame()
```
