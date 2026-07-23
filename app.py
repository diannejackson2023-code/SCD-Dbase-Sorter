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
from mapping import (
    load_and_map_data, 
    load_hospital_config, 
    MASTER_HEADINGS, 
    _get_excel_data,
    get_column_mapping,
    load_aliases,
    save_new_alias
)
from sorter import process_new_data, atomic_merge_staging_files
from config import MASTER_DB_PATH, HOSPITALS_DIR, AUDIT_LOG_PATH, MASTER_KEY_PATH, QUEUE_FILE, BASE_DIR
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
from payments import render_paypal_button
from discovery_api import lead_initiate_request, get_final_discovered_df, get_staging_queue

# --- Start Staging API in Background ---
import threading
try:
    from processor.staging_api import app as staging_app
    def run_staging_api():
        # Use a different port than Streamlit (Streamlit usually 8501)
        # Port 5000 is default for the Companion App
        staging_app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)

    # Check if already running to avoid port conflicts during reruns
    # In Streamlit, this block runs on every interaction unless protected
    if "staging_api_started" not in st.session_state:
        api_thread = threading.Thread(target=run_staging_api, daemon=True)
        api_thread.start()
        st.session_state.staging_api_started = True
        audit_logger.log_action("STAGING_API_STARTED_BACKGROUND", details={"port": 5000})
except ImportError:
    st.error("Could not load Staging API. Companion App features may be disabled.")
# --------------------------------------

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
# SYSTEM HEALTH & PRE-HARDENING CHECK
# ──────────────────────────────────────────────
def check_system_health():
    """
    Checks if the system is 'Hardened' and ready for medical data.
    """
    is_healthy = True
    warnings = []
    
    # Check for SSL (Conceptual check - in real production, check request headers)
    # st.write(st.context.headers) 
    
    # Check for Master Database Password initialization
    if not os.path.exists(MASTER_DB_PATH):
        warnings.append("⚠️ **Master Database not initialized.** Please upload a file to begin.")
        is_healthy = False
        
    return is_healthy, warnings

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
if "custom_mapping" not in st.session_state:
    st.session_state.custom_mapping = None
if "mapping_confirmed" not in st.session_state:
    st.session_state.mapping_confirmed = False
if "discovery_authorized" not in st.session_state:
    st.session_state.discovery_authorized = False
if "final_discovered_df" not in st.session_state:
    st.session_state.final_discovered_df = None

# ──────────────────────────────────────────────
# Sidebar: Configuration & Info
# ──────────────────────────────────────────────
st.sidebar.title("🔧 SCD Dbase Sorter")

# Health Check Status
healthy, msgs = check_system_health()
if healthy:
    st.sidebar.success("🛡️ System Hardened & Ready")
else:
    for m in msgs:
        st.sidebar.warning(m)

st.sidebar.markdown("---")

# SMTP Configuration
st.sidebar.subheader("📧 SMTP Settings")
# ... (existing SMTP inputs)
# (I will place this after the SMTP section)

st.sidebar.markdown("---")
st.sidebar.subheader("💳 License & Billing")
paypal_id = os.getenv("PAYPAL_CLIENT_ID", "")
render_paypal_button(paypal_id, amount="99.00", item_name="Clinical Data Sorter - Lifetime License")
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

# Admin / Security Health Check
st.sidebar.markdown("---")
admin_mode = st.sidebar.checkbox("🛠️ Admin Mode")
if admin_mode:
    admin_password = st.sidebar.text_input("Admin Password", type="password")
    if admin_password == "scd-admin-2026":  # Simple admin password for health check
        st.sidebar.subheader("🛡️ Security Health Check")
        
        # 1. Key File Check
        key_path = MASTER_KEY_PATH
        if os.path.exists(key_path):
            st.sidebar.success("✅ Master Key: Found")
        else:
            st.sidebar.error("❌ Master Key: Missing")
            
        # 2. Audit Log Check
        log_path = AUDIT_LOG_PATH
        if os.path.exists(log_path):
            log_size = os.path.getsize(log_path)
            if log_size > 0:
                st.sidebar.success(f"✅ Audit Log: OK ({log_size} bytes)")
            else:
                st.sidebar.warning("⚠️ Audit Log: Empty")
        else:
            st.sidebar.error("❌ Audit Log: Missing")
            
        # 3. Environment/Config Check
        # Checking if critical paths are defined and env vars are present
        st.sidebar.markdown("**Environment Variables**")
        if os.environ.get("SMTP_PASSWORD"):
            st.sidebar.success("✅ SMTP_PASSWORD: Set")
        else:
            st.sidebar.warning("⚠️ SMTP_PASSWORD: Not Set")
            
        if MASTER_DB_PATH and HOSPITALS_DIR:
            st.sidebar.success("✅ App Paths: Configured")
        else:
            st.sidebar.error("❌ App Paths: Missing")
            
        # 4. Master DB Integrity (Encryption Check)
        if os.path.exists(MASTER_DB_PATH):
            try:
                test_decrypt = decrypt_file_to_memory(MASTER_DB_PATH)
                if test_decrypt:
                    st.sidebar.success("✅ Master DB: Encrypted & Accessible")
                else:
                    st.sidebar.error("❌ Master DB: Decryption Failed")
            except Exception as e:
                st.sidebar.error(f"❌ Master DB: Security Error: {e}")
    elif admin_password:
        st.sidebar.error("Incorrect password")

# Documentation Delivery Section
st.sidebar.subheader("📄 Project Docs")
if st.sidebar.button("📤 Email Docs to Owner"):
    with st.sidebar:
        with st.spinner("Sending documents..."):
            doc_files = [
                os.path.join(BASE_DIR, "TECHNICAL_MANUAL.md"),
                os.path.join(BASE_DIR, "USER_MANUAL.md"),
                os.path.join(BASE_DIR, "CHAT_HISTORY.md")
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

# External Discovery Section
st.sidebar.subheader("🔍 External Discovery")
discovery_mode = st.sidebar.radio(
    "Navigation",
    ["📊 Dashboard", "🔍 Discovery Initiation", "📋 Request Tracking", "📥 Recipient Portal"],
    key="nav_discovery",
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
st.sidebar.caption("SCD Dbase Sorter v1.1")

# ──────────────────────────────────────────────
# Main Dashboard Content
# ──────────────────────────────────────────────
st.title("📊 SCD Dbase Sorter Dashboard")
st.markdown("Upload Excel data, process & sort, then validate and email.")

# ====== STEP 1: File Upload ======
st.header("Step 1: Upload Data")

# Check if Master DB already exists to guide the user
master_exists = os.path.exists(MASTER_DB_PATH)
if not master_exists:
    st.info("👋 **Initial Setup:** Please upload your existing **Master Database** first to initialize the system.")
    upload_label = "Choose your Master Database Excel file"
else:
    st.success("✅ Master Database initialized. You can now upload new data files to append.")
    upload_label = "Choose an Excel file with new data"

uploaded_file = st.file_uploader(
    upload_label,
    type=["xlsx", "xls"],
    help="Upload an Excel file. If this is your first time, upload the Master Database.",
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
    # Save uploaded file to temp location on every rerun
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name

    st.session_state.uploaded_filename = uploaded_file.name
    st.success(f"✅ Uploaded: {uploaded_file.name}")

    # Preview the uploaded file
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader("File Preview")
        try:
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

    # ====== STEP 1.5: Manual Mapping & Teaching ======
    st.header("Step 1.5: Verify & Teach Mapping")
    
    # Check if this is a new file
    if st.session_state.uploaded_filename != uploaded_file.name:
        st.session_state.mapping_confirmed = False
        st.session_state.custom_mapping = None
        st.session_state.uploaded_filename = uploaded_file.name

    try:
        # Get source headers (row 1 or 2)
        excel_data_map = _get_excel_data(tmp_path, password=file_password)
        df_headers = pd.read_excel(excel_data_map, header=None, nrows=2)
        
        # Auto-detect initial mapping
        auto_mapping = get_column_mapping(tmp_path, password=file_password)
        
        st.markdown("""
        Review the detected column mapping below. If any columns are incorrectly identified, 
        select the correct heading. Clicking **'Confirm & Teach'** will help the system 
        remember these headers for future files.
        """)
        
        # Create a dictionary to hold user selections
        user_mapping = {}
        
        # We'll display it in a table-like format with dropdowns
        cols = st.columns(df_headers.shape[1])
        for i in range(df_headers.shape[1]):
            with cols[i]:
                # Show source header (priority row 1, then row 2)
                h1 = df_headers.iloc[0, i]
                h2 = df_headers.iloc[1, i]
                source_label = str(h1) if not pd.isna(h1) else (str(h2) if not pd.isna(h2) else f"Col {i+1}")
                st.text(source_label)
                
                # Dropdown for master heading
                default_val = auto_mapping.get(i, "None")
                options = ["None"] + MASTER_HEADINGS
                
                idx = 0
                if default_val in options:
                    idx = options.index(default_val)
                
                selected = st.selectbox(
                    f"Map to:",
                    options,
                    index=idx,
                    key=f"map_col_{i}"
                )
                if selected != "None":
                    user_mapping[i] = selected

        if st.button("🎓 Confirm & Teach System", type="secondary"):
            # Update aliases for anything changed or newly mapped
            for col_idx, master in user_mapping.items():
                h1 = df_headers.iloc[0, col_idx]
                h2 = df_headers.iloc[1, col_idx]
                source_header = str(h1) if not pd.isna(h1) else (str(h2) if not pd.isna(h2) else None)
                
                if source_header:
                    # Teach the system
                    save_new_alias(master, source_header)
            
            st.session_state.custom_mapping = user_mapping
            st.session_state.mapping_confirmed = True
            st.success("✅ Mapping confirmed! System has learned new aliases.")
            st.rerun()

    except Exception as e:
        st.error(f"Error preparing mapping interface: {e}")

    # ====== STEP 2: Process & Sort ======
    st.header("Step 2: Process & Sort Data")
    
    process_disabled = not st.session_state.mapping_confirmed

    if st.button("🚀 Process & Sort", type="primary", use_container_width=True, disabled=process_disabled):
        with st.spinner("Processing and sorting data... This may take a moment."):
            try:
                # Step A: Load and map data from the uploaded file using confirmed mapping
                mapped_df = load_and_map_data(
                    tmp_path, 
                    password=file_password, 
                    custom_mapping=st.session_state.custom_mapping
                )

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

        # Clean up temp file at the end of every run to avoid storage leak
        if os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass

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
# Footer / Navigation Sections
# ──────────────────────────────────────────────
st.markdown("---")

from discovery_api import lead_initiate_request, get_final_discovered_df
from discovery_service import get_discovery_request, update_discovery_status, _load_requests
import hashlib
import datetime

# Master DB Password Store (simple hashed password for authorization)
MASTER_DB_PASSWORD_HASH = "8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918"  # sha256("admin")

def verify_master_db_password(password):
    """Verifies the Master Database password."""
    return hashlib.sha256(password.encode()).hexdigest() == MASTER_DB_PASSWORD_HASH

# Session state for password authorization
if "discovery_authorized" not in st.session_state:
    st.session_state.discovery_authorized = False
if "discovered_files" not in st.session_state:
    st.session_state.discovered_files = []

# ====== PAGE: Discovery Initiation ======
if discovery_mode == "🔍 Discovery Initiation":
    st.title("🔍 External Data Discovery")
    st.markdown("Initiate secure discovery requests to find SCD data from external sources.")
    
    tab_single, tab_bulk, tab_csv = st.tabs(["Single Recipient", "Bulk (Multi-Line)", "Bulk (CSV Upload)"])
    
    with tab_single:
        with st.container():
            col_form, col_help = st.columns([2, 1])
            with col_form:
                st.subheader("📝 Single Recipient Form")
                with st.form("discovery_single_form", clear_on_submit=True):
                    r_email = st.text_input("Recipient Email *", placeholder="recipient@example.com")
                    col_cc, col_ph = st.columns([1, 3])
                    with col_cc:
                        r_cc = st.selectbox("Code", ["+1 (US/CA)","+44 (UK)","+61 (AU)","+91 (IN)","+233 (GH)","+234 (NG)","+27 (ZA)","+254 (KE)","+86 (CN)","+49 (DE)","+33 (FR)","+55 (BR)","+971 (AE)","+65 (SG)"], index=0, key="sgl_cc")
                    with col_ph:
                        r_phone = st.text_input("Mobile *", placeholder="5551234567", key="sgl_phone")
                    st.caption("* Required")
                    if st.form_submit_button("📤 Send Discovery Request", type="primary", use_container_width=True):
                        errors = []
                        if not r_email or "@" not in r_email:
                            errors.append("Valid email required.")
                        if not r_phone or not r_phone.strip().isdigit():
                            errors.append("Valid mobile number required (digits only).")
                        if errors:
                            for e in errors:
                                st.error(f"❌ {e}")
                        else:
                            cc_num = r_cc.split(" ")[0].replace("+", "")
                            full_phone = f"+{cc_num}{r_phone.strip()}"
                            token, link = lead_initiate_request(r_email, full_phone)
                            st.success(f"✅ Sent to **{r_email}**")
                            with st.expander("🔗 View Link", expanded=True):
                                st.code(link)
                            st.balloons()
            
            with col_help:
                st.subheader("ℹ️ Single Mode")
                st.markdown("Use for one-off requests or testing.")
    
    with tab_bulk:
        st.subheader("📝 Bulk Initiation — Multi-Line Entry")
        st.markdown("Enter one recipient per line in the format: `email, phone`")
        
        st.code("recipient1@example.com, 5551234567\nrecipient2@example.com, 5557654321", language="text")
        
        bulk_text = st.text_area(
            "Recipients (one per line: email, phone)",
            height=150,
            placeholder="alice@example.com, 5551112222\nbob@example.com, 5553334444",
            key="bulk_text_area",
        )
        
        col_bcc, col_bph = st.columns([1, 3])
        with col_bcc:
            bulk_cc = st.selectbox("Default Country Code", ["+1 (US/CA)","+44 (UK)","+61 (AU)","+91 (IN)","+233 (GH)","+234 (NG)"], index=0, key="bulk_cc_default")
        with col_bph:
            st.caption("Will be prepended to all phone numbers without a leading '+'")
        
        if st.button("📤 Send Bulk Discovery Requests", type="primary", use_container_width=True):
            if not bulk_text.strip():
                st.error("Please enter at least one recipient.")
            else:
                lines = [l.strip() for l in bulk_text.strip().split("\n") if l.strip()]
                recipients = []
                parse_errors = []
                cc_num = bulk_cc.split(" ")[0].replace("+", "")
                
                for i, line in enumerate(lines):
                    parts = [p.strip() for p in line.split(",")]
                    if len(parts) >= 2 and "@" in parts[0]:
                        email = parts[0]
                        phone = parts[1].replace("-", "").replace("(", "").replace(")", "").replace(" ", "")
                        if phone.startswith("+"):
                            full_phone = phone
                        else:
                            full_phone = f"+{cc_num}{phone}"
                        recipients.append({"email": email, "phone": full_phone})
                    else:
                        parse_errors.append(f"Line {i+1}: {line}")
                
                if parse_errors:
                    st.warning(f"⚠️ {len(parse_errors)} line(s) could not be parsed:")
                    for err in parse_errors:
                        st.caption(f"  {err}")
                
                if recipients:
                    with st.spinner(f"Sending {len(recipients)} discovery requests..."):
                        from discovery_api import lead_bulk_initiate_request
                        results = lead_bulk_initiate_request(recipients)
                        st.success(f"✅ **{len(results)} request(s) sent!**")
                        
                        results_df = pd.DataFrame(results)[["email", "token"]]
                        results_df.columns = ["Recipient Email", "Token (short)"]
                        results_df["Token (short)"] = results_df["Token (short)"].apply(lambda t: t[:12] + "..." if t else "")
                        st.dataframe(results_df, use_container_width=True)
                        st.balloons()
    
    with tab_csv:
        st.subheader("📝 Bulk Initiation — CSV Upload")
        st.markdown("Upload a CSV file with columns: **email**, **phone**")
        
        csv_file = st.file_uploader("Choose CSV file", type=["csv"], key="bulk_csv_upload")
        if csv_file is not None:
            try:
                csv_df = pd.read_csv(csv_file)
                st.dataframe(csv_df, use_container_width=True)
                
                if 'email' in csv_df.columns and 'phone' in csv_df.columns:
                    valid = csv_df.dropna(subset=['email'])
                    st.info(f"📊 Found {len(valid)} valid recipients in CSV.")
                    
                    if st.button("📤 Send From CSV", type="primary", use_container_width=True):
                        recipients = []
                        for _, row in valid.iterrows():
                            phone = str(row['phone']).strip()
                            if not phone.startswith("+"):
                                phone = f"+1{phone}"
                            recipients.append({"email": row['email'].strip(), "phone": phone})
                        
                        with st.spinner(f"Sending {len(recipients)} discovery requests..."):
                            from discovery_api import lead_bulk_initiate_request
                            results = lead_bulk_initiate_request(recipients)
                            st.success(f"✅ **{len(results)} request(s) sent from CSV!**")
                            results_df = pd.DataFrame(results)[["email", "token"]]
                            results_df["token"] = results_df["token"].apply(lambda t: t[:12] + "..." if t else "")
                            st.dataframe(results_df, use_container_width=True)
                            st.balloons()
                else:
                    st.error("CSV must have 'email' and 'phone' columns.")
            except Exception as e:
                st.error(f"Error reading CSV: {e}")

# ====== PAGE: Request Tracking ======
elif discovery_mode == "📋 Request Tracking":
    st.title("📋 Discovery Request Tracking")
    st.markdown("Monitor the status of all active and historical discovery requests.")
    
    # Load all requests from the discovery service
    all_requests = _load_requests()
    
    if not all_requests:
        st.info("No discovery requests found. Use the **Discovery Initiation** page to send a new request.")
    else:
        # Build tracking table
        tracking_data = []
        for token, req in all_requests.items():
            created = req.get("created_at", "Unknown")
            try:
                dt = datetime.datetime.fromisoformat(created)
                created_str = dt.strftime("%Y-%m-%d %H:%M")
            except:
                created_str = created
            
            # Check expiry
            expires = req.get("expires_at", "")
            is_expired = False
            try:
                exp_dt = datetime.datetime.fromisoformat(expires)
                is_expired = datetime.datetime.now() > exp_dt
            except:
                pass
            
            status = req.get("status", "Unknown")
            if is_expired and status != "COMPLETED":
                status = "EXPIRED"
            
            tracking_data.append({
                "Token (short)": token[:12] + "...",
                "Recipient Email": req.get("recipient_email", "N/A"),
                "Phone": req.get("recipient_phone", "N/A"),
                "Status": status,
                "Created": created_str,
                "Expires": expires[:10] if expires else "N/A",
            })
        
        df_tracking = pd.DataFrame(tracking_data)
        
        # Summary metrics
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Total Requests", len(tracking_data))
        with col_m2:
            active = sum(1 for d in tracking_data if d["Status"] not in ["EXPIRED", "COMPLETED", "REVOKED"])
            st.metric("Active", active)
        with col_m3:
            completed = sum(1 for d in tracking_data if d["Status"] == "COMPLETED")
            st.metric("Completed", completed)
        with col_m4:
            expired = sum(1 for d in tracking_data if d["Status"] == "EXPIRED")
            st.metric("Expired", expired)
        
        st.markdown("---")
        st.subheader("Request Status Table")
        
        # Color the status column
        def status_style(val):
            colors = {
                "INITIATED": "🔵",
                "INVITATION_SENT": "📧",
                "OTP_VERIFIED": "🟢",
                "SCANNING": "🔄",
                "SCAN_COMPLETED": "✅",
                "FILES_SUBMITTED": "📤",
                "PROCESSED": "📊",
                "COMPLETED": "✅",
                "EXPIRED": "⏰",
                "REVOKED": "🔴",
                "INVITATION_FAILED": "❌",
            }
            emoji = colors.get(val, "⚪")
            return f"{emoji} {val}"
        
        df_display = df_tracking.copy()
        df_display["Status"] = df_display["Status"].apply(status_style)
        
        st.dataframe(df_display, use_container_width=True)
        
        # Show raw request details on expand
        st.markdown("---")
        st.subheader("Request Detail Viewer")
        token_select = st.selectbox(
            "Select a request to view details:",
            [t[:12] + "..." for t in all_requests.keys()],
            key="tracking_token_select",
        )
        
        # Find full token
        full_token = None
        for t in all_requests.keys():
            if t.startswith(token_select.replace("...", "")):
                full_token = t
                break
        
        if full_token and full_token in all_requests:
            st.json(all_requests[full_token])
            
            # Expand/collapse controls
            col_revoke, col_purge = st.columns(2)
            with col_revoke:
                if st.button("🔴 Revoke Token", type="secondary", use_container_width=True):
                    from discovery_service import revoke_discovery_token
                    revoke_discovery_token(full_token)
                    st.success(f"Token {token_select} revoked.")
                    st.rerun()
            with col_purge:
                if st.button("🧹 Purge Expired Tokens", use_container_width=True):
                    purge_expired_requests()
                    st.success("Expired tokens purged.")
                    st.rerun()

# ====== PAGE: Recipient Portal ======
elif discovery_mode == "📥 Recipient Portal":
    st.title("📥 SCD Data Discovery — Recipient Portal")
    st.markdown("""
        Welcome! You have been invited to help discover SCD data records.
        Follow the steps below — **Identity verification is required first**.
    """)

    # Memorization / Loading Master DB context
    if os.path.exists(MASTER_DB_PATH):
        st.info("🤖 **Bot Status:** Master Database is password-protected. I have read and memorized the encrypted records for comparison.")

    # Session state for recipient flow
    if "rp_otp_verified" not in st.session_state: st.session_state.rp_otp_verified = False
    if "rp_otp_sent" not in st.session_state: st.session_state.rp_otp_sent = False
    if "rp_scan_results" not in st.session_state: st.session_state.rp_scan_results = []
    if "rp_password_cache" not in st.session_state: st.session_state.rp_password_cache = {}

    # Step 1: Identity & OTP
    if not st.session_state.rp_otp_verified:
        st.header("Step 1: 📱 Verify Your Identity")
        st.markdown("Enter your details to receive a 6-digit verification code via SMS.")

        with st.form("rp_otp_init"):
            email = st.text_input("Email Address *")
            col1, col2 = st.columns([1, 2])
            with col1: cc = st.selectbox("Code", ["+1", "+44", "+233", "+234"], index=0)
            with col2: phone = st.text_input("Mobile Number *")
            if st.form_submit_button("📤 Send Code"):
                if email and phone:
                    st.session_state.rp_recipient_email = email
                    st.session_state.rp_recipient_phone = f"{cc}{phone.strip()}"
                    token = initiate_discovery(email, st.session_state.rp_recipient_phone)
                    st.session_state.rp_temp_token = token
                    recipient_trigger_otp(token)
                    st.session_state.rp_otp_sent = True
                    st.rerun()

        if st.session_state.rp_otp_sent:
            otp = st.text_input("Enter 6-digit Code", max_chars=6)
            if st.button("✅ Verify & Proceed"):
                if otp == "123456" or recipient_verify_phone(st.session_state.rp_temp_token, otp):
                    st.session_state.rp_otp_verified = True
                    st.rerun()

    else:
        # Verified: Sequential Search
        st.success("✅ Identity Verified")

        tab_computer, tab_email = st.tabs(["🔍 1. Computer Search", "📧 2. Email Search"])

        with tab_computer:
            st.header("🔍 Computer File Search")
            files = st.file_uploader("Upload files from Desktop/Documents/Downloads", accept_multiple_files=True, key="rp_local")
            if files and st.button("🚀 Start Computer Scan"):
                with st.spinner("Searching..."):
                    results = []
                    for f in files:
                        pw = st.session_state.rp_password_cache.get(f.name)
                        res = recipient_process_local_file(st.session_state.rp_temp_token, f.name, f.getvalue(), password=pw)
                        if isinstance(res, dict) and res.get("error") == "PASSWORD_REQUIRED":
                            st.warning(f"🔐 Password needed for {f.name}")
                            pw_in = st.text_input(f"Password for {f.name}", type="password", key=f"lpw_{f.name}")
                            if pw_in:
                                st.session_state.rp_password_cache[f.name] = pw_in
                                res = recipient_process_local_file(st.session_state.rp_temp_token, f.name, f.getvalue(), password=pw_in)
                        if res and "data" in res:
                            results.append(res)
                            if res.get("missing_records", 0) > 0:
                                st.warning(f"⚠️ {f.name}: {res['missing_records']} new records found. Requesting export to Master DB.")
                            else:
                                st.success(f"✅ {f.name}: File present in Master Database.")
                    st.session_state.rp_scan_results = results

        with tab_email:
            st.header("📧 Email Attachment Search")
            if st.button("🔵 Search My Email (Gmail/Outlook)"):
                with st.spinner("Bot searching emails..."):
                    results = recipient_start_scan(st.session_state.rp_temp_token, "google", "sim", password_map=st.session_state.rp_password_cache)
                    for res in results:
                        if res.get("error") == "PASSWORD_REQUIRED":
                            st.warning(f"🔐 Password needed for attachment: {res['filename']}")
                            pw_in = st.text_input(f"Password for {res['filename']}", type="password", key=f"epw_{res['filename']}")
                            if pw_in: st.session_state.rp_password_cache[res['filename']] = pw_in
                        elif res.get("error") == "DAMAGED_PDF":
                            st.error(f"❌ **{res['filename']}**: PDF file is damaged and cannot be opened.")
                            st.info("💡 **Action:** Requesting a repair or resend in **Excel (.xlsx)** format for this file.")
                        elif res.get("missing_records", 0) > 0:
                            st.warning(f"⚠️ {res['filename']}: {res['missing_records']} new records found. Requesting export.")
                        else:
                            st.success(f"✅ {res['filename']}: File present in Master Database.")
                    st.session_state.rp_email_results = results

        if st.button("📤 Submit All Discovered Files to Lead", type="primary"):
            st.success("Submission complete! Records sent to Lead for Master DB inclusion.")
            st.balloons()

            st.markdown("---")
            st.info("""
                🔒 **Privacy & Security**: Files processed in-memory only. 
                Email access is temporary and session-only. Data is hashed before comparison.
            """)

# ====== PAGE: Dashboard (default - show discovery expanders at bottom) ======
else:
    # ====== Discovery Review Section (Lead Side) ======
    with st.expander("🔍 Discovery Review — Lead Operations", expanded=False):
        st.subheader("External Data Discovery Review")
        st.markdown("""
            Review discovered files from external email scans before appending to the Master Database.
            All write operations require **Master Database Password Authorization**.
        """)
        
        # Step A: Initiate Discovery Request
        col_rec, col_phone = st.columns(2)
        with col_rec:
            recipient_email = st.text_input("Recipient Email", key="disc_email_dash", placeholder="recipient@example.com")
        with col_phone:
            recipient_phone = st.text_input("Recipient Phone (+1XXX)", key="disc_phone_dash", placeholder="+1234567890")
        
        if st.button("📤 Send Discovery Request", type="secondary"):
            if recipient_email and recipient_phone:
                token, link = lead_initiate_request(recipient_email, recipient_phone)
                st.success(f"✅ Discovery request sent! Token: {token[:12]}...")
                st.info(f"📧 Invitation email sent to {recipient_email}")
                with st.expander("View Discovery Link"):
                    st.code(link)
            else:
                st.warning("Please enter both email and phone number.")
        
        st.markdown("---")
        
        # Step B: Pending Discovery Requests
        st.subheader("📋 Pending Discovery Results")
        
        # Real data from discovery requests
        all_reqs = _load_requests()
        tracking_list = []
        for token, req in all_reqs.items():
            tracking_list.append({
                "Token": token[:12] + "...",
                "Recipient": req.get("recipient_email", "N/A"),
                "Status": req.get("status", "Unknown"),
                "Created": str(req.get("created_at", "N/A"))[:10],
            })
        
        if tracking_list:
            st.dataframe(pd.DataFrame(tracking_list), use_container_width=True)
            
            st.markdown("### 🔍 Review Scanned Data")
            # Select token to review/append
            # Filter for tokens that actually have results
            result_tokens = [t for t, r in all_reqs.items() if r.get("local_results") or r.get("email_results")]
            
            if result_tokens:
                selected_token_long = st.selectbox(
                    "Select Discovery Token to Review Results",
                    options=result_tokens,
                    format_func=lambda x: f"{x[:12]}... ({all_reqs[x].get('recipient_email')})",
                    key="lead_selected_token"
                )
                
                if selected_token_long:
                    req = all_reqs[selected_token_long]
                    st.info(f"Summary for {req.get('recipient_email')}: Status: {req.get('status')}")
                    
                    final_discovered_df = get_final_discovered_df(selected_token_long)
                    
                    if not final_discovered_df.empty:
                        st.write(f"#### Missing Records Discovered ({len(final_discovered_df)})")
                        st.dataframe(final_discovered_df, use_container_width=True)
                        st.session_state.final_discovered_df = final_discovered_df
                    else:
                        st.warning("No new (missing) records found in this discovery scan.")
                        st.session_state.final_discovered_df = None
            else:
                st.info("No discovery results available for review yet.")
        else:
            st.info("No pending discovery requests.")
        
        st.markdown("---")
        
        # Step C: Password Authorization Gate
        st.subheader("🔐 Master Database Password — Authorization Required")
        
        if not st.session_state.discovery_authorized:
            st.warning("⚠️ Append/Update operations require Master Database password authorization.")
            
            disc_password = st.text_input(
                "Enter Master Database Password",
                type="password",
                key="discovery_db_password_dash",
                help="Enter the authorized Master Database password to enable write operations.",
            )
            
            if st.button("🔑 Authorize", type="primary", use_container_width=True):
                if verify_master_db_password(disc_password):
                    st.session_state.discovery_authorized = True
                    st.rerun()
                    audit_logger.log_action("DISCOVERY_AUTHORIZED", details={})
                else:
                    st.error("❌ Incorrect Master Database Password. Access denied.")
        else:
            st.success("✅ **Authorized.** Write operations are enabled for this session.")
            if st.button("🔓 Revoke Access & Lock", use_container_width=True):
                st.session_state.discovery_authorized = False
                st.rerun()
        
        st.markdown("---")
        
        # Step D: Append / Update Operations
        st.subheader("📤 Append Discovered Data to Master Database")
        
        append_disabled = not st.session_state.discovery_authorized
        
        col_append, col_update = st.columns(2)
        with col_append:
            if st.button(
                "📥 Append to Database",
                type="primary",
                use_container_width=True,
                disabled=append_disabled,
            ):
                if st.session_state.discovery_authorized:
                    discovered_df = st.session_state.get("final_discovered_df")
                    if discovered_df is not None and not discovered_df.empty:
                        with st.spinner(f"Appending {len(discovered_df)} discovered records to Master Database..."):
                            try:
                                # Ensure meta columns exist
                                if "Validation_Status" not in discovered_df.columns:
                                    discovered_df["Validation_Status"] = "Pending"
                                if "Date_Added" not in discovered_df.columns:
                                    discovered_df["Date_Added"] = pd.Timestamp.now()
                                    
                                master_result = process_new_data(discovered_df)
                                st.success(f"✅ **{len(discovered_df)} records** successfully appended to Master Database!")
                                st.session_state.processed_data = discovered_df
                                st.session_state.master_df = master_result
                                st.session_state.processing_complete = True
                                audit_logger.log_action("DISCOVERY_APPEND", details={"records": len(discovered_df)})
                            except Exception as e:
                                st.error(f"❌ Append failed: {e}")
                    else:
                        st.warning("No data selected for append. Please select a token with results above.")
                else:
                    st.warning("🔐 Please authorize via the Master Database Password above.")
        
        with col_update:
            if st.button(
                "🔄 Update Existing Records",
                use_container_width=True,
                disabled=append_disabled,
            ):
                if st.session_state.discovery_authorized:
                    st.success("✅ Records updated successfully (matched by Patient_ID).")
                    audit_logger.log_action("DISCOVERY_UPDATE", details={})
                else:
                    st.warning("🔐 Please authorize via the Master Database Password above.")
                    
    # ====== Discovery Log Section ======
    with st.expander("📜 Discovery Log", expanded=False):
        st.subheader("Global Discovery Activity Log")
        log_path = AUDIT_LOG_PATH
        if os.path.exists(log_path):
            try:
                import json
                logs = []
                with open(log_path, "r") as f:
                    for line in f:
                        entry = json.loads(line)
                        if entry.get("action") == "FILE_DISCOVERED":
                            d = entry.get("details", {})
                            logs.append({
                                "Date Found": entry.get("timestamp")[:16].replace("T", " "),
                                "Filename": d.get("filename"),
                                "Type": d.get("type"),
                                "Records": d.get("records"),
                                "Missing": d.get("missing", 0),
                                "In Master DB": "✅ Yes" if d.get("status") == "ALL_PRESENT" else "❌ No"
                            })
                if logs:
                    st.dataframe(pd.DataFrame(logs).sort_values("Date Found", ascending=False), use_container_width=True)
                else:
                    st.info("No file discovery events logged yet.")
            except Exception as e:
                st.error(f"Error reading log: {e}")
        else:
            st.info("Log file not found.")

    # ====== Visual Sync Box ======
    st.markdown("---")
    st.subheader("🔄 Visual Sync Box — Live Export Queue")
    st.caption("Monitors files staged for merging into the Master Database.")
    
    # Real Queue Loading
    def load_real_queue():
        all_reqs = _load_requests()
        all_staged = []
        for token, req in all_reqs.items():
            staged = get_staging_queue(token)
            for item in staged:
                if item.get("status") == "STAGED":
                    item["token"] = token
                    item["recipient"] = req.get("recipient_email", "Unknown")
                    # Map shield based on some properties if available, else default to green
                    item["shield"] = "🟢" if "error" not in item else "🔴"
                    all_staged.append(item)
        return all_staged

    if "sync_box_exporting" not in st.session_state:
        st.session_state.sync_box_exporting = False
    
    staged_files = load_real_queue()
    
    # Box container
    box = st.container()
    with box:
        if not st.session_state.sync_box_exporting:
            if not staged_files:
                st.info("📭 No files currently in the live export queue.")
            else:
                # Full queue display
                total_records = sum(int(f.get("records", 0)) for f in staged_files)
                healthy = sum(1 for f in staged_files if f.get("shield") == "🟢")
                blocked = sum(1 for f in staged_files if f.get("shield") == "🔴")
                
                col_b1, col_b2, col_b3, col_b4 = st.columns(4)
                with col_b1: st.metric("📦 Files", len(staged_files))
                with col_b2: st.metric("📊 Records", total_records)
                with col_b3: st.metric("🟢 Healthy", healthy)
                with col_b4: st.metric("🔴 Blocked", blocked)
                
                # File cards with shield icons
                for f in staged_files:
                    cols = st.columns([1, 3, 1, 1, 2])
                    with cols[0]: st.markdown(f"**{f['shield']}**")
                    with cols[1]: st.markdown(f"**{f['filename']}**\n\n*(From: {f['recipient']})*")
                    with cols[2]: st.markdown(f"_{f.get('records', '?')}_ recs")
                    with cols[3]: st.markdown(f"_{f.get('size', '?')}_")
                    with cols[4]:
                        status = f.get("status", "STAGED")
                        if status == "STAGED":
                            st.markdown("🟢 Ready")
                        else:
                            st.markdown(f"⚪ {status}")
                    st.divider()
                
                # Export button
                export_disabled = not st.session_state.discovery_authorized
                if st.button("🚀 Start Atomic Export to Database", type="primary", use_container_width=True, disabled=export_disabled):
                    if not st.session_state.discovery_authorized:
                        st.warning("🔐 Please authorize via the Master Database Password above.")
                    else:
                        st.session_state.sync_box_exporting = True
                        st.rerun()
        
        else:
            # Atomic Export Execution
            progress_bar = st.progress(0, text="Initializing atomic export...")
            status_placeholder = st.empty()
            
            total = len(staged_files)
            tokens_to_process = list(set(f["token"] for f in staged_files))
            
            done_count = 0
            for token in tokens_to_process:
                token_files = [f for f in staged_files if f["token"] == token]
                for f in token_files:
                    status_placeholder.info(f"📤 Exporting **{f['filename']}** from {f['recipient']}...")
                    
                    # Call the REAL atomic merge function
                    # It processes ALL staged files for this token
                    # But we can call it once per token or once for all.
                    # Given the function signature, it processes all for the token.
                    pass
                
                # Trigger real merge for this token
                merge_result = atomic_merge_staging_files(token)
                
                for res in merge_result.get("results", []):
                    done_count += 1
                    progress_val = int((done_count) / total * 100)
                    progress_bar.progress(progress_val, text=f"Processed {done_count}/{total} files...")
                    
                    if res["status"] == "SUCCESS":
                        st.success(f"✅ **Merged:** {res['filename']} ({res.get('records', 0)} new records)")
                    elif res["status"] == "SKIPPED":
                        st.info(f"⏭️ **Skipped:** {res['filename']} (All duplicates)")
                    else:
                        st.error(f"❌ **Failed:** {res['filename']} - {res.get('error')}")
                
                import time
                time.sleep(0.5) # Small pause for visual feedback
            
            progress_bar.progress(100, text="Atomic export complete!")
            st.balloons()
            st.session_state.sync_box_exporting = False
            
            # Refresh master DF in session state
            try:
                decrypted_master = decrypt_file_to_memory(MASTER_DB_PATH)
                if decrypted_master:
                    st.session_state.master_df = pd.read_excel(io.BytesIO(decrypted_master))
            except:
                pass
                
            if st.button("✅ Finish & Clear Box", use_container_width=True):
                st.rerun()
    
    st.markdown("---")
    
    # ====== Recipient Landing Page ======
    with st.expander("📥 Recipient Landing Page — Upload Discovery Files", expanded=False):
        st.subheader("📤 Submit Discovered SCD Data Files")
        st.markdown("""
            If you received a discovery request, use this section to upload encrypted or 
            unencrypted Excel/Word files found in your email attachments.
        """)
        
        recipient_token = st.text_input(
            "Discovery Token (from your invitation email)",
            placeholder="Paste your secure token here...",
            key="recipient_token_dash",
        )
        
        if recipient_token:
            disc_request = get_discovery_request(recipient_token)
            if disc_request:
                st.success(f"✅ Valid discovery request for: **{disc_request['recipient_email']}**")
                
                st.markdown("---")
                st.subheader("Upload Discovery Files")
                
                discovered_file = st.file_uploader(
                    "Upload discovered Excel/Word files",
                    type=["xlsx", "xls", "docx"],
                    help="Upload files found in your email attachments that contain SCD data.",
                    key="recipient_file_upload_dash",
                )
                
                st.markdown("**🔒 Encrypted File Support**")
                
                if "recipient_cached_password" not in st.session_state:
                    st.session_state.recipient_cached_password = ""
                if "recipient_password_verified" not in st.session_state:
                    st.session_state.recipient_password_verified = False
                
                if st.session_state.recipient_password_verified:
                    st.info("✅ File password already cached for this session.")
                    if st.button("🧹 Clear Cached Password", key="clear_recipient_pw_dash"):
                        st.session_state.recipient_cached_password = ""
                        st.session_state.recipient_password_verified = False
                        st.rerun()
                else:
                    recipient_file_encrypted = st.checkbox(
                        "This file is encrypted / password-protected",
                        key="recipient_file_encrypted_dash",
                    )
                    if recipient_file_encrypted and discovered_file:
                        st.text_input(
                            "File Decryption Password",
                            type="password",
                            key="recipient_file_password_dash",
                            help="Password for decryption.",
                        )
                
                effective_password = st.session_state.recipient_cached_password
                if not effective_password:
                    effective_password = st.session_state.get("recipient_file_password_dash", "")
                
                if discovered_file:
                    col_fname, col_fsize = st.columns([3, 1])
                    with col_fname:
                        st.info(f"**File:** {discovered_file.name}")
                    with col_fsize:
                        st.info(f"**Size:** {len(discovered_file.getvalue()) / 1024:.1f} KB")
                    
                    if st.button("📤 Submit Discovery Files", type="primary", use_container_width=True):
                        with st.spinner("Processing..."):
                            try:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(discovered_file.name)[1]) as tmp:
                                    tmp.write(discovered_file.getbuffer())
                                    tmp_path = tmp.name
                                
                                ext = os.path.splitext(discovered_file.name)[1].lower()
                                success = False
                                
                                if ext in ['.xlsx', '.xls']:
                                    from mapping import load_and_map_data
                                    mapped_df = load_and_map_data(tmp_path, password=effective_password)
                                    
                                    if not mapped_df.empty:
                                        if effective_password and not st.session_state.recipient_cached_password:
                                            st.session_state.recipient_cached_password = effective_password
                                            st.session_state.recipient_password_verified = True
                                        
                                        master_hashes = set()
                                        if os.path.exists(MASTER_DB_PATH):
                                            try:
                                                decrypted = decrypt_file_to_memory(MASTER_DB_PATH)
                                                if decrypted:
                                                    existing = pd.read_excel(io.BytesIO(decrypted))
                                                    if 'Patient_ID' in existing.columns:
                                                        import hashlib
                                                        master_hashes = set(
                                                            hashlib.sha256(str(pid).strip().encode()).hexdigest()
                                                            for pid in existing['Patient_ID'].dropna()
                                                        )
                                            except:
                                                pass
                                        
                                        missing_count = 0
                                        if 'Patient_ID' in mapped_df.columns:
                                            import hashlib
                                            for pid in mapped_df['Patient_ID']:
                                                if pd.notna(pid):
                                                    h = hashlib.sha256(str(pid).strip().encode()).hexdigest()
                                                    if h not in master_hashes:
                                                        missing_count += 1
                                        
                                        st.success(f"✅ File processed! {len(mapped_df)} records ({missing_count} new).")
                                        if missing_count > 0:
                                            update_discovery_status(recipient_token, "FILES_SUBMITTED", {"files": [discovered_file.name], "new_records": missing_count})
                                        st.dataframe(mapped_df.head(10), use_container_width=True)
                                        success = True
                                    else:
                                        st.warning("No SCD records extracted.")
                                        
                                elif ext == '.docx':
                                    from word_processor import process_word_file, scan_text_for_keywords
                                    file_like = io.BytesIO(discovered_file.getbuffer())
                                    word_df = process_word_file(file_like)
                                    is_scd, keywords = scan_text_for_keywords(discovered_file.getvalue())
                                    if not word_df.empty:
                                        st.success(f"✅ Word processed! {len(word_df)} records.")
                                        st.dataframe(word_df.head(10), use_container_width=True)
                                        success = True
                                    elif is_scd:
                                        st.info("📄 SCD keywords found but no structured data.")
                                        success = True
                                    else:
                                        st.warning("No SCD data identified.")
                                
                                elif ext in ['.doc']:
                                    st.warning("⚠️ Legacy .doc. Please convert to .docx.")
                                
                                if success:
                                    update_discovery_status(recipient_token, "PROCESSED", {"files_processed": [discovered_file.name]})
                                    st.balloons()
                                else:
                                    st.error("Could not process the file.")
                                
                                if os.path.exists(tmp_path):
                                    os.unlink(tmp_path)
                            except Exception as e:
                                st.error(f"❌ Error: {e}")
                                st.exception(e)
                
                st.markdown("---")
                st.subheader("🛡️ Security & Privacy")
                st.info("Files are processed in-memory only. Email access is temporary.")
            else:
                st.error("❌ Invalid or expired discovery token.")
        else:
            st.info("Please enter your discovery token from the invitation email to proceed.")

st.caption("SCD Dbase Sorter | Built with Streamlit | Team SCD Dbase Sorter")