import pandas as pd
import numpy as np
import os
import io
import msoffcrypto
import json

from sanitization import sanitize_dataframe
from logger import audit_logger

# Standard Master Headings
MASTER_HEADINGS = [
    "Patient_ID", "Patient_Name", "Hospital", "Year", "Validation_Status", 
    "Validator_Email", "Hospital_Email", "Date_Added", "Treatment", "Outcome"
]

ALIAS_FILE = "/home/team/shared/SCD_Dbase_Sorter/data/config/aliases.json"

def load_aliases():
    """Loads aliases from the JSON file."""
    if os.path.exists(ALIAS_FILE):
        with open(ALIAS_FILE, 'r') as f:
            return json.load(f)
    return {
        "Hospital": ["Hosp_Name", "Hosp Name", "Facility", "Center", "Hospital Name"],
        "Year": ["Yr", "Data_Year", "Period", "Year of Data"],
        "Patient_ID": ["Pt_No", "Patient ID", "ID", "Case_No", "Patient_ID"],
        "Patient_Name": ["Name", "Patient Name", "Full Name", "Pt Name"],
        "Treatment": ["Rx", "Therapy", "Treatment"],
        "Outcome": ["Result", "Status", "Outcome"]
    }

def save_new_alias(master_heading, new_alias):
    """Saves a new alias for a master heading."""
    aliases = load_aliases()
    if master_heading in aliases:
        if new_alias not in aliases[master_heading]:
            aliases[master_heading].append(new_alias)
    else:
        aliases[master_heading] = [new_alias]
    
    os.makedirs(os.path.dirname(ALIAS_FILE), exist_ok=True)
    with open(ALIAS_FILE, 'w') as f:
        json.dump(aliases, f, indent=4)
    audit_logger.log_action("SAVE_ALIAS", details={"master": master_heading, "alias": new_alias})

def _get_excel_data(file_input, password=None):
    """
    Helper to handle password-protected Excel files.
    Returns a file-like object or file path.
    file_input can be a path (str) or file-like object.
    """
    if password:
        decrypted_data = io.BytesIO()
        if isinstance(file_input, str):
            with open(file_input, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted_data)
        else:
            # Assume file-like
            office_file = msoffcrypto.OfficeFile(file_input)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_data)
        decrypted_data.seek(0)
        return decrypted_data
    return file_input

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
    aliases = load_aliases()
    for master, alias_list in aliases.items():
        for alias in alias_list:
            if alias.lower() == clean_text:
                return master
                
    return None

def get_column_mapping(file_input, password=None):
    """
    Analyzes the first two rows of the Excel file to determine column mapping.
    Returns a dict: {column_index: master_heading}
    """
    excel_data = _get_excel_data(file_input, password)
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

def load_and_map_data(file_input, password=None, custom_mapping=None):
    """
    Loads data from file_input, applies mapping, and returns a normalized DataFrame.
    
    Args:
        file_input: Path to the Excel file or file-like object.
        password: Password for encrypted files.
        custom_mapping: Optional dict {col_idx: master_heading}. If None, uses auto-detection.
    """
    if custom_mapping:
        mapping = custom_mapping
    else:
        mapping = get_column_mapping(file_input, password)
    
    excel_data = _get_excel_data(file_input, password)
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
    
    file_label = file_input if isinstance(file_input, str) else "file-like-object"
    audit_logger.log_action("LOAD_AND_MAP", details={"file": file_label, "records": len(df_mapped)})
    
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
