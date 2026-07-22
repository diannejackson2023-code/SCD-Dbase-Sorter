import pandas as pd
import numpy as np
import os
import io
import msoffcrypto
import json
import re

from sanitization import sanitize_dataframe
from logger import audit_logger

# Standard Master Headings
MASTER_HEADINGS = [
    "Patient_ID", "Patient_Name", "Hospital", "Year", "Validation_Status", 
    "Validator_Email", "Hospital_Email", "Date_Added", "Treatment", "Outcome"
]

# Learned Patterns for Header Recovery
PATTERNS = {
    "Patient_ID": r"^[A-Z0-9]{4,12}$",
    "Validator_Email": r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$",
    "Hospital_Email": r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$",
    "Year": r"^(19|20)\d{2}$"
}

def infer_heading_from_data(data_sample):
    """
    Attempts to infer the master heading based on data patterns.
    """
    if data_sample.empty:
        return None
        
    for master, pattern in PATTERNS.items():
        # Check if at least 50% of non-null samples match the pattern
        valid_samples = data_sample.dropna()
        if valid_samples.empty:
            continue
            
        matches = valid_samples.astype(str).apply(lambda x: bool(re.match(pattern, x)))
        if matches.mean() >= 0.5:
            return master
    return None

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

def levenshtein_distance(s1, s2):
    if len(s1) < len(s2):
        return levenshtein_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)
    previous_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        current_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = previous_row[j + 1] + 1
            deletions = current_row[j] + 1
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row
    return previous_row[-1]

def find_master_match(header_text, use_fuzzy=True):
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
    
    # Fuzzy Matching (Milestone 4)
    if use_fuzzy:
        best_match = None
        min_dist = 3 # Levenshtein distance <= 2
        
        # Check against master headings
        for master in MASTER_HEADINGS:
            dist = levenshtein_distance(master.lower(), clean_text)
            if dist < min_dist:
                min_dist = dist
                best_match = master
        
        # Check against aliases
        for master, alias_list in aliases.items():
            for alias in alias_list:
                dist = levenshtein_distance(alias.lower(), clean_text)
                if dist < min_dist:
                    min_dist = dist
                    best_match = master
        
        if best_match:
            # Healing Action: Add learned alias
            # We don't save it immediately here to avoid side effects during scanning
            # but we return the match
            return best_match
                
    return None

def get_column_mapping(file_input, password=None):
    """
    Analyzes the rows of the Excel file to determine column mapping.
    Scans up to Row 10 to find a valid header row.
    Returns a tuple: (mapping_dict, header_row_index)
    """
    excel_data = _get_excel_data(file_input, password)
    # Read first 20 rows to have enough data for inference if needed
    try:
        df_scan = pd.read_excel(excel_data, header=None, nrows=20)
    except Exception as e:
        audit_logger.log_action("ERROR", details={"msg": f"Failed to read file for mapping: {e}"})
        return {}, 0
    
    num_rows = df_scan.shape[0]
    num_cols = df_scan.shape[1]
    
    best_row_idx = 0
    best_mapping = {}
    max_matches = 0
    
    # Milestone 4: Search up to Row 10
    scan_limit = min(num_rows, 10)
    for row_idx in range(scan_limit):
        current_mapping = {}
        matches = 0
        
        # Merged Cell Unstacking (Protocol 3.3)
        last_match = None
        
        for col_idx in range(num_cols):
            val = df_scan.iloc[row_idx, col_idx]
            
            # If empty, try unstacking from previous column in the same row
            if pd.isna(val) or str(val).strip() == "":
                if last_match:
                    current_mapping[col_idx] = last_match
                    continue
            
            match = find_master_match(val, use_fuzzy=True)
            if match:
                current_mapping[col_idx] = match
                matches += 1
                last_match = match
                
                # Log healing if fuzzy
                original_text = str(val).strip()
                if original_text.lower() != match.lower() and original_text not in load_aliases().get(match, []):
                     file_label = file_input if isinstance(file_input, str) else "stream"
                     audit_logger.log_action("HEALING_HEADER", details={
                         "file_source": file_label,
                         "details": f"Healed '{original_text}' to '{match}'"
                     })
            else:
                last_match = None
        
        if matches > max_matches:
            max_matches = matches
            best_mapping = current_mapping
            best_row_idx = row_idx
            
        # Check 30% threshold for deep scan rows (as per Sanitization Protocol 3.2)
        if row_idx > 1 and num_cols > 0:
            if (matches / num_cols) >= 0.3:
                break
                
    # Empty Header Recovery (Protocol 3.2)
    if num_rows > best_row_idx + 1:
        data_area = df_scan.iloc[best_row_idx + 1:]
        for col_idx in range(num_cols):
            if col_idx not in best_mapping:
                inferred = infer_heading_from_data(data_area.iloc[:, col_idx])
                if inferred:
                    best_mapping[col_idx] = inferred
                    file_label = file_input if isinstance(file_input, str) else "stream"
                    audit_logger.log_action("HEALING_HEADER", details={
                        "file_source": file_label,
                        "details": f"Inferred header for column {col_idx} as '{inferred}' based on data pattern"
                    })

    # If no matches found in deep scan, try combined Row 1 & 2 logic as fallback
    if max_matches == 0 and num_rows >= 2:
         for col_idx in range(num_cols):
             r1 = df_scan.iloc[0, col_idx]
             r2 = df_scan.iloc[1, col_idx]
             
             match = find_master_match(r1) or find_master_match(r2)
             if match:
                 best_mapping[col_idx] = match
         best_row_idx = 1
         
    return best_mapping, best_row_idx

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
    """
    if custom_mapping:
        mapping = custom_mapping
        header_row_idx = 1
    else:
        mapping, header_row_idx = get_column_mapping(file_input, password)

    excel_data = _get_excel_data(file_input, password)
    # Read the data
    df = pd.read_excel(excel_data, header=None, skiprows=header_row_idx + 1)

    # Rename columns based on mapping
    df_mapped = pd.DataFrame()

    # Deduplication (Protocol 3.3): Merge columns mapping to same Master Heading
    master_to_cols = {}
    for col_idx, master_name in mapping.items():
        if col_idx < df.shape[1]:
            if master_name not in master_to_cols:
                master_to_cols[master_name] = []
            master_to_cols[master_name].append(df.iloc[:, col_idx])

    for master_name, series_list in master_to_cols.items():
        if len(series_list) > 1:
            # Merge: prioritize non-null values
            merged = series_list[0].copy()
            for i in range(1, len(series_list)):
                merged = merged.fillna(series_list[i])
            df_mapped[master_name] = merged

            file_label = file_input if isinstance(file_input, str) else "stream"
            audit_logger.log_action("HEALING_HEADER", details={
                "file_source": file_label,
                "details": f"Merged multiple columns for '{master_name}'"
            })
        else:
            df_mapped[master_name] = series_list[0]

    # Add default values for missing master columns
    for master in MASTER_HEADINGS:
        if master not in df_mapped.columns:
            df_mapped[master] = np.nan

    # Apply PII masking
    df_mapped = mask_pii_data(df_mapped)

    # Apply Sanitization
    df_mapped = sanitize_dataframe(df_mapped)

    # Set default Validation_Status if missing
    if 'Validation_Status' not in df_mapped.columns or df_mapped['Validation_Status'].isnull().all():
        df_mapped['Validation_Status'] = 'Pending'

    # Set Date_Added
    df_mapped['Date_Added'] = pd.Timestamp.now()

    # Record new aliases
    excel_data_orig = _get_excel_data(file_input, password)
    try:
        df_header_row = pd.read_excel(excel_data_orig, header=None, skiprows=header_row_idx, nrows=1)
        for col_idx, master_name in mapping.items():
            if col_idx < df_header_row.shape[1]:
                original_header = str(df_header_row.iloc[0, col_idx]).strip()
                if original_header and original_header.lower() != master_name.lower():
                    # Check if already an alias
                    if original_header not in load_aliases().get(master_name, []):
                        save_new_alias(master_name, original_header)
    except Exception:
        pass

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
