import pandas as pd
import os
import io
import json
import hashlib
from datetime import datetime
from encryption import encrypt_file, decrypt_file_to_memory
from logger import audit_logger
from hashing_service import get_master_patient_hashes, compare_hashes
from mapping import load_and_map_data

try:
    from .config import MASTER_DB_PATH, HOSPITALS_DIR, STAGING_DIR as STAGING_BASE_DIR, QUEUE_FILE
except ImportError:
    from config import MASTER_DB_PATH, HOSPITALS_DIR, STAGING_DIR as STAGING_BASE_DIR, QUEUE_FILE

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

def update_queue_status_local(token, filename, status, details=None):
    """Local helper to update queue without full Flask dependency."""
    if not os.path.exists(QUEUE_FILE):
        return
    try:
        with open(QUEUE_FILE, 'r') as f:
            queue = json.load(f)
    except Exception:
        return
        
    if token in queue:
        for item in queue[token]:
            if item.get('filename') == filename:
                item['status'] = status
                item['updated_at'] = datetime.now().isoformat()
                if details:
                    item['details'] = details
                break
        with open(QUEUE_FILE, 'w') as f:
            json.dump(queue, f, indent=4)

def atomic_merge_staging_files(token):
    """
    Milestone 4: Atomic Export
    Merges all staged files for a token into the Master Database.
    Deletes files from staging ONLY after successful merge.
    """
    if not os.path.exists(QUEUE_FILE):
        return {"error": "Queue not found"}
        
    try:
        with open(QUEUE_FILE, 'r') as f:
            queue = json.load(f)
    except Exception as e:
        return {"error": f"Failed to read queue: {e}"}
        
    if token not in queue:
        return {"error": "Token not found in queue"}
        
    staged_items = [item for item in queue[token] if item.get('status') == 'STAGED']
    if not staged_items:
        return {"message": "No staged files to merge"}
        
    results = []
    master_hashes = get_master_patient_hashes()
    
    for item in staged_items:
        filename = item['filename']
        file_path = os.path.join(STAGING_BASE_DIR, token, filename)
        
        if not os.path.exists(file_path):
            update_queue_status_local(token, filename, "ERROR", "File missing on disk")
            continue
            
        try:
            # 1. Load, Map, and Heal (Heal logic is inside load_and_map_data)
            df = load_and_map_data(file_path)
            
            # Check for macros (extension-based check for logging)
            if filename.lower().endswith(('.xlsm', '.xlsb', '.docm')):
                audit_logger.log_action("SANITIZATION_MACRO", details={"file_source": filename, "msg": "Stripped VBA macros"})
            
            # 2. De-duplication
            if not df.empty and "Patient_ID" in df.columns:
                initial_count = len(df)
                # Filter out rows that are already in master
                is_duplicate = df["Patient_ID"].apply(lambda pid: hashlib.sha256(str(pid).strip().encode()).hexdigest() in master_hashes if pd.notna(pid) else False)
                df = df[~is_duplicate]
                removed_count = initial_count - len(df)
                if removed_count > 0:
                    audit_logger.log_action("DE_DUPLICATION", details={"file": filename, "duplicates_removed": removed_count})
            
            if not df.empty:
                # 3. Merge into Master DB
                process_new_data(df)
                
                # 4. Verify and Clean up
                # If we reached here without exception, assume success
                os.remove(file_path)
                update_queue_status_local(token, filename, "MERGED", f"Merged {len(df)} new records")
                audit_logger.log_action("ATOMIC_EXPORT", details={"file": filename, "status": "SUCCESS", "new_records": len(df)})
                results.append({"filename": filename, "status": "SUCCESS", "records": len(df)})
                
                # Refresh master hashes for next file in loop
                master_hashes = get_master_patient_hashes()
            else:
                os.remove(file_path)
                update_queue_status_local(token, filename, "MERGED", "No new records found (all duplicates)")
                results.append({"filename": filename, "status": "SKIPPED", "msg": "All duplicates"})
                
        except Exception as e:
            update_queue_status_local(token, filename, "FAILED", str(e))
            audit_logger.log_action("ATOMIC_EXPORT", details={"file": filename, "status": "FAILED", "error": str(e)})
            results.append({"filename": filename, "status": "FAILED", "error": str(e)})
            
    return {"results": results}

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
