import pandas as pd
import os
import io
import json
import hashlib
from datetime import datetime
from encryption import encrypt_file, decrypt_file_to_memory
from logger import audit_logger
from hashing_service import get_master_patient_hashes, compare_hashes
from mapping import load_and_map_data, MASTER_HEADINGS

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
            
            # Read only the Master_Data sheet as source of truth
            try:
                master_df = pd.read_excel(io.BytesIO(decrypted_data), sheet_name='Master_Data')
            except Exception:
                # Fallback: try reading first sheet
                master_df = pd.read_excel(io.BytesIO(decrypted_data), sheet_name=0)
            
            # Ensure Date_Added is datetime
            if 'Date_Added' in master_df.columns:
                master_df['Date_Added'] = pd.to_datetime(master_df['Date_Added'])
            
            combined_df = pd.concat([master_df, new_data_df], ignore_index=True)
            
            # Ensure Region column exists and is filled (Migration/Multi-tenancy)
            if 'Region' not in combined_df.columns:
                combined_df['Region'] = 'SERHA'
            else:
                combined_df['Region'] = combined_df['Region'].fillna('SERHA')
        except Exception as e:
            print(f"Error reading master database: {e}")
            audit_logger.log_action("ERROR", details={"msg": f"Error reading master database: {e}"})
            combined_df = new_data_df
    else:
        combined_df = new_data_df

    # We don't save here anymore, we let process_new_data call the refined generator
    return combined_df

def generate_hospital_sheets(master_df):
    """
    Owner Requirement: Master Database organized by year-of-birth tabs.
    Splits the master dataframe into:
    1. Individual hospital Excel files in Hospitals/ directory.
    2. DOB Year tabs in Master_Database.xlsx.
    3. Hospital tabs in Master_Database.xlsx (for reporting).
    """
    ensure_directories()
    
    if master_df.empty:
        return

    # 1. Generate individual hospital files
    hospital_groups = master_df.groupby('Hospital')
    for hospital, group in hospital_groups:
        if pd.isna(hospital) or str(hospital).strip() == "":
            hospital_name = "Unassigned"
        else:
            hospital_name = str(hospital).strip()
            
        safe_name = "".join([c for c in hospital_name if c.isalnum() or c in (' ', '_')]).strip()
        safe_filename = safe_name.replace(' ', '_')
        
        # Sort by Year (collection year) and Date_Added
        if 'Year' in group.columns:
            group_sorted = group.assign(Year_Numeric=pd.to_numeric(group['Year'], errors='coerce'))
            group_sorted = group_sorted.sort_values(by=['Year_Numeric', 'Date_Added'], ascending=[False, False])
            group_to_save = group_sorted.drop(columns=['Year_Numeric'])
        else:
            group_to_save = group.sort_values(by='Date_Added', ascending=False)
            
        file_path = os.path.join(HOSPITALS_DIR, f"{safe_filename}.xlsx")
        group_to_save.to_excel(file_path, index=False)
        encrypt_file(file_path)

    # 2. Update Master_Database.xlsx with Year-of-Birth tabs and Hospital tabs
    with pd.ExcelWriter(MASTER_DB_PATH, engine='openpyxl') as writer:
        # A. Master Data Tab (Source of Truth)
        master_df.to_excel(writer, sheet_name='Master_Data', index=False)
        
        # B. DOB Year Tabs (OWNER REQUIREMENT)
        if 'DOB' in master_df.columns:
            # Extract year from DOB
            master_df_with_dob_year = master_df.copy()
            # Ensure DOB is datetime
            master_df_with_dob_year['DOB'] = pd.to_datetime(master_df_with_dob_year['DOB'], errors='coerce')
            master_df_with_dob_year['DOB_Year'] = master_df_with_dob_year['DOB'].dt.year
            
            # Filter out records with no DOB year
            valid_dob_df = master_df_with_dob_year.dropna(subset=['DOB_Year'])
            if not valid_dob_df.empty:
                valid_dob_df = valid_dob_df.assign(DOB_Year=valid_dob_df['DOB_Year'].astype(int))
                year_groups = valid_dob_df.groupby('DOB_Year')
                
                # Sort years ascending (oldest to newest)
                sorted_years = sorted(year_groups.groups.keys())
                for year in sorted_years:
                    year_group = year_groups.get_group(year).drop(columns=['DOB_Year'])
                    writer.book.create_sheet(str(year))
                    year_group.to_excel(writer, sheet_name=str(year), index=False)
            
            # Ambiguous/Missing DOB records go to a specific tab
            missing_dob_df = master_df_with_dob_year[master_df_with_dob_year['DOB_Year'].isna()].drop(columns=['DOB_Year'])
            if not missing_dob_df.empty:
                missing_dob_df.to_excel(writer, sheet_name='Missing_DOB', index=False)

        # C. Hospital Tabs (Reporting)
        for hospital, group in hospital_groups:
            if pd.isna(hospital) or str(hospital).strip() == "":
                sheet_name = "Unassigned"
            else:
                sheet_name = "".join([c for c in str(hospital) if c.isalnum() or c in (' ', '_')])[:31].strip()
            
            group.to_excel(writer, sheet_name=sheet_name, index=False)
            
    # Finally encrypt the Master Database
    encrypt_file(MASTER_DB_PATH)
    audit_logger.log_action("GENERATE_MASTER_TABS", details={"hospitals": len(hospital_groups)})

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

def approve_and_merge_staged_file(token, filename):
    """
    Milestone 5: Lead Verification Queue
    Force merges a file that was flagged for review.
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
        
    item = next((i for i in queue[token] if i.get('filename') == filename), None)
    if not item:
        return {"error": "File not found in queue"}
        
    if item.get('status') != 'NEEDS_REVIEW':
        return {"error": f"File is in status {item.get('status')}, cannot approve."}
        
    file_path = os.path.join(STAGING_BASE_DIR, token, filename)
    if not os.path.exists(file_path):
        return {"error": "File missing on disk"}
        
    try:
        # Load and map (ignoring triggers because this is an explicit approval)
        df = load_and_map_data(file_path)
        
        # De-duplication
        master_hashes = get_master_patient_hashes()
        if not df.empty and "Patient_ID" in df.columns:
            is_duplicate = df["Patient_ID"].apply(lambda pid: hashlib.sha256(str(pid).strip().encode()).hexdigest() in master_hashes if pd.notna(pid) else False)
            df = df[~is_duplicate]
            
        if not df.empty:
            process_new_data(df)
            
            # Record approved aliases to Learning Loop (Protocol 5.3)
            # This is partly handled in load_and_map_data, but only if no triggers.
            # Here we can force save them because the lead approved.
            _save_approved_aliases(file_path)
            
            os.remove(file_path)
            update_queue_status_local(token, filename, "MERGED", "Lead approved and merged")
            audit_logger.log_action("LEAD_APPROVAL", details={"file": filename, "token": token})
            return {"status": "SUCCESS", "records": len(df)}
        else:
            os.remove(file_path)
            update_queue_status_local(token, filename, "MERGED", "Lead approved (all duplicates)")
            return {"status": "SUCCESS", "msg": "All duplicates"}
            
    except Exception as e:
        return {"error": str(e)}

def _save_approved_aliases(file_path):
    """Extracts and saves aliases from an approved file."""
    try:
        from mapping import get_column_mapping, save_new_alias, find_master_match, load_aliases
        mapping, header_row_idx, _ = get_column_mapping(file_path)
        
        import pandas as pd
        df_header_row = pd.read_excel(file_path, header=None, skiprows=header_row_idx, nrows=1)
        
        for col_idx, master_name in mapping.items():
            if col_idx < df_header_row.shape[1]:
                original_header = str(df_header_row.iloc[0, col_idx]).strip()
                if original_header and original_header.lower() != master_name.lower():
                    if original_header not in load_aliases().get(master_name, []):
                        save_new_alias(master_name, original_header)
    except Exception:
        pass

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
            
            # Milestone 5: Accuracy Guardrails - Check for Review Flag
            if not df.empty and df['Review_Required'].any():
                triggers = df['Review_Triggers'].iloc[0]
                # Extract the column mapping used for informative review
                cols_found = [c for c in df.columns if c in MASTER_HEADINGS]
                mapping_str = ", ".join(cols_found)
                detail_msg = f"Requires lead approval due to: {triggers}. Suggested columns: {mapping_str}"
                
                update_queue_status_local(token, filename, "NEEDS_REVIEW", detail_msg)
                audit_logger.log_action("ACCURACY_GUARDRAIL", details={"file": filename, "status": "NEEDS_REVIEW", "triggers": triggers, "mapping": mapping_str})
                results.append({"filename": filename, "status": "NEEDS_REVIEW", "triggers": triggers})
                continue
            
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
