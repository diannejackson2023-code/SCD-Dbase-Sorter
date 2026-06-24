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
