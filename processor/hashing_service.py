import hashlib
import pandas as pd
import io
import os
from encryption import decrypt_file_to_memory

MASTER_DB_PATH = "/home/team/shared/SCD_Dbase_Sorter/data/master/Master_Database.xlsx"

def generate_hash(value):
    """Generates a SHA-256 hash for a given value."""
    if pd.isna(value):
        return None
    return hashlib.sha256(str(value).strip().encode()).hexdigest()

def get_master_patient_hashes():
    """
    Reads the Master Database and returns a set of SHA-256 hashes of Patient IDs.
    """
    if not os.path.exists(MASTER_DB_PATH):
        return set()

    try:
        # The database is encrypted at rest
        decrypted_data = decrypt_file_to_memory(MASTER_DB_PATH)
        if not decrypted_data:
            return set()
            
        df = pd.read_excel(io.BytesIO(decrypted_data))
        if "Patient_ID" not in df.columns:
            return set()
            
        # Generate hashes for all Patient IDs
        hashes = set(df["Patient_ID"].dropna().apply(generate_hash))
        return hashes
    except Exception as e:
        print(f"Error generating patient hashes: {e}")
        return set()

def compare_hashes(new_ids, master_hashes):
    """
    Compares a list of new IDs against the master hashes.
    Returns a list of booleans indicating if each ID is already in the database.
    """
    results = []
    for patient_id in new_ids:
        h = generate_hash(patient_id)
        results.append(h in master_hashes if h else False)
    return results
