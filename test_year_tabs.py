import pandas as pd
import os
import sys
import io

# Add processor to path
sys.path.append(os.path.join(os.path.dirname(__file__), "processor"))

from sorter import process_new_data
from config import MASTER_DB_PATH
from encryption import decrypt_file_to_memory

def test_year_tabs():
    # 1. Create test data
    data = {
        'Patient_ID': ['P1', 'P2', 'P3', 'P4', 'P5'],
        'Patient_Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
        'Hospital': ['Hosp A', 'Hosp A', 'Hosp B', 'Hosp B', 'Hosp A'],
        'Region': ['North', 'North', 'South', 'South', 'North'],
        'Year': [2026, 2026, 2026, 2026, 2026],
        'DOB': ['2020-01-01', '2021-05-15', '2020-11-20', '2022-03-10', 'Invalid Date'],
        'Treatment': ['None', 'A', 'B', 'C', 'None'],
        'Outcome': ['Good', 'Better', 'Best', 'OK', 'Unknown']
    }
    df = pd.DataFrame(data)
    
    # Pre-process DF to match what load_and_map_data would return
    df['Date_Added'] = pd.Timestamp.now()
    
    print("Processing test data...")
    process_new_data(df)
    
    print("Verifying Master_Database.xlsx...")
    if os.path.exists(MASTER_DB_PATH):
        decrypted_data = decrypt_file_to_memory(MASTER_DB_PATH)
        xlsx = pd.ExcelFile(io.BytesIO(decrypted_data))
        print(f"Sheet names found: {xlsx.sheet_names}")
        
        expected_sheets = ['Master_Data', '2020', '2021', '2022', 'Missing_DOB', 'Hosp_A', 'Hosp_B']
        for s in expected_sheets:
            if s in xlsx.sheet_names:
                print(f"✓ Found expected sheet: {s}")
                sheet_df = pd.read_excel(xlsx, sheet_name=s)
                print(f"  Records in {s}: {len(sheet_df)}")
            else:
                print(f"✗ Missing expected sheet: {s}")
    else:
        print("✗ Master_Database.xlsx not found!")

if __name__ == "__main__":
    # Clean up previous runs if needed
    if os.path.exists(MASTER_DB_PATH):
        os.remove(MASTER_DB_PATH)
    
    test_year_tabs()
