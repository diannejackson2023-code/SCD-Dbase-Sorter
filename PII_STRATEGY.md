# PII Security Strategy - SCD Dbase Sorter

This document outlines the security measures implemented to protect Personally Identifiable Information (PII) within the SCD Dbase Sorter system.

## Level 1: Source File Protection (Reading Encrypted Excels)
The system supports the ingestion of password-protected Excel files (.xlsx). 
- **Mechanism**: Data providers can encrypt their Excel files before transmission.
- **Implementation**: The `processor/mapping.py` uses `msoffcrypto-tool` to decrypt files in-memory using a password provided via the dashboard. This ensures that PII is never stored in unencrypted form on disk during the ingestion phase if the source is protected.

## Level 2: Transport Security (Email Communication)
All data shared via email (validation requests and finalized reports) is protected during transit.
- **Mechanism**: The `mailer.py` module is configured to use STARTTLS/SSL for secure communication with the SMTP server.
- **Status**: Implemented.

## Level 3: Application Data Masking
To prevent the exposure of sensitive patient names within the dashboard or secondary reports, masking strategies are employed.
- **Strategy**: 
    - **Patient_ID over Names**: The system primarily uses anonymized `Patient_ID` for all internal processing.
    - **Name Masking**: If patient names are ingested, they should be masked (e.g., "John Doe" -> "J*** D**") in all non-validated views or removed before distribution if not strictly necessary.
- **Implementation**: Masking logic can be applied in `processor/mapping.py` during the normalization phase.

## Level 4: At-Rest Encryption
Data stored in the `data/` directory, including the `Master_Database.xlsx` and hospital-specific sheets, must be secured.
- **Mechanism**: 
    - **Excel Encryption**: The `Master_Database.xlsx` can be saved with a password to prevent unauthorized access.
    - **OS Level**: It is recommended to deploy the application on a filesystem with full-disk encryption (e.g., LUKS on Linux).
- **Implementation**: Future enhancement to `processor/sorter.py` to support saving encrypted workbooks.

## Summary of Security Controls
| Control Level | Component | Status |
|---------------|-----------|--------|
| Source Protection | `processor/mapping.py` | Implemented (Supports password-protected input) |
| Transport Security | `processor/mailer.py` | Implemented (TLS/SSL) |
| Data Masking | `processor/mapping.py` | Strategy defined (Patient_ID primary) |
| Storage Security | `data/` directory | Strategy defined (FDE / Excel Password) |
