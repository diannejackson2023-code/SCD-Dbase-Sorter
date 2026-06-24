# SCD Dbase Sorter - System Architecture Design

## 1. Project Overview
The SCD Dbase Sorter is an automated system designed to manage and sort medical data from Excel sheets. It centralizes data into a master database and distributes it into hospital-specific sheets, facilitating validation and communication with hospital contacts.

## 2. Core Components

### 2.1. Input Interface (The Dashboard)
- **Technology**: Streamlit (Python-based web dashboard).
- **Features**:
    - File uploader for new Excel data.
    - "Process & Sort" button.
    - "Validate & Email" buttons for each hospital.
    - Status overview (data counts per hospital/year).

### 2.2. Data Processing Engine
- **Technology**: Python with `pandas` and `openpyxl`.
- **Functions**:
    - **Header Mapping**: Logic to identify columns even if headers are in Row 1, Row 2, or split across both.
    - **Normalization**: Standardizing hospital names and year formats.
    - **Distribution**: Splitting master data into individual hospital records.

### 2.3. Storage Layer
- **Primary Storage**: A Master Excel Workbook (`Master_Database.xlsx`).
- **Dynamic Sheets**: Automatically generated tabs for each hospital.
- **Reference Data**: A configuration file (or sheet) mapping Hospitals to Email Addresses and Validators.

## 3. Detailed Logic

### 3.1. Excel Sheet Structure
#### Master Database Columns (Recommended)
| Column Name | Type | Description |
|-------------|------|-------------|
| Patient_ID  | String | Unique identifier |
| Hospital    | String | Full name of the hospital |
| Year        | Integer| Year of data |
| Validation_Status | Enum | Pending, Validated, Rejected |
| Validator_Email | String | Email of the person who validated |
| Date_Added  | DateTime | When the record was pasted |
| ...         | ...    | Medical specific fields (e.g., Treatment, Outcome) |

#### Mapping Config (Reference Sheet)
| Source Heading | Target Master Heading |
|----------------|-----------------------|
| Hosp_Name      | Hospital              |
| Yr             | Year                  |
| Pt_No          | Patient_ID            |

### 3.2. Mapping Logic (Matching Headings)
- The system will scan Row 1 and Row 2 of the uploaded sheet.
- It will compare found strings against a "Header Map" using:
    1. Exact Match.
    2. Case-insensitive Match.
    3. Alias Match (e.g., "Hosp Name" -> "Hospital Name").
- Data under matched headings is appended to the Master Database under the correct column.

### 3.3. Autosorting (Hospital & Year)
- Every record must have a `Hospital` and `Year` value.
- Upon processing, the system:
    1. Groups data by `Hospital`.
    2. For each Hospital, checks if a sheet exists in the `Master_Database.xlsx` or a separate `Hospitals/` directory.
    3. Appends/Updates data in the specific hospital sheet.
    4. Sub-sorts by `Year` within those sheets.

### 3.4. Email & Validation Workflow
- **Validation**: A "Validator" (assigned per hospital or globally) reviews the data.
- **Trigger**: Dashboard button "Send Validation Request".
- **Execution**: 
    - System looks up the `Validator Email` for the hospital.
    - Sends an email (via SMTP or Outlook Integration) with the hospital's data attached or a link to it.
- **Hospital Emailing**: After validation, "Send to Hospital" button sends the finalized data to the `Hospital Email`.

## 4. Directory Structure
```text
/home/team/shared/SCD_Dbase_Sorter/
├── app.py                # Streamlit Dashboard
├── processor/
│   ├── mapping.py        # Header mapping logic
│   ├── sorter.py         # Autosort and sheet generation
│   └── mailer.py         # Email automation
├── data/
│   ├── master/           # Master Excel files
│   ├── hospitals/        # Generated hospital-specific files
│   └── config/           # Hospital-Email mapping (csv/xlsx)
└── DESIGN.md             # This document
```

## 5. Technology Stack
- **Language**: Python 3.10+
- **Data Handling**: `pandas`, `openpyxl`
- **Dashboard**: `Streamlit`
- **Emailing**: `smtplib` or `MS Graph API` (depending on environment)
- **Deployment**: Local network or Sandbox server

## 6. Communication Protocols
- The **Dashboard** communicates with the **Processor** via function calls.
- The **Processor** reads/writes directly to the **Data** directory.
- Validation status is tracked in a hidden column in the Master Excel or a sidecar database/JSON.
