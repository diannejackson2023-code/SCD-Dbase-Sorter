# Security & Sanitization Protocol (Milestone 4)

This document defines the protocols for handling external SCD data discovery, focusing on security sanitization and structural repair of ingested medical records.

## 1. Formula Injection Defense (CSV/Excel)
To prevent malicious formula execution in spreadsheet applications (Excel, Google Sheets), all string-based inputs must be neutralized before being stored or displayed.

### 1.1 Neutralization Rules
Any cell value starting with the following characters must be neutralized:
- `=` (Equals)
- `+` (Plus)
- `-` (Minus)
- `@` (At symbol)

### 1.2 Fixing Mechanism
- **Prepend Single Quote**: The companion app and backend ingress must prepend a single quote (`'`) to the string.
- **Example**: `=SUM(A1:A10)` becomes `'=SUM(A1:A10)`.
- **Purpose**: This forces the spreadsheet software to interpret the content as literal text rather than an executable formula, neutralizing "CSV Injection" or "Formula Injection" attacks.

---

## 2. VBA Macro Protocol
VBA (Visual Basic for Applications) macros are a common vector for malware delivery.

### 2.1 Detection
- Files with extensions `.xlsm`, `.xlsb`, `.docm`, or legacy `.xls`/`.doc` must be flagged for macro presence.

### 2.2 Handling Policy: Automated Stripping
The system prioritizes security over macro preservation.
- **Action**: All VBA macros and `vbaProject.bin` streams must be stripped from the file during the sanitization phase.
- **Implementation**: 
    - For `.xlsx` conversion from `.xlsm`, use `openpyxl` or `pandas` which discard macros by default.
    - For `.docx` from `.docm`, use `python-docx` to extract text/tables into a fresh, macro-free document.
- **User Warning**: The UI must display a non-blocking notification: *"Security Check: Macros detected and removed from [Filename] to ensure safe processing."*

---

## 3. Structural Healing (Header Repair)
"Healing" refers to the automated correction of broken, misspelled, or structurally non-standard Excel headers using the system's persistent Intelligence.

### 3.1 Alias-Driven Healing
- **Knowledge Base**: Use the `aliases.json` Knowledge Base (persistent mapping history).
- **Fuzzy Matching**: If a header does not have an exact match or known alias, perform a fuzzy match (Levenshtein distance <= 2) against the alias list.
- **Healing Action**: If a high-confidence match is found, rename the column to the Master Heading and add the misspelled version to the `aliases.json` as a "learned" variation for future automation.

### 3.2 Positional & Contextual Healing
- **Header Search Depth**: If Row 1 and Row 2 do not yield valid mappings, scan up to Row 10 to find a row where at least 30% of columns match known aliases.
- **Empty Header Recovery**: If a column has data but no header, check the data format. If the data matches a "Learned Pattern" (e.g., `[A-Z0-9]{8}` for Patient_ID or `*@*.*` for Email), assign the corresponding Master Heading.

### 3.3 Structural Repair
- **Merged Cell Unstacking**: Detect merged cells in header rows and propagate the label to all sub-columns to prevent data misalignment.
- **Deduplication**: If structural damage results in two columns mapping to the same Master Heading, merge the columns into a single field, prioritizing non-null values and logging the conflict.

---

## 5. Strict Accuracy Guardrails (Milestone 5)
To ensure zero tolerance for data errors, all "healed" or "inferred" data must undergo a mandatory human review.

### 5.1 Mandatory Review Flag
The system must automatically tag records with a `Review_Required` flag (boolean) if any of the following conditions are met during ingestion:
- **Fuzzy Match Trigger**: A header was mapped using fuzzy matching (Levenshtein distance > 0).
- **Deep Search Trigger**: Valid headers were only found after scanning beyond Row 2 (Deep Scan).
- **Contextual Inference Trigger**: A header was inferred from data patterns (Protocol 3.2).
- **Structural Repair Trigger**: Merged cells were unstacked or duplicate columns were merged (Protocol 3.3).

### 5.2 Lead Verification Queue (Suggested Healing)
- **Suggested vs. Automatic**: Milestone 5 transitions the system to a "Suggested Healing" model. High-risk automated repairs (fuzzy matching, deep search, contextual inference) are strictly prohibited from merging until a human Lead reviews the suggestion.
- **Gating**: Records tagged with `Review_Required = True` are diverted to a persistent **Staging Queue** with status `NEEDS_REVIEW` and are **NOT** merged into the Master Database during atomic exports.
- **Verification UI**: The Lead Dashboard must provide a "Verification Queue" interface showing:
    - Original "Broken" Header vs. Healed Master Heading.
    - Data samples from the affected column.
    - Reason for the Review Flag (e.g., "Fuzzy Match: 'Ptnt ID' -> 'Patient_ID'").
- **Actions**:
    - **Approve**: Removes the flag and moves the record to the Master Database.
    - **Correct**: Allows the Lead to manually re-map the column before merging.
    - **Reject**: Purges the record from the Staging Queue.

### 5.3 Learning Loop
- Upon Lead approval, the system must permanently add the approved mapping to `aliases.json` to reduce future flags for the same variation.

---

## 6. Audit & Logging
Every sanitization, healing, and verification action must be logged in the `audit_log.jsonl` with the following attributes:
- `timestamp`: ISO-8601
- `action`: `SANITIZATION_FORMULA`, `SANITIZATION_MACRO`, `HEALING_HEADER`, `ACCURACY_GUARDRAIL`, or `LEAD_APPROVAL`
- `file_source`: Original filename
- `details`: Specific change made or trigger detected (e.g., "Healed 'Pt ID' to 'Patient_ID'", "Triggered Fuzzy Match", "Lead approved and merged")

---

## 7. International Data Privacy Compliance (Global Scaling)
To support regional multi-tenancy (Caribbean, North/South America, Africa) and ensure compliance with international standards (HIPAA, GDPR, Jamaica Data Protection Act), the following data handling rules apply:

### 7.1 PII Masking (Dashboard Display)
Sensitive data (Patient_Name) must be masked by default in the dashboard to prevent unauthorized exposure.
- **Rule**: Replace internal characters with asterisks (e.g., "John Doe" -> "J**n D*e").
- **Exemption**: Only authorized clinical staff with explicit "View PII" permissions (Master Password authorized) may view unmasked names.

### 7.2 Data Residency & Multi-Tenancy
- **Regional Tags**: Every record must carry a mandatory `Region` tag.
- **Access Control**: Future iterations will implement Regional Profile filters, ensuring data from one region is not accessible by administrators of another unless explicitly shared.
- **Encryption**: All data at rest is encrypted with AES-256 using the `processor/encryption.py` module.
