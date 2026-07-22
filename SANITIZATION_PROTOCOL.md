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

## 4. Audit & Logging
Every sanitization and healing action must be logged in the `audit_log.jsonl` with the following attributes:
- `timestamp`: ISO-8601
- `action`: `SANITIZATION_FORMULA`, `SANITIZATION_MACRO`, or `HEALING_HEADER`
- `file_source`: Original filename
- `details`: Specific change made (e.g., "Prepended ' to cell A4", "Stripped VBA macros", "Healed 'Pt ID' to 'Patient_ID'")
