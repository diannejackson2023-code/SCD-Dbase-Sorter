# SCD Companion Scanner - Superbot v1

The Companion Scanner is a standalone Python application designed for recipients to securely scan their local drives for Sickle Cell Disease (SCD) records.

## Features
- **Local Scanning**: Automatically scans Desktop, Documents, and Downloads for `.xlsx`, `.xls`, and `.docx` files.
- **Security Sanitization**: Applies the team's sanitization protocol to prevent Excel formula injection and XSS.
- **Malware Scanning**: Conceptual hook for integrating antivirus checks before upload.
- **Secure Streaming**: Encrypts and streams discovered data directly to the Discovery Staging API.

## Usage
1. Ensure Python 3.10+ is installed.
2. Install dependencies:
   ```bash
   pip install requests pandas openpyxl
   ```
3. Run the scanner with your secure discovery token:
   ```bash
   python scanner.py <YOUR_TOKEN>
   ```

## Development
To compile into a single-file executable for Windows/macOS:
```bash
pip install pyinstaller
pyinstaller --onefile scanner.py
```
