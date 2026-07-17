# Handoff: Transitioning to General Medical Data Sorter (GMDS)

This document provides the roadmap for transforming the specialized **SCD Dbase Sorter** into a universal **General Medical Data Sorter (GMDS)**. Use this as the first prompt for the Agent Lead in your new business.

## 1. Project Vision
The GMDS is a high-security, disease-agnostic framework for clinical data ingestion. Its value lies in its ability to "learn" any hospital's messy Excel formatting and proactively "discover" missing patient records across local files and emails, regardless of the diagnosis.

## 2. Generalization Instructions for New Agent
*Prompt the new Agent Lead with the following:*
> "I have uploaded the codebase for a high-security medical data sorter. Your mission is to generalize it. Please perform the following:
> 1. **String Replacement**: Scan the entire project and replace 'SCD Dbase Sorter' with 'General Medical Data Sorter'. 
> 2. **Refactor `processor/mapping.py`**: Change the `MASTER_HEADINGS` to a generic clinical set: `['Patient_ID', 'Patient_Name', 'Facility', 'Record_Date', 'Diagnosis', 'Treatment', 'Validation_Status']`.
> 3. **Refactor `processor/search_bot.py`**: Change the discovery keywords from SCD-specific terms to a dynamic list that can be configured via the UI.
> 4. **Update UI**: In `app.py`, change the sidebar and headers to reflect a multi-disease clinical environment."

## 3. Key Files to Export/Import
| File | Role | Change Level |
| :--- | :--- | :--- |
| `app.py` | Main Dashboard & Portal | Medium (UI Text/Headers) |
| `processor/mapping.py` | Column Alias Engine | High (Headings/Schema) |
| `processor/search_bot.py` | Discovery Engine | Medium (Search Keywords) |
| `processor/encryption.py` | AES-256 Security | None (Keep as-is) |
| `processor/otp_service.py` | SMS Verification | None (Keep as-is) |
| `processor/sanitization.py` | Injection Defense | None (Keep as-is) |
| `requirements.txt` | Environment Setup | None (Keep as-is) |
| `Dockerfile` | Production Packaging | None (Keep as-is) |

## 4. One-Click Deployment & "Self-Hardening"
To sell to non-technical buyers, the app must be "Secure by Default."
*   **Dockerization**: Provide a `Dockerfile` so the app runs in a sealed, pre-configured environment.
*   **Setup Wizard**: Implement a first-run check in `app.py` that forces the creation of a Master Password and verifies SSL/TLS.
*   **Environment Vaults**: Use platform-specific "Secret" managers to store Twilio and SMTP keys so the user never touches raw code.

## 5. Monetization Path for GMDS
*   **Marketplace Listing**: Package as a "Universal Clinical Data Framework."
*   **White-Labeling**: Sell to research firms as a secure ingestion layer.
*   **SaaS**: Subscription-based patient registry management for small clinics.

## 6. Deployment Note
Ensure you transfer the `HARDENING_HANDOFF.md` file to the new project. It contains the essential instructions for the Hercules/Security team to move the app from a sandbox to a production-hardened environment.

---
**SCD Dbase Sorter Team**
