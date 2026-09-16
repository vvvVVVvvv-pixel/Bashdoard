# Bashdoard

A PowerShell utility built to automate tedious N1 support workflows, bypass manual copy-pasting, and standardize ticket escalation. 

Originally created to eliminate repetitive CRM data entry and streamline technical diagnostics (like extracting ONT serial numbers and OLT IPs for SSH/Paramiko monitoring).

---

## Features
* *CRM Data Extraction:* Connects to internal CRM APIs via Invoke-WebRequest to automatically pull customer plans, site names, and network access methods.
* *Streamlined Questionnaires:* Rapidly captures structured client contact details and failure descriptions during live calls.
* *Technical Profiling & Automation:* Detects fiber connections, extracts backend ONT/OLT parameters, and formats them into a clean string optimized for auxiliary Python/Paramiko scripts that automate Huawei OLT checks (optical power, line status, etc.).
* *Instant Ticket Formatting:* Compiles everything into a uniform layout, copies it straight to the clipboard, and logs it locally.

---

## Installation

This script runs as a standalone PowerShell utility. 

1. Ensure you have *PowerShell 5.1+* (tested on 5.1.26100.8655).
2. Clone or download the script file:
   ```powershell
   git clone [https://github.com/your-username/bashdoard.git](https://github.com/your-username/bashdoard.git)
   
# Usage
Open PowerShell and load or execute the script.

Run the main loop function to initialize the interactive terminal:

PowerShell
Run
Enter the target account number when prompted.

Refresh-Cookies
If session cookies have expired, paste the fresh credentials when requested.

Hit Ctrl + C to cancel an active process, or type Finish to reset the session state.
