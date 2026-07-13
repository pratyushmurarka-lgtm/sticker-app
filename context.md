# Project Context: Intech QR Code Repository System

## 1. GitHub Repository Information
* **Repository URL**: `https://github.com/pratyushmurarka-lgtm/sticker-app.git`
* **Default Branch**: `main`
* **Push/Pull Guideline**: All code modifications must be committed and pushed to this repository so they can be easily shared between the Development PC and the Production Server PC.

## 2. Directory Paths
* **Development PC Workspace**: `C:\Users\intec\.gemini\antigravity\scratch\QRCode_Project`
* **Production Server PC Workspace**: `C:\QRCode_Project` on server `192.168.1.218`

## 3. Database Credentials (C:\dbconfig.txt)
The application uses the connection settings specified in `C:\dbconfig.txt`:
* **Development (TEST)**:
  * Server: `ERPINTECH\SQLEXPRESS`
  * Database: `Es_Comp0001_2026`
  * User: `sa` / Password: `12345678`
* **Production (LIVE)**:
  * Server: `localhost` (when running on 192.168.1.218)
  * Database: `Es_Comp0001_2026`
  * User: `sa` / Password: `Intechipl@12345`

## 4. Key Systems & Architectures
* **Web App (app.py)**: Hosted on port `5000` (on production server `192.168.1.218:5000`).
* **Print Agent (print_agent.py)**: Runs on local client PCs (port `5011` / loopback) physically connected to the label printers (TSC/Zebra). Handles slow GDI print spooling asynchronously in background threads.
* **Granular Permissions Checkboxes**: Individual features (Generate, Import, Reprint, Reports, Manage Users) can be turned on/off per user inside the User Management tab.
* **Auto-Fallback ODBC Driver**: `database.py` dynamically loops and tests available ODBC drivers (18, 17, 13, Native Client, legacy) and catches failures to bypass registry corruption bugs.
* **Sequential Generation Speedup**: Duplicate queries were removed from the generation loop (100x speedup).
* **Android-Style Dropdown Search**: Custom dropdown in sequential tab allows real-time fuzzy matching by typing.

---

## 5. Deployment/Pull Commands for Server Agent
If running an agent on the Server PC, use the following commands to pull the latest code:
```powershell
# Navigate to workspace
cd C:\QRCode_Project

# Pull latest files from Git
git pull origin main

# Install any new dependencies
"C:\Users\Dell\AppData\Local\Python\bin\python.exe" -m pip install -r requirements.txt
```
