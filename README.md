# 🖥️ Windows Inventory Audit Script

This script collects system and hardware information from Windows computers and uploads it to a secure network share. It is intended for use in inventory audits and asset tracking.

---

## ✅ Features

- Collects:
  - PC name, manufacturer, model, serial number
  - Windows edition & build
  - MAC addresses (Ethernet/Wi-Fi)
  - CPU name, RAM size/speed
  - Drive type and capacity
- Saves output as a `.txt` file named after the device serial number
- Uploads to a selected subfolder on a secure network share
- Option to compile into a standalone `.exe`

---

## 📁 Output Location

```text
\\sharecrypt.oakland.edu\SAT Secure\Inventory\2025 Inventory
```

## Usage 

1. Run the script with Python:
   python inventory_audit.py

2. Enter your admnet username and password when prompted

3. Select the destination folder from the list provided

4. Script will copy the file to the selected folder and remove the local copy

## Executable:

To create an executable version of this script:
pyinstaller --onefile inventory_audit.py

## Destination Path:

\\sharecrypt.oakland.edu\SAT Secure\Inventory\2025 Inventory

## Requirements:

- Windows OS
- Python 3.x
- Modules: psutil, wmi, pywin32

## Note:

Ensure the user running the script has permission to access the network share.

## Author:

Oakland University STC
