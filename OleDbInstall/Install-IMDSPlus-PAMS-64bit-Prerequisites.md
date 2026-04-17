# OLE DB Driver 19 for SQL Server - Installation Instructions

## Overview

This process will install **Microsoft OLE DB Driver 19 for SQL Server** and its prerequisites on your Windows computer. This is required for the 64-bit version of IMDSPlus/PAMS.

---

## ⚠️ IMPORTANT: Administrator Privileges Required

**This installation requires Administrator privileges.** If you do not have administrator access on your computer, please contact your **IT Support team** for assistance.

---

## Prerequisites

Before running the installer, ensure:
- Windows 10 or later (64-bit)
- Administrator access to the computer

---

## Installation Steps

### Step 1: Run the Installer

1. Open IMDSPlus/PAMS Logon.
2. Go to **Diagnostics → Choose a Diagnostic → Install-OleDbDriver19**

The installer will automatically:
- Install Visual C++ Redistributable 2015-2022 (x86)
- Install Visual C++ Redistributable 2015-2022 (x64)
- Install Microsoft OLE DB Driver 19 for SQL Server

### Step 2: Follow On-Screen Prompts

- The script will display progress and status messages
- If existing installer processes are detected, you may be prompted to terminate them
- Wait for the installation to complete

### Step 3: Restart if Required

**A system restart may be required** after installation. If prompted, please restart your computer to complete the installation.

---

## 🔧 Manual Installation (If Script Fails)

If the automated script fails, you can manually download and install the components from Microsoft:

### Download Links:

1. **Visual C++ Redistributable (x64)**  
   https://aka.ms/vc14/vc_redist.x64.exe

2. **Visual C++ Redistributable (x86)**  
   https://aka.ms/vc14/vc_redist.x86.exe

3. **Microsoft OLE DB Driver 19 for SQL Server**  
   https://go.microsoft.com/fwlink/?linkid=2318101

### Manual Installation Order:

1. Download all three files above
2. Run `vc_redist.x86.exe` - click "Install" and wait for completion
3. Run `vc_redist.x64.exe` - click "Install" and wait for completion
4. Run `msoledbsql19.msi` - follow the setup wizard (see details below)
5. **Restart your computer**

### Installing MSOLEDBSQL19 Manually

When running the `msoledbsql19.msi` installer manually:

1. Accept the license agreement
2. On the **Feature Selection** screen:
   - ✅ Keep **"OLE DB Driver"** selected (required)
   - ❌ **UNCHECK "SDK"** - this is not needed for normal use
3. Click **Install** and wait for completion
4. Click **Finish**

> **Note:** The SDK (Software Development Kit) is only needed for developers building applications. End users should **not** install the SDK.

---

## Troubleshooting

### "Access Denied" or "Administrator Required" Error
- You must run the installer as Administrator
- **Contact your IT Support team** if you don't have admin access

### Installation Hangs or Times Out
- Another installation may be in progress
- Close any pending Windows updates or other installers
- Restart your computer and try again

### Components Not Detected After Installation
- A system restart may be required
- Restart your computer and run the script again to verify

### Still Having Issues?
- Check the log file: `Install-OleDbDriver19.log` (same folder as the script)
- Contact your **IT Support team** with the log file for assistance

---

## Contact IT Support

If you encounter any issues during installation or do not have administrator privileges, please contact your IT Support team for assistance. Provide them with:
- The error message (if any) and a screenshot
- The log file: `Install-OleDbDriver19.log`
- The diagnostic output from **View Diagnostic** (copy and paste)

Email all of this information to REJIS IMDSPlus/PAMS support.

---

## Version Information

- **OLE DB Driver Version:** 19.x
- **VC++ Redistributable:** 2015-2022 (v14.x)
- **Script Date:** April 2026
