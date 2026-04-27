
# IT Admin Guide to IMDSPlus-PAMS 26.1.0.0 (64-bit) Prerequisites


## Required Microsoft Runtime Components
This application depends on several Microsoft‑provided runtime components that must be installed on the target workstation before launch. These components are standard Microsoft redistributables commonly used across enterprise Windows environments:

### 1. Visual C++ Redistributable (x86 – 32-bit)

**What it is:** Runtime libraries for 32-bit applications  
**What it's for:** Although Microsoft OLE DB Driver 19 (x64) installs both 64-bit and 32-bit driver binaries, the x64 installer of the C++ Redistributable does not install 32-bit runtime libraries. The x86 C++ Redistributable must be installed separately to provide required 32-bit runtime dependencies that the x64 OLE DB Driver needs.  
**File:** `vc_redist.x86.exe` (~13 MB)  
**Reference:** [Microsoft Learn - Download OLE DB Driver](https://learn.microsoft.com/en-us/sql/connect/oledb/download-oledb-driver-for-sql-server?view=sql-server-ver17)

### 2. Visual C++ Redistributable (x64 – 64-bit)

**What it is:** Runtime libraries for 64-bit applications  
**What it's for:** Provides the 64-bit runtime environment required by MS OLE DB Driver.
**File:** `vc_redist.x64.exe` (~14 MB)

### 3. Microsoft OLE DB Driver 19 for SQL Server (x64)

**What it is:** Database connectivity driver  
**What it's for:** Installs both 64-bit and 32-bit MS OLE DB driver components to enable IMDSPlus-PAMS 64-bit to establish secure connections to SQL Server databases for data retrieval and transactions  
**File:** `msoledbsql19.msi` (~64 MB)

These components must be installed prior to running the application to ensure proper database connectivity and runtime stability.

---

## ⚠️ IMPORTANT: Administrator Privileges Required

**This installation requires Administrator privileges.** 

---

## Prerequisites

Before running the installer, ensure:
- Windows 10 or later (64-bit)
- Administrator access to the computer

---

## 🔧 Manual Installation 

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


---

## Version Information

- **VC++ Redistributable:** 2015-2022 (v14.5)
- **OLE DB Driver Version:** 19.4


