# 🚀 Automation Traceroute Report with PowerShell code and export it in Excel Sheet v6

> **Document Type:** IT Operations Guide
> 
> **Category:** Network Diagnostics & Automation
> 
> **Platform:** Windows PowerShell 5.1+
> 
> **Version:** 6.0
> 
> **Status:** ✅ Production Ready
> 

---

## 📑 Table of Contents

* 📘 Overview
* 🎯 Purpose
* 🏢 Business Use Case
* ⭐ Key Features
* 🛠️ Prerequisites
* 📦 Dependencies
* 🌐 Network Detection Logic
* 📡 Target Services
* ⚙️ Script Workflow
* 🧠 Script Architecture
* 🚀 Execution Guide
* 📊 Report Structure
* 📁 Output
* 📄 Sample Report
* 💻 Complete Script
* 🔍 Code Walkthrough
* 🔒 Security Considerations
* ⚡ Performance
* 🧪 Testing
* 🐞 Troubleshooting
* 📝 Operational Notes
* 🔄 Maintenance
* 📚 Appendix

---

# 📘 Overview

This PowerShell automation script performs **parallel traceroute diagnostics** against multiple Microsoft 365 and Google Workspace services and exports the results to a structured Microsoft Excel workbook.

The solution is designed to help IT administrators quickly collect consistent network diagnostic information while minimizing manual effort and improving incident response times.

> [!NOTE]
> The script automatically detects the workstation's network, gathers system information, executes traceroutes in parallel, and generates a multi-sheet Excel report.

---

# 🎯 Purpose

The primary objectives of this automation are to:

* ✅ Standardize traceroute collection across workstations.
* ✅ Reduce manual troubleshooting time.
* ✅ Improve consistency of diagnostic reports.
* ✅ Simplify escalation to network providers and vendors.
* ✅ Maintain historical network diagnostic records.

---

# 🏢 Business Use Case

This script is particularly useful for:

* 🌐 Network Operations Center (NOC)
* 🖥️ Desktop Support Teams
* ☁️ Microsoft 365 Administrators
* 📡 Network Engineers
* 🛠️ IT Help Desk
* 📈 Infrastructure Operations

---

# ⭐ Key Features

| Feature                           | Description                                                                             |
| --------------------------------- | --------------------------------------------------------------------------------------- |
| ⚡ Parallel Processing             | Executes all traceroutes simultaneously using PowerShell Jobs.                          |
| 🌐 Automatic Network Detection    | Maps IPv4 ranges to predefined corporate network names.                                 |
| 📊 Excel Reporting                | Creates a formatted workbook with one worksheet per destination.                        |
| 📦 Dependency Management          | Automatically installs required PowerShell modules when missing.                        |
| 🖥️ System Information Collection | Captures workstation, user, IP address, operating system, and execution timestamp.      |
| 🧹 Automatic Cleanup              | Removes completed background jobs and replaces existing reports with the same filename. |

---

# 🛠️ Prerequisites

| Requirement                           | Required |
| ------------------------------------- | -------- |
| Windows 10 / 11                       | ✅        |
| Windows PowerShell 5.1 or later       | ✅        |
| Internet Connectivity                 | ✅        |
| PowerShell Gallery Access (first run) | ✅        |
| Write Access to Desktop               | ✅        |

> [!TIP]
> Microsoft Excel is **not required**. The script generates a standard `.xlsx` file using the `ImportExcel` PowerShell module.

---

# 📦 Dependencies

The script validates the following components before execution:

* **NuGet Package Provider**
* **ImportExcel PowerShell Module**

If these components are unavailable, the script installs them automatically.

> [!IMPORTANT]
> Internet access is required during the initial execution if dependencies are not already installed.

---

# 🌐 Network Detection Logic

The workstation IPv4 address is evaluated and mapped to a corporate network identifier.

| IPv4 Range      | Network Name           |
| --------------- | ---------------------- |
| `192.168.10.x`  | 🏢 Corporate_Office    |
| `192.168.11.x`  | 👔 Corporate_Executive |
| `192.168.12.x`  | 👥 Corporate_Guest     |
| `192.168.13.x`  | 📶 Corporate_WIFI      |
| Any Other Range | ❓ Unknown_Network      |

> [!NOTE]
> This mapping is customizable to align with your organization's IP addressing scheme.

---

# 📡 Target Services

The following cloud services are included in the diagnostic process:

| Worksheet     | Target Host                |
| ------------- | -------------------------- |
| ☁️ Drive      | `drive.google.com`         |
| 🎥 Meet       | `meet.google.com`          |
| 📧 Outlook    | `outlook.office365.com`    |
| 🪟 Microsoft  | `www.microsoft.com`        |
| 📂 SharePoint | `microsoft.sharepoint.com` |
| ☁️ OneDrive   | `onedrive.live.com`        |

---

# ⚙️ Script Workflow

```text
Start
 │
 ▼
Detect Desktop
 │
 ▼
Retrieve IPv4 Address
 │
 ▼
Determine Corporate Network
 │
 ▼
Validate Dependencies
 │
 ▼
Install Missing Components (if required)
 │
 ▼
Launch Parallel Traceroutes
 │
 ▼
Collect Results
 │
 ▼
Generate Excel Workbook
 │
 ▼
Create Individual Worksheets
 │
 ▼
Write Device Information
 │
 ▼
Write Traceroute Results
 │
 ▼
Save Workbook
 │
 ▼
Remove Background Jobs
 │
 ▼
Complete
```

---

# 🧠 Script Architecture

```text
PowerShell Script
        │
        ├── Environment Detection
        ├── Network Identification
        ├── Dependency Validation
        ├── Parallel Job Processing
        ├── Traceroute Collection
        ├── Excel Workbook Creation
        └── Report Export
```

---

# 🚀 Execution Guide

1. Open **Windows PowerShell**.
2. Navigate to the directory containing the script.
3. Run the script:

```powershell
.\Automation_Tracert_Report_v6.ps1
```

4. Wait for all traceroutes to complete.
5. Open the generated Excel workbook from your Desktop.

---

# 📊 Report Structure

Each worksheet contains:

| Property             | Description                 |
| -------------------- | --------------------------- |
| 🖥️ Computer Name    | Local workstation hostname  |
| 👤 Logged-In User    | Current Windows account     |
| 🌐 IPv4 Address      | Active IPv4 address         |
| 📶 Wi-Fi SSID        | Detected network name       |
| 💻 Operating System  | Windows edition             |
| 📅 Date              | Execution date              |
| 🕒 Time              | Execution time              |
| 📍 Traceroute Output | Complete diagnostic results |

---

# 📁 Output

Example workbook:

```text
Automation_Tracert_Corporate_Office_Report_July_05_2026.xlsx
```

Workbook layout:

```text
Automation_Tracert_Corporate_Office_Report_July_05_2026.xlsx
│
├── Drive
├── Meet
├── Outlook
├── Microsoft
├── SharePoint
└── OneDrive
```

---

# 📄 Sample Report

| Property         | Example               |
| ---------------- | --------------------- |
| Computer Name    | LAPTOP-001            |
| Logged-In User   | CORP\User01           |
| IPv4 Address     | 192.168.10.25         |
| Wi-Fi SSID       | Corporate_Office      |
| Operating System | Windows 11 Enterprise |
| Date             | 2026-07-05            |
| Time             | 14:25:30              |

---

# 🔍 Code Walkthrough

Each major section of the script is responsible for a specific task:

| Section               | Responsibility                                                  |
| --------------------- | --------------------------------------------------------------- |
| Environment Detection | Identifies the Desktop location and retrieves the IPv4 address. |
| Network Mapping       | Maps IP ranges to predefined network names.                     |
| Dependency Validation | Confirms that required PowerShell components are available.     |
| Parallel Processing   | Executes multiple traceroutes concurrently.                     |
| Report Generation     | Creates a structured Excel workbook and populates worksheets.   |
| Cleanup               | Removes completed background jobs to free system resources.     |

---

# 🔒 Security Considerations

> [!WARNING]
> The script requires internet connectivity to perform traceroutes and may install PowerShell modules during the initial execution.

Recommendations:

* Execute the script from trusted sources only.
* Restrict modifications to authorized IT personnel.
* Review module installation policies in environments with restricted PowerShell execution.

---

# ⚡ Performance

Key optimization features include:

* Parallel background jobs to reduce execution time.
* Automatic resource cleanup after completion.
* Efficient worksheet generation using the `ImportExcel` module.

---

# 🧪 Testing

Before deploying in production:

* Verify connectivity to all destination hosts.
* Confirm Excel workbook creation.
* Validate worksheet contents.
* Review generated traceroute output for accuracy.

---

# 🐞 Troubleshooting

| Issue                       | Possible Cause          | Recommended Action                                                   |
| --------------------------- | ----------------------- | -------------------------------------------------------------------- |
| `ImportExcel` not found     | Module missing          | Allow the script to install dependencies automatically.              |
| Empty report                | Traceroute failed       | Verify internet connectivity and DNS resolution.                     |
| `Unknown_Network` displayed | IP range not configured | Update the IP-to-network mapping logic.                              |
| Existing report replaced    | Same filename detected  | Expected behavior; the previous report is intentionally overwritten. |

---

# 📝 Operational Notes

* Existing reports with the same filename are replaced automatically.
* Each destination is exported to a dedicated worksheet.
* The workbook remains compatible with Microsoft Excel and other spreadsheet applications supporting `.xlsx`.

---

# 🔄 Maintenance

To maintain this script:

* Update destination hosts as required.
* Review IP range mappings after network changes.
* Periodically validate compatibility with the latest PowerShell and module versions.

---

# 📚 Appendix

## Output File Naming Convention

```text
Automation_Tracert_<NetworkName>_Report_<Month_Day_Year>.xlsx
```

Example:

```text
Automation_Tracert_Corporate_Office_Report_July_05_2026.xlsx
```

> [!TIP]
> The **Complete PowerShell Script** section can be placed below this documentation without modification, ensuring the source code remains identical to the original.


# ⚙️Powershell Code

```text
<#
Automation_Tracert_Report_OptionA_Fast_v6.ps1
Parallel traceroute monitoring with SSID naming by IP range.
#>
$ErrorActionPreference='Continue'
$Desktop=[Environment]::GetFolderPath('Desktop')
try {
 $IPv4=(Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -notlike '169.254*' -and $_.IPAddress -ne '127.0.0.1'} | Select-Object -First 1 -ExpandProperty IPAddress)
} catch {$IPv4='Unavailable'}
$WiFiSSID='Unknown_Network'
if($IPv4 -like '192.168.10.*'){ $WiFiSSID='Corporate_Office' }
elseif($IPv4 -like '192.168.11.*'){ $WiFiSSID='Corporate_Executive' }
elseif($IPv4 -like '192.168.12.*'){ $WiFiSSID='Corporate_Guest' }
elseif($IPv4 -like '192.168.13.*'){ $WiFiSSID='Corporate_WIFI' }
$DateStamp=Get-Date -Format 'MMMM_dd_yyyy'
$ExcelFile=Join-Path $Desktop "Automation_Tracert_${WiFiSSID}_Report_$DateStamp.xlsx"
if(Test-Path $ExcelFile){Remove-Item $ExcelFile -Force -ErrorAction SilentlyContinue}
if(-not (Get-Module -ListAvailable -Name ImportExcel)){
 if(-not(Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)){Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null}
 Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
}
Import-Module ImportExcel -Force
try{$OS=(Get-CimInstance Win32_OperatingSystem).Caption}catch{$OS='Unavailable'}
$Targets=@(
@{Sheet='Drive';Host='drive.google.com'},
@{Sheet='Meet';Host='meet.google.com'},
@{Sheet='Outlook';Host='outlook.office365.com'},
@{Sheet='Microsoft';Host='www.microsoft.com'},
@{Sheet='SharePoint';Host='microsoft.sharepoint.com'},
@{Sheet='OneDrive';Host='onedrive.live.com'})
foreach($t in $Targets){Start-Job -Name $t.Sheet -ArgumentList $t.Host -ScriptBlock {param($h) tracert -d -h 20 -w 1000 $h 2>&1} | Out-Null}
Get-Job | Wait-Job | Out-Null
foreach($t in $Targets){
 $rows=@()
 $rows += [pscustomobject]@{Property='Computer Name';Value=$env:COMPUTERNAME}
 $rows += [pscustomobject]@{Property='Logged In User';Value="$env:USERDOMAIN\$env:USERNAME"}
 $rows += [pscustomobject]@{Property='IPv4 Address';Value=$IPv4}
 $rows += [pscustomobject]@{Property='WiFi SSID';Value=$WiFiSSID}
 $rows += [pscustomobject]@{Property='Operating System';Value=$OS}
 $rows += [pscustomobject]@{Property='Date';Value=(Get-Date -Format 'yyyy-MM-dd')}
 $rows += [pscustomobject]@{Property='Time';Value=(Get-Date -Format 'HH:mm:ss')}
 foreach($line in (Receive-Job -Name $t.Sheet)){ $rows += [pscustomobject]@{Property='Traceroute Output';Value=$line.ToString()} }
 $rows | Export-Excel -Path $ExcelFile -WorksheetName $t.Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
}
Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
Write-Host "Report saved to: $ExcelFile"
Read-Host 'Press Enter to close'

```

---

# 🚀 How to Use the Script
---

## 🛠️ Step-by-Step Execution Guide

### 1. 📁 Save the Script

Save the PowerShell script file as:

```text
Traceroute Report.ps1
```

Ensure the file is stored in a known directory (e.g., Desktop or Documents).

---

### 2. ▶️ Run the Script

You can execute the script using one of the following methods:

#### Option A — Right-click Execution (GUI Method)

* Right-click on `Traceroute Report.ps1`
* Select **“Run with PowerShell”**
* The PowerShell window will open and start execution automatically

---

#### Option B — PowerShell Terminal (Recommended)

1. Open **PowerShell**
2. Navigate to the script location:

```powershell
cd $env:USERPROFILE\Desktop
```

3. Run the script:

```powershell
.\Traceroute Report.ps1
```

---

## ⏳ 3. Wait for Execution to Complete

During execution:

* The script will run multiple **traceroute operations in parallel**
* System and network details will be collected automatically
* No user interaction is required during processing

> [!NOTE]
> Execution time may vary depending on network latency and system performance.

---

## 📊 4. Successful Completion Output

Once the script finishes successfully:

* An **Excel (.xlsx) report** will be generated automatically
* The file will be saved on the **Desktop**

Example output:

```text
Automation_Tracert_Corporate_Office_Report_July_05_2026.xlsx
```

---

## 📁 5. View the Report

* Navigate to your **Desktop**
* Open the generated `.xlsx` file using:

  * Microsoft Excel
  * LibreOffice Calc (optional)
  * Any compatible spreadsheet application

---

## 📄 Report Contents

The Excel file includes:

* 🖥️ Computer Name
* 👤 Logged-in User
* 🌐 IP Address
* 📶 Wi-Fi / Network Name
* 💻 Operating System
* 📅 Date & Time
* 📡 Full Traceroute Results (per service)

---

## ⚠️ Important Notes

> [!WARNING]
> Do not close the PowerShell window during execution, as this will interrupt traceroute collection.

> [!TIP]
> If execution fails, try running PowerShell as **Administrator** for better network access permissions.

---

## 🧾 Summary

In simple terms:

* Save script → Run it → Wait → Get Excel report on Desktop

---


