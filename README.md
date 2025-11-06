
# Hyper-V Report Script Documentation

## Overview

This PowerShell script automates the collection of detailed information about a Hyper-V environment, including hosts, virtual machines, snapshots, replication status, VHDX files, network adapters, virtual switches, and cluster-specific metrics. It generates a customizable HTML report and can optionally send it via email using MS Graph or MailKit.

The script supports both standalone and clustered deployments.

> ⚠️ Note: The script installs required modules (MS Graph or MailKit) if not already present. Avoid running it if module installation could impact your environment.
> 
> 🚨 Important: This script must be run in an elevated PowerShell session by a user with administrator rights on the Hyper-V server.

---

## Features

- Collects host system information
- Enumerates all virtual machines and their configurations
- Lists VM snapshots and calculates their age
- Reports replication status for VMs
- Gathers VHDX file details
- Extracts VM and management OS network adapter data
- Lists virtual switch configurations
- Reports CSV health and space utilization (clustered only)
- Includes cluster configuration and network details
- Generates a comprehensive HTML report with selectable styles
- Sends the report via email using MS Graph or MailKit

---

## Prerequisites

- PowerShell 5.1 or later
- Hyper-V role installed
- Required modules:
  - `Hyper-V`
  - `CimCmdlets`
- External scripts in the same directory:
  - `GlobalVariables.ps1`
  - `HtmlCode.ps1`
  - `Functions.ps1`
  - CSS styles in the `Style` subfolder:
    - `StyleCSS-Minimal.ps1`
    - `StyleCSS-Pro.ps1`
    - `StyleCSS-ProDark.ps1`
    - `StyleCSS-Colorful.ps1`

---

## Script Parameters

Defined in `GlobalVariables.ps1`:

- `$reportHtmlDir` – Directory to save the HTML report
- `$reportHtmlName` – Base name for the report file
- `$reportStyle` – Style of the HTML report (`minimal`, `pro`, `prodark`, `colorful`, `professional`)
- `$clusterDeployment` – Boolean flag for cluster support
- `$csvHealthInfoNeeded` – Include CSV health info (clustered only)
- `$csvSpaceInfoNeeded` – Include CSV space utilization (clustered only)
- `$clusterConfigInfoNeeded` – Include cluster configuration details
- `$clusterNetworksInfoNeeded` – Include cluster network details
- `$replicationInfoNeeded` – Include replication info
- `$vhdxInfoNeeded` – Include VHDX info
- `$vmnetInfoNeeded` – Include VM network adapter info
- `$osNetInfoNeeded` – Include management OS network adapter info
- `$vswitchInfoNeeded` – Include virtual switch info
- `$reportHtmlRequired` – Generate HTML report
- `$emailReport` – Send report via email
- `$emailSystem` – Email system to use (`msgraph` or `mailkit`)
- `$encryptedSMTPCredsFileName` – Filename for encrypted SMTP credentials

---

## Output

- **HTML Report**: Saved in `$reportHtmlDir` with a timestamped filename.
- **Console Output**: Displays formatted tables for each section.
- **Email**: Sent if `$emailReport` is enabled and `$reportHtmlRequired` is true.

---

## Usage

```powershell
 'powershell.exe -ExecutionPolicy Bypass -File .\Hyper-V-Report.ps1'
```

Ensure all required variables and modules are properly configured before execution.

---

## Sections in the Report

1. Host Info – CPU, RAM, OS version, VHD volume stats
2. CSV Health Info – CSV status and health (clustered only)
3. CSV Space Utilization – CSV volume usage (clustered only)
4. VM Info – Name, generation, memory, IP, state, uptime, replication
5. Snapshots – Snapshot name, age, parent snapshot
6. Replication – Status, health, last replication time
7. VHDX Info – Format, type, size, fragmentation
8. VM Network Adapters – MAC, IP, vSwitch, VLAN
9. Management OS Adapters – IP, MAC, vSwitch, VLAN
10. Virtual Switches – Name, type, uplinks, SET status
11. Cluster Configuration – Cluster settings and roles
12. Cluster Networks – Cluster network topology and status

---

## Support Script: Save-SafeCreds.ps1

To securely store SMTP credentials for email delivery, use the `Save-SafeCreds.ps1` script:

### Usage

```powershell
'powershell.exe -ExecutionPolicy Bypass -File .\Save-SafeCreds.ps1'
```

This script:
- Prompts for SMTP username and password
- Encrypts the credentials using the current user's context
- Saves them to an XML file (e.g., `EncryptedCreds.xml`)
- Ensures only the user who created the file can decrypt it

Make sure the filename matches the value of `$encryptedSMTPCredsFileName` in `GlobalVariables.ps1`.

---


# EntraAuthAnalyzer

**EntraAuthAnalyzer** is a PowerShell-based tool designed to audit and analyze authentication methods, Conditional Access (CA) policies, and login behaviors in Microsoft Entra ID (formerly Azure AD). It generates detailed reports and optionally exports them to Excel or sends them via email.

## Features
- Authentication Methods Audit
- Security Info Registrations
- CA Policy Exclusions
- CA Policy Changes
- Login Analysis
- Excel Export
- Email Report

## Configuration
Edit the `GlobalVariables.ps1` file to configure the script for your environment.

### Authentication
If you want to run the script unattended (e.g. using task scheduler), please
see `AppOnly-Auth.txt` and set the following values:

```powershell
$tenantId = "your-tenant-id"
$clientId = "your-client-id"
$certificateThumbprint = "your-cert-thumbprint"
$runUnattended = $true
```

### Report Options
Infos to include in the report: 

```powershell
$entraAuthMethodsReq = $true
$entraSecInfoRegReq = $true
$entraCaExclusionsReq = $true
$expandGroups = $true
$entraCaExclusionsData = "list"
$caPolicyChangesReq = $true
$entraLoginsNoCaReq = $true
$entraLoginsNoCaConsoleLimit = 100
```

### Observation Windows
```powershell
$loginsLookupPeriod = 30
$loginsLogsLookupPeriod = 3
```

### Excel Export
```powershell
$exportToExcel = $true
$tenantName = "MyTenant"
$resultsFileDir = "C:\Temp"
$resultsFileName = "Entra_Auth_Analyzer"
$resultsFileExt = ".xlsx"
```

### Email Delivery
```powershell
$sendEmail = $false
$emailSender = "mySender@domain.com"
$emailRecipient = "myrecipient@domain.com"
$emailCcrecipient = $null
$subject = "Entra Id Auth Report - $($tenantName)"
$smtpServer = "mySmtp.server.com"
$smtpServerPort = 587
$smtpAuthRequired = $true
```

Use encrypted SMTP credentials:
```powershell
$encryptedSMTPCreds = $true
$encryptedSMTPCredsFileName = "EncryptedCreds.xml"
```

Or use plain text credentials (not recommended):
```powershell
$smtpServerUser = "smtpserver.user"
$smtpServerPwd = "mySecretPwd"
```

## Usage
```powershell
git clone https://github.com/tsmagnum/EntraAuthAnalyzer.git
cd EntraAuthAnalyzer
.\EntraAuthAnalyzer.ps1
```

## Output
- Excel Workbook with multiple sheets:
  - AuthMethods
  - SecInfoReg
  - CaExclusions
  - CaPolCh
  - Logins_NoCA
- Email Report (if enabled)

## Notes
- The Excel file always contains all login records, regardless of console display limits.
- Use encrypted XML for SMTP credentials to enhance security.
- Ensure proper permissions and compliance with your organization's policies.


## 🔧 Variable Explanations
The following variables are defined in `GlobalVariables.ps1` and control script behavior:

- **$tenantId**: Azure AD tenant ID used for app-only authentication.
- **$clientId**: Client ID of the registered application in Entra ID.
- **$certificateThumbprint**: Thumbprint of the certificate used for authentication.
- **$runUnattended**: If set to $true, runs the script without user interaction.
- **$entraAuthMethodsReq**: Include authentication methods section in the report.
- **$entraSecInfoRegReq**: Include security info registration status in the report.
- **$entraCaExclusionsReq**: Include users excluded from Conditional Access policies.
- **$caPolicyChangesReq**: Include recent changes to Conditional Access policies.
- **$entraLoginsNoCaReq**: Include successful logins not covered by Conditional Access.
- **$exportToExcel**: Enable export of results to an Excel file.
- **$resultsFileDir**: Directory path where the Excel report will be saved.
- **$resultsFileName**: Base name of the Excel report file.
- **$resultsFileExt**: File extension for the report (usually .xlsx).
- **$sendEmail**: Enable sending the report via email.
- **$emailSender**: Email address used as sender.
- **$emailRecipient**: Recipient email address.
- **$smtpServer**: SMTP server used to send the email.
- **$smtpServerPort**: Port number for the SMTP server.
- **$encryptedSMTPCreds**: If true, uses encrypted credentials for SMTP.
- **$encryptedSMTPCredsFileName**: Filename of the encrypted SMTP credentials XML file.

---

## Support Script: Save-SafeCreds.ps1

To securely store SMTP credentials for email delivery, use the `Save-SafeCreds.ps1` script:

### Usage

```powershell
'powershell.exe -ExecutionPolicy Bypass -File .\Save-SafeCreds.ps1'
```

This script:
- Prompts for SMTP username and password
- Encrypts the credentials using the current user's context
- Saves them to an XML file (e.g., `EncryptedCreds.xml`)
- Ensures only the user who created the file can decrypt it

Make sure the filename matches the value of `$encryptedSMTPCredsFileName` in `GlobalVariables.ps1`.

---

## ⚠️ Disclaimer

These scripts are provided **"as-is"**, without any express or implied warranties. The author assumes **no responsibility or liability** for any errors, issues, or damages that may arise from its use. By running these scripts, you acknowledge that you do so **at your own risk** and that you are solely responsible for ensuring it complies with your organization's policies and security standards.
