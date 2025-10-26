
# Hyper-V Report Script Documentation

## Overview

This PowerShell script automates the collection of detailed information about a Hyper-V environment, including hosts, virtual machines, snapshots, replication status, VHDX files, network adapters, virtual switches, and—new in this version—CSV health and space utilization, cluster configuration, and cluster network details. It generates a customizable HTML report and can optionally send it via email using MS Graph or MailKit.

The script supports both standalone and clustered deployments.

> ⚠️ Note: The script installs required modules (MS Graph or MailKit) if not already present. Avoid running it if module installation could impact your environment.

## Features

- Collects host system information
- Enumerates all virtual machines and their configurations
- Lists VM snapshots and calculates their age
- Reports replication status for VMs
- Gathers VHDX file details
- Extracts VM and management OS network adapter data
- Lists virtual switch configurations
- NEW: Reports CSV health and space utilization (clustered only)
- NEW: Includes cluster configuration and network details
- Generates a comprehensive HTML report with selectable styles
- Sends the report via email using MS Graph or MailKit

## Prerequisites

- PowerShell 5.1 or later
- Hyper-V role installed
- Required modules:
  - Hyper-V
  - CimCmdlets
- External scripts in the same directory:
  - GlobalVariables.ps1
  - StyleCSS.ps1 (or variants like StyleCSS-Pro.ps1, StyleCSS-Colorful.ps1)
  - HtmlCode.ps1
  - Functions.ps1

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

## Output

- **HTML Report**: Saved in `$reportHtmlDir` with a timestamped filename.
- **Console Output**: Displays formatted tables for each section.
- **Email**: Sent if `$emailReport` is enabled and `$reportHtmlRequired` is true.

## Usage

```powershell
.\Hyper-V-Report.ps1
```

Ensure all required variables and modules are properly configured before execution.

## Sections in the Report

1. **Host Info** – CPU, RAM, OS version, VHD volume stats
2. **CSV Health Info** – CSV status and health (clustered only)
3. **CSV Space Utilization** – CSV volume usage (clustered only)
4. **VM Info** – Name, generation, memory, IP, state, uptime, replication
5. **Snapshots** – Snapshot name, age, parent snapshot
6. **Replication** – Status, health, last replication time
7. **VHDX Info** – Format, type, size, fragmentation
8. **VM Network Adapters** – MAC, IP, vSwitch, VLAN
9. **Management OS Adapters** – IP, MAC, vSwitch, VLAN
10. **Virtual Switches** – Name, type, uplinks, SET status
11. **Cluster Configuration** – Cluster settings and roles
12. **Cluster Networks** – Cluster network topology and status

## Notes

- Cluster support is now implemented.
- Email system must be explicitly selected (`msgraph` or `mailkit`).
- Only the user who encrypted SMTP credentials can decrypt them.
