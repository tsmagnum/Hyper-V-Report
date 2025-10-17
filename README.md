
# Hyper-V Report Script Documentation

## Overview

This PowerShell script automates the collection of detailed information about a Hyper-V environment, including hosts, virtual machines, snapshots, replication status, virtual hard disks (VHDX), network adapters, and virtual switches. It generates an HTML report and optionally sends it via email.

---

## Features

- Collects host system information
- Enumerates all virtual machines and their configurations
- Lists VM snapshots and calculates their age
- Reports replication status for VMs
- Gathers VHDX file details
- Extracts VM and management OS network adapter data
- Lists virtual switch configurations
- Generates a comprehensive HTML report
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
  - `StyleCSS.ps1`
  - `HtmlCode.ps1`
  - `Functions.ps1`

---

## Script Parameters

These are expected to be defined in `GlobalVariables.ps1`:

- `$reportHtmlDir` – Directory to save the HTML report
- `$reportHtmlName` – Base name for the report file
- `$clusterDeployment` – Boolean flag for cluster support
- `$replicationInfoNeeded` – Boolean flag to include replication info
- `$vhdxInfoNeeded` – Boolean flag to include VHDX info
- `$vmnetInfoNeeded` – Boolean flag to include VM network adapter info
- `$osNetInfoNeeded` – Boolean flag to include management OS network adapter info
- `$vswitchInfoNeeded` – Boolean flag to include virtual switch info
- `$reportHtmlRequired` – Boolean flag to generate HTML report
- `$emailReport` – Boolean flag to send report via email
- `$emailSystem` – Email system to use (`msgraph` or `mailkit`)

---

## Output

- **HTML Report**: Saved in `$reportHtmlDir` with timestamped filename.
- **Console Output**: Displays formatted tables for each section.
- **Email**: Sent if `$emailReport` is enabled and `$reportHtmlRequired` is true.

---

## Usage

```powershell
.\Hyper-V-Report.ps1
```

Ensure all required variables and modules are properly configured before execution.

---

## Sections in the Report

1. **Host Info** – CPU, RAM, OS version, VHD volume stats
2. **VM Info** – Name, generation, memory, IP, state, uptime, replication
3. **Snapshots** – Snapshot name, age, parent snapshot
4. **Replication** – Status, health, last replication time
5. **VHDX Info** – Format, type, size, fragmentation
6. **VM Network Adapters** – MAC, IP, vSwitch, VLAN
7. **Management OS Adapters** – IP, MAC, vSwitch, VLAN
8. **Virtual Switches** – Name, type, uplinks, SET status

---

## Notes

- Cluster support is marked as "Coming soon".
- Email system must be explicitly selected (`msgraph` or `mailkit`).