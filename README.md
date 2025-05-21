# 🛠️ PSScripts

A growing collection of practical, production-tested **PowerShell scripts** used in real-world **IT system administration** environments.

These scripts are written and maintained by [Benjamin Rain (TulakHordia)](https://github.com/TulakHordia) to automate routine administrative tasks, improve visibility, and streamline Windows infrastructure management.

---

## 📁 Repository Overview

This repository is organized by function. Each script includes inline documentation and many export results to CSV for reporting or audit purposes.

| Folder                 | Description |
|------------------------|-------------|
| `AD Scripts`           | Manage Active Directory users, computers, groups, and stale object cleanup. |
| `Complete Toolboxes`   | Fully interactive PowerShell UIs that combine multiple utilities in one menu-based script. |
| `Entra`                | Scripts related to Microsoft Entra (Azure AD), including Conditional Access and audit tools. |
| `Exchange`             | Scripts related to Microsoft Exchange, Mailbox confifguring and auditing. |
| `Printer`              | Scripts for querying and managing printers and printer assignments across endpoints. |
| `Sentinel`             | Scripts related to SentinelOne, log management, and security automation. |
| `SharePoint`           | Scripts related to Microsoft SharePoint, OneDrive and Cloud storage. |
| `Windows`              | Miscellaneous Windows-focused scripts for OS maintenance, version queries, device audits, and admin tools. |

> Most scripts support CSV export to: `C:\Twistech\Script Results`

---

## ✅ Common Features

- Designed for use by IT administrators and support engineers
- Organized menus for easier execution and task selection
- Includes:
  - Inactivity detection and remediation (AD)
  - Dynamic date threshold and custom tagging
  - Mailbox and calendar permissions audit (Exchange)
  - Microsoft 365 and Azure integration
  - Conditional Access filtering
  - Printer discovery and export
- Export results for documentation or compliance

---

## 🧠 Getting Started

1. Clone or download the repository:
   ```bash
   git clone https://github.com/TulakHordia/PSScripts.git
