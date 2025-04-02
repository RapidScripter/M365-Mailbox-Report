# Microsoft 365 Mailbox Size Report PowerShell Script

![PowerShell](https://img.shields.io/badge/PowerShell-%235391FE.svg?style=for-the-badge&logo=powershell&logoColor=white)
![Microsoft Exchange](https://img.shields.io/badge/Microsoft_Exchange-0078D4?style=for-the-badge&logo=microsoft-exchange&logoColor=white)

This PowerShell script generates detailed reports of Microsoft 365 mailbox sizes, quotas, and usage statistics.

## Features

- 📊 Comprehensive mailbox size reporting
- 🔍 Multiple filtering options (user/shared mailboxes)
- 🔐 Supports MFA, device code, and basic authentication
- 📁 Input from file or all mailboxes in tenant
- 📈 Progress tracking and error handling
- 📂 CSV export with automatic file opening

## Prerequisites

- PowerShell 5.1 or later
- Exchange Online PowerShell V2 module
- Appropriate Exchange Online admin permissions

## Installation

1. Clone this repository:
   ```powershell
   git clone https://github.com/RapidScripter/M365-Mailbox-Report.git
   cd M365-Mailbox-Report
   ```

2. Install the required module:
   ```powershell
   Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
   ```

## Usage

### Basic Commands

```powershell
# Interactive MFA authentication
.\GetMailboxSizeReport.ps1 -MFA

# Device code flow (for non-interactive environments)
.\GetMailboxSizeReport.ps1 -MFA -DeviceAuth

# Shared mailboxes only
.\GetMailboxSizeReport.ps1 -SharedMBOnly

# User mailboxes only
.\GetMailboxSizeReport.ps1 -UserMBOnly
```

### Advanced Options

| Parameter          | Description                          | Example                          |
|--------------------|--------------------------------------|----------------------------------|
| `-MBNamesFile`     | Path to CSV with specific mailboxes  | `-MBNamesFile .\mailboxes.csv`   |
| `-UserName`        | Admin username for automation        | `-UserName admin@domain.com`     |
| `-Password`        | Password for automation              | `-Password "yourpassword"`       |
| `-OutputPath`      | Custom output directory              | `-OutputPath "C:\Reports"`       |

### Input File Format

For `-MBNamesFile`, create a CSV with one column header "Identity":
```
Identity
user1@domain.com
user2@domain.com
sharedmbx@domain.com
```

## Output

The script generates a CSV report with these columns:

- Display Name
- User Principal Name
- Mailbox Type
- Primary SMTP Address
- Item Count
- Total Size (readable + bytes)
- Deleted Items (count + size)
- Archive Status
- Quota Settings (Warning/Send/Receive)
- Last Logon Time

Sample output filename: `MailboxSizeReport_2023-08-15_143022.csv`

## Troubleshooting

| Error | Solution |
|-------|----------|
| "Window handle must be configured" | Use `-DeviceAuth` parameter |
| "Cannot validate argument on parameter 'Filter'" | Update to latest script version |
| "No mailboxes found" | Check your filters and permissions |
