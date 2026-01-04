# 🔍 Enhanced Windows Event Log Viewer

## 📋 Executive Summary

The Enhanced Windows Event Log Viewer is a powerful PowerShell-based GUI tool designed to simplify the investigation and analysis of Windows Event Logs. Built with Windows Presentation Foundation (WPF), this tool provides system administrators and IT professionals with an intuitive interface to search, filter, analyze, and export event log data across multiple categories including security events, system health, and application issues.

### ✨ Key Highlights

- **🎯 Pre-configured Event Categories**: Five curated categories covering the most critical Windows events
- **🖥️ Local & Remote Support**: Query event logs from your local machine or remote computers
- **📊 Advanced Filtering**: Filter by date range, event ID, and message text
- **💾 Multiple Export Formats**: Export results to CSV or Excel with full formatting
- **🔄 Real-time Monitoring**: Auto-refresh capability for continuous monitoring
- **🌐 Web Search Integration**: Instantly search for event information using Google or DuckDuckGo
- **📄 Pagination System**: Efficiently handle large result sets with customizable page sizes
- **👁️ Enhanced Visibility**: Alternating row colors and hover effects for easy reading

---

## 🚀 Quick Start

### Prerequisites

- Windows PowerShell 5.1 or later
- Windows Operating System (Windows 10/11 or Windows Server 2016+)
- Administrative privileges (required for accessing Security logs)
- Microsoft Excel (optional, for Excel export functionality)

### Installation

1. **Download the script**
   ```powershell
   # Clone the repository
   git clone https://github.com/ChrisMunnPS/EventIDChecker.git
   cd EventIDChecker
   ```

2. **Run the script**
   ```powershell
   # Execute with PowerShell (requires Administrator privileges)
   .\EventIDSearch.ps1
   ```

   Or right-click the script and select "Run with PowerShell"

---

## 📖 Features Overview

### 🗂️ Event Categories

The tool organizes Windows events into five pre-configured categories:

| Category | Log Source | Event Count | Description |
|----------|------------|-------------|-------------|
| **Account Activity** | Security | 10 events | Login/logout events, authentication attempts |
| **AD Account Changes** | Security | 12 events | User account modifications, group changes |
| **Security Threat Indicators** | Security | 7 events | Suspicious activities, audit log tampering |
| **Server Health & Reliability** | System | 10 events | System crashes, service failures, DNS issues |
| **Application-Level Issues** | Application | 5 events | Application crashes, hangs, installation events |

### 🔧 Core Capabilities

#### 🎛️ Filtering Options
- **📅 Date Range**: Select custom start and end dates (default: last 3 days)
- **🔢 Event ID**: Filter by specific event IDs or view all events in a category
- **🔎 Text Search**: Search within event messages for specific keywords
- **💻 Computer Name**: Query local or remote computers

#### 📤 Export Functions
- **CSV Export**: 
  - Exports all filtered results
  - Includes timestamp, log name, source, event ID, meaning, level, and message
  - Auto-generated filename with timestamp
  
- **Excel Export**:
  - Formatted workbook with headers
  - Auto-fit columns for readability
  - Color-coded header row
  - Requires Microsoft Excel installed

#### 🔄 Auto-Refresh
- Enable automatic refresh every 30 seconds
- Perfect for monitoring ongoing issues
- Visual indicator when enabled
- Can be toggled on/off at any time

#### 🌐 Web Search Integration
- Select any event and click "Web Search"
- Choose between Google or DuckDuckGo
- Automatically searches for: `Event ID [number] [Source] Windows`
- Opens results in default browser

#### 📊 Pagination
- Configurable page sizes: 50, 100, 250, or 500 events
- Navigation buttons: First, Previous, Next, Last
- Shows current page and total pages
- Displays event count range

---

## 🛠️ Technical Documentation

### Architecture

The application is built using:
- **WPF (Windows Presentation Foundation)** for the UI layer
- **PowerShell** for business logic and Windows Event Log API integration
- **XAML** for declarative UI design
- **COM Interop** for Excel export functionality

### Event ID Reference

<details>
<summary>📌 Account Activity Events (Security Log)</summary>

| Event ID | Meaning |
|----------|---------|
| 4624 | Successful logon |
| 4625 | Failed logon attempt |
| 4634 | Logoff |
| 4647 | User-initiated logoff |
| 4648 | Explicit credentials used |
| 4672 | Security ID assigned to user |
| 4768 | Kerberos TGT request |
| 4769 | Kerberos service ticket request |
| 4771 | Kerberos pre-authentication failed |
| 4776 | NTLM authentication attempt |

</details>

<details>
<summary>🔐 AD Account Changes (Security Log)</summary>

| Event ID | Meaning |
|----------|---------|
| 4720 | User account created |
| 4722 | User account enabled |
| 4723 | Password change attempt |
| 4724 | Password reset attempt |
| 4725 | User account disabled |
| 4726 | User account deleted |
| 4732 | User added to security group |
| 4735 | Security group modified |
| 4738 | User account modified |
| 4740 | Account locked out |
| 4756 | Security group membership change |
| 4767 | Account unlocked |

</details>

<details>
<summary>⚠️ Security Threat Indicators (Security Log)</summary>

| Event ID | Meaning |
|----------|---------|
| 1102 | Audit log cleared |
| 2886 | LDAP unsigned/simple bind detected |
| 2887 | Count of unsigned/simple bind attempts |
| 2889 | Source of unsigned/simple bind |
| 1644 | Expensive LDAP query detected |
| 4627 | Group membership information |
| 4663 | Access to an object |

</details>

<details>
<summary>🖥️ Server Health and Reliability (System Log)</summary>

| Event ID | Meaning |
|----------|---------|
| 41 | Kernel-Power: unexpected restart/shutdown |
| 55 | NTFS file system corruption detected |
| 6005 | Event log service started |
| 6006 | Event log service stopped |
| 6008 | Unexpected shutdown |
| 6009 | System startup information |
| 1074 | System shutdown/restart initiated |
| 1014 | DNS name resolution failure |
| 1058 | Group Policy failure to read from DC |
| 5719 | Netlogon: no DC available |

</details>

<details>
<summary>🖥️ Application-Level Issues (Application Log)</summary>

| Event ID | Meaning |
|----------|---------|
| 1000 | Application error (crash) |
| 1001 | Application hang or bugcheck info |
| 1002 | Application hang |
| 1309 | ASP.NET application error (IIS) |
| 11707 | Application installation |

</details>

### Code Structure

```
EventLogViewer.ps1
├── Assembly Loading (WPF, Windows Forms)
├── Category Definitions (Hash table with event mappings)
├── XAML UI Definition
│   ├── Category Selection (Radio buttons)
│   ├── Date/Computer Selection
│   ├── Event ID and Text Filters
│   ├── Action Buttons
│   ├── Progress Bar
│   ├── Results DataGrid
│   └── Pagination Controls
├── UI Control References
├── Helper Functions
│   ├── Update-EventIDComboBox
│   ├── Update-PaginationDisplay
│   └── Invoke-EventSearch
└── Event Handlers
    ├── Search Operations
    ├── Export Functions
    ├── Auto-Refresh Timer
    ├── Pagination Navigation
    └── Web Search Integration
```

### PowerShell Cmdlets Used

- `Get-WinEvent`: Primary cmdlet for retrieving Windows Event Logs
- `Add-Type`: Loading .NET assemblies
- `New-Object`: Creating COM objects and WPF controls
- `Start-Process`: Launching web browser for searches
- `Export-Csv`: CSV export functionality

### Performance Considerations

- **Pagination**: Results are paginated to prevent UI freezing with large datasets
- **Lazy Loading**: Only the current page is rendered in the DataGrid
- **Progress Indicators**: Visual feedback during long-running queries
- **COM Object Cleanup**: Proper disposal of Excel COM objects to prevent memory leaks

---

## 📸 Screenshots

### Main Interface
![Main Interface](00%20-%20EventID_Main.jpg)

### Account Activity Filter
![Account Activity Filter](01%20-%20EventID_Account_Filter.jpg)

### Account Activity - All Items View
![Account Activity All Items](01a%20-%20EventID_Account_All_Items.jpg)

### AD Account Changes Filter
![AD Account Changes](02%20-%20EventID_ADAccount_Changes_Filter.jpg)

### Security Threat Indicators Filter
![Security Threat Indicators](03%20-%20EventID_SecurityThreatIndicators_Filter.jpg)

### Server Health and Reliability Filter
![Server Health](04%20-%20EventID_ServerHealthAndReliability_Filter.jpg)

### Application Level Issues Filter
![Application Issues](05%20-%20EventID_ApplicationLevelIssues_Filter.jpg)

---

## 🔒 Security Considerations

### Required Permissions

- **Security Log Access**: Requires membership in:
  - Administrators group, or
  - Event Log Readers group
  
- **Remote Computer Access**: Additional requirements:
  - Remote Registry service running
  - Appropriate firewall rules
  - Remote Event Log Management enabled
  - Network access to target computer

### Best Practices

1. ⚠️ Run with least privilege necessary
2. 🔐 Use service accounts for remote queries
3. 📝 Audit script usage in production environments
4. 🛡️ Review and validate event IDs before deployment
5. 🔒 Secure exported files containing sensitive information

---

## 🐛 Troubleshooting

### Common Issues

**Problem**: "Access Denied" when running script
- **Solution**: Run PowerShell as Administrator

**Problem**: No events returned from Security log
- **Solution**: Verify you're in the Event Log Readers or Administrators group

**Problem**: Excel export fails
- **Solution**: Ensure Microsoft Excel is installed and not already running

**Problem**: Remote computer connection fails
- **Solution**: 
  - Verify network connectivity (`Test-Connection`)
  - Check Windows Firewall rules
  - Ensure Remote Event Log Management is enabled

**Problem**: Script execution blocked
- **Solution**: Set execution policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for:

- 🐛 Bug fixes
- ✨ New features
- 📚 Documentation improvements
- 🎨 UI/UX enhancements

### Development Guidelines

1. Follow PowerShell best practices
2. Use approved PowerShell verbs for functions
3. Include error handling for all operations
4. Test on multiple Windows versions
5. Update documentation for new features

---

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

---

## 🙏 Acknowledgments

- Microsoft Documentation for Windows Event Log reference
- PowerShell Community for WPF guidance
- All contributors and users providing feedback and support

---

## ⭐ Show Your Support

If this tool has been helpful, please consider:
- ⭐ Starring the repository
- 🐛 Reporting bugs or suggesting features
- 🤝 Contributing improvements
- ☕ [Buying me a coffee](https://buymeacoffee.com/ChrisMunnPS)

---

## 📞 Support

For issues, questions, or suggestions:
- 🐛 [Open an issue on GitHub](https://github.com/ChrisMunnPS/EventIDChecker/issues)
- 💬 [GitHub Discussions](https://github.com/ChrisMunnPS/EventIDChecker/discussions)
- ☕ [Buy me a coffee](https://buymeacoffee.com/ChrisMunnPS)

---

**Version**: 1.0.0  
**Last Updated**: January 2026  
**Compatibility**: Windows 10/11, Windows Server 2016+
