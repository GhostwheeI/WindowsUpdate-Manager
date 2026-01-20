### `README.md`
```md
# Windows Update Manager v2
A **production-ready, menu-driven PowerShell tool** that wraps **PSWindowsUpdate** to make Windows Update management safer and easier: scan, install, hide/unhide, uninstall, export history, and run remote update jobs — with **sane defaults** and **audit-friendly logging**.

> ⭐ If this saved you time, starring the repo helps prioritize maintenance and improvements.

## 30-second quick start
1) Run PowerShell as Admin (the script can auto-elevate):
```powershell
.\WindowsUpdate-Manager_v2.ps1
```

2) First run will guide you through setup (module/dependency checks + preferences).

## What you can do
- Scan for updates (Windows Update or Microsoft Update)
- Install **all** or **selected** updates
- Hide/unhide updates (with a persistent KB exclusion list)
- Uninstall updates by KB
- View/export update history + scan results
- Generate a compliance report
- Run remote installs via Invoke-WUJob (SYSTEM scheduled task)

## Who this is for
- Sysadmins who want a repeatable update workflow
- Helpdesk/junior staff who need guardrails (menus + confirmations)
- Servers/workstations where GUI tools are undesirable
- Environments where update actions must be logged/auditable

## Documentation
- Full user guide: **docs/USER_GUIDE.md**
- PSWindowsUpdate: see the PowerShell Gallery package page

## Safety
Windows Updates can reboot systems and remove updates. Test in non-production first. Use maintenance windows in server environments.

## Contributing
PRs welcome. If you’re not sure where to start, look for **good first issue**.
```

---

### `docs/USER_GUIDE.md`
```md
# Windows Update Manager v2.0 - User Guide

**Author:** Ghostwheel
**Version:** 2.0.0
**Created:** 2026-01-19

---

## Overview

Windows Update Manager v2 is a professional, production-ready PowerShell script that provides comprehensive Windows Update management through an intuitive menu-driven interface. This enhanced version includes automatic dependency management, system safety features, and advanced configuration options.

---

## 🎯 Key Features

### Automatic Dependency Management
- ✅ **Auto-detection** of PowerShell version with upgrade recommendations
- ✅ **Auto-installation** of NuGet provider when needed
- ✅ **Auto-installation** of PSWindowsUpdate module
- ✅ **PowerShell 7.5+ installation** offering with clear benefits explanation
- ✅ **Intelligent fallback** to Windows PowerShell 5.1 if user declines upgrade

### Enhanced User Interface
- 🎨 **Color-coded messages** (Info: Cyan, Success: Green, Warning: Yellow, Error: Red)
- 📊 **Status indicators** (✓ Success, ✗ Error, ⚠ Warning, ℹ Info)
- 📈 **Progress tracking** for long operations
- 🖥️ **Professional banner** and menu system
- 👁️ **Visual elevation indicator** (shows when running as Administrator)

### Safety Features
- 🛡️ **Pre-flight system checks** (internet connectivity, disk space, service status)
- 💾 **Automatic restore point creation** before updates (configurable)
- 📝 **Comprehensive error logging** with timestamps
- 🔄 **Automatic self-elevation** to Administrator when needed
- ⚠️ **Confirmation prompts** for destructive operations

### Update Management
- 🔍 **Smart update scanning** with category filtering
- 📦 **Install all or selected** updates
- 🚫 **KB exclusion list** (persistent hide list)
- 👁️ **Hide/unhide updates** by selection or KB number
- ❌ **Uninstall updates** by KB number
- 📜 **Update history** viewing and export
- 🌐 **Remote update management** via Invoke-WUJob

### Configuration & Reporting
- ⚙️ **Persistent JSON configuration** with validation
- 🎛️ **First-run wizard** for initial setup
- 📊 **Compliance report generation**
- 📁 **Export to CSV** for scan results and history
- 🔧 **Flexible logging** with transcript support

---

## 📋 Requirements

### Minimum Requirements
- Windows 10/11 or Windows Server 2016+
- Windows PowerShell 5.1 or PowerShell 7.x
- Administrator privileges (script will auto-elevate)
- Internet connectivity for updates

### Recommended
- PowerShell 7.5 or later (script will offer to install)
- 10GB+ free disk space
- Active Windows Update service

### Dependencies (Auto-Installed)
- NuGet package provider (2.8.5.201+)
- PSWindowsUpdate module (latest version)

---

## 🚀 Getting Started

### First Run

1. **Right-click** the script and select "Run with PowerShell"
   - Or open PowerShell and run: `.\WindowsUpdate-Manager_v2.ps1`

2. The script will:
   - Check for Administrator privileges and **elevate automatically**
   - Check dependencies and **install if missing**
   - Offer **PowerShell 7.5+ installation** (optional)
   - Run the **first-run configuration wizard**

3. Follow the on-screen prompts to configure your preferences

### Configuration Wizard

On first run, you'll be asked to configure:

1. **Update Source**
   - Microsoft Update (includes Windows + Microsoft products) - Recommended
   - Windows Update only

2. **Auto Reboot**
   - Whether to automatically reboot after updates if required

3. **Create Restore Point**
   - Whether to create system restore points before installing updates (Recommended)

---

## 📚 Menu Options

### Update Operations

**[1] Status**
- View installer status, reboot requirements, last results, and service managers
- Quick health check of Windows Update system

**[2] Scan for Updates**
- Scan for available updates with optional filters
- Shows summary by severity (Critical, Important, Moderate, Low)
- Export scan results to CSV

**[3] Install ALL Updates**
- Install all available updates automatically
- Creates restore point before installation
- Includes pre-flight system checks

**[4] Install SELECTED Updates**
- Choose specific updates from last scan
- Select multiple updates using Out-GridView or console selection
- Checks against KB exclusion list

### Update Management

**[5] Hide SELECTED Updates**
- Hide unwanted updates
- Automatically adds to KB exclusion list
- Prevents updates from appearing in future scans

**[6] Unhide by KB**
- Unhide previously hidden updates
- Removes from KB exclusion list

**[7] Uninstall by KB**
- Uninstall specific updates by KB number
- Creates restore point before uninstallation
- Use with caution

**[8] Show Update History**
- View update installation history
- Filter by number of entries
- Export history to CSV

### System Maintenance

**[9] Reset Windows Update Components**
- Resets Windows Update services and cache
- Use when experiencing persistent update problems
- Creates restore point before reset
- May require reboot

**[10] Update PSWindowsUpdate Module**
- Check for and install latest version of PSWindowsUpdate
- Ensures you have the latest features and fixes

### Configuration

**[11] Target Computers**
- Configure remote computers for update management
- Supports comma-separated list of computer names
- Leave blank for local computer only

**[12] Settings**
- Toggle update source (Windows Update / Microsoft Update)
- Toggle auto reboot
- Toggle ignore reboot
- Toggle verbose output
- Toggle system restore point creation
- Toggle logging (transcript)
- Manage KB exclusion list

### Advanced

**[13] Remote Install Job**
- Create SYSTEM-level scheduled task on remote computers
- Reliable remote update installation
- Requires PowerShell remoting enabled

**[14] Generate Compliance Report**
- Generate detailed compliance report
- Includes system info, update status, and available updates
- Exports to text file

---

## 🎨 Color Guide

The script uses colors to make information easier to understand:

- **Cyan** (ℹ) - Informational messages
- **Green** (✓) - Success messages
- **Yellow** (⚠) - Warning messages
- **Red** (✗) - Error messages
- **Gray** - Secondary information

---

## 🛡️ Pre-Flight Checks

Before major operations, the script performs automatic checks:

1. **Internet Connectivity** - Tests connection to update servers
2. **Disk Space** - Ensures sufficient free space (10GB+ recommended)
3. **Windows Update Service** - Verifies service is running
4. **PowerShell Version** - Shows version information (detailed mode)

Failed checks will prompt you to decide whether to continue.

---

## ⚙️ Configuration File

The script stores configuration in JSON format at:
```
%APPDATA%\WindowsUpdateManager\config.json
```

### Configuration Options

```json
{
  "UseMicrosoftUpdate": true,
  "AutoReboot": false,
  "IgnoreReboot": false,
  "Verbose": false,
  "TargetComputers": [],
  "CreateRestorePoint": true,
  "ExcludedKBs": [],
  "AutoAcceptEULA": false
}
```

- **UseMicrosoftUpdate**: Use Microsoft Update (true) or Windows Update only (false)
- **AutoReboot**: Automatically reboot after updates if required
- **IgnoreReboot**: Install updates even if reboot is pending
- **Verbose**: Show detailed output from PSWindowsUpdate cmdlets
- **TargetComputers**: Array of remote computer names (empty = local)
- **CreateRestorePoint**: Create system restore point before updates
- **ExcludedKBs**: Array of KB numbers to exclude (hidden updates)
- **AutoAcceptEULA**: Automatically accept EULAs (use with caution)

---

## 📝 Logging

### Transcript Logging
Enable via Settings menu ([12] → [6])

Logs stored in:
```
%APPDATA%\WindowsUpdateManager\logs\Transcript_YYYYMMDD_HHMMSS.log
```

### Error Logging
Automatic error logging to:
```
%APPDATA%\WindowsUpdateManager\logs\ErrorLog_YYYYMMDD_HHMMSS.txt
```

Logs all errors with timestamps and exception details.

---

## 🔧 Command-Line Parameters

```powershell
.\WindowsUpdate-Manager_v2.ps1 [parameters]
```

### Available Parameters

**-NoElevation**
- Prevents automatic elevation to Administrator
- Use only if already running elevated

**-ConfigPath** <path>
- Custom configuration file path
- Default: `%APPDATA%\WindowsUpdateManager\config.json`

**-LogPath** <path>
- Custom transcript log path
- Default: Auto-generated in logs folder

**-SkipDependencyCheck**
- Skips automatic dependency checking and installation
- Use only if dependencies are confirmed installed

**-Silent**
- Minimizes user prompts for automated operation
- Use with caution

### Examples

```powershell
# Run with default settings
.\WindowsUpdate-Manager_v2.ps1

# Run without elevation (already admin)
.\WindowsUpdate-Manager_v2.ps1 -NoElevation

# Custom config and log paths
.\WindowsUpdate-Manager_v2.ps1 -ConfigPath "C:\Config\wu.json" -LogPath "C:\Logs\wu.log"

# Skip dependency checks (faster startup)
.\WindowsUpdate-Manager_v2.ps1 -SkipDependencyCheck
```

---

## 🎯 Common Use Cases

### Scenario 1: First-time Setup
1. Run the script
2. Follow the first-run wizard
3. Let dependencies auto-install
4. Consider upgrading to PowerShell 7.5+

### Scenario 2: Scan and Install Specific Updates
1. Select **[2] Scan for updates**
2. Review the available updates
3. Select **[4] Install SELECTED updates**
4. Choose updates from the list
5. Confirm installation

### Scenario 3: Hide Unwanted Updates
1. Select **[2] Scan for updates**
2. Select **[5] Hide SELECTED updates**
3. Choose updates to hide
4. Updates are hidden and added to exclusion list

### Scenario 4: Remote Update Management
1. Select **[11] Target computers**
2. Enter comma-separated computer names
3. Select **[13] Remote install job**
4. Specify targets and options
5. Job runs as SYSTEM on remote computers

### Scenario 5: Windows Update Problems
1. Select **[9] Reset Windows Update components**
2. Confirm the reset operation
3. Restore point is created automatically
4. Reboot if prompted

### Scenario 6: Compliance Reporting
1. Select **[14] Generate compliance report**
2. Report is generated with system info and update status
3. Open report for review or archival

---

## 🔒 Security Considerations

### Administrator Privileges
- Required for most update operations
- Script auto-elevates when needed
- Visual indicator shows elevation status

### System Restore Points
- Created before major operations
- Provides rollback capability
- Can be disabled in settings if needed

### KB Exclusion List
- Persistent list of hidden updates
- Prevents unwanted updates from installing
- Manage via Settings → Manage KB Exclusion List

### Remote Operations
- Requires PowerShell remoting enabled
- Uses SYSTEM account for reliability
- Ensure appropriate network security

---

## 🐛 Troubleshooting

### Issue: PSWindowsUpdate Won't Install

**Solution:**
1. Ensure you're running as Administrator
2. Check internet connectivity
3. Verify NuGet provider is installed
4. Try manual installation:
   ```powershell
   Install-PackageProvider -Name NuGet -Force
   Install-Module -Name PSWindowsUpdate -Force
   ```

### Issue: Updates Won't Install

**Solution:**
1. Run pre-flight checks
2. Check disk space (need 10GB+)
3. Verify Windows Update service is running
4. Try resetting Windows Update components ([9])

### Issue: Elevation Fails

**Solution:**
1. Right-click PowerShell → "Run as Administrator"
2. Then run the script manually
3. Or use `-NoElevation` parameter

### Issue: Remote Computers Not Responding

**Solution:**
1. Verify PowerShell remoting is enabled:
   ```powershell
   Enable-PSRemoting -Force
   ```
2. Check firewall rules
3. Verify network connectivity
4. Ensure PSWindowsUpdate is installed on targets

### Issue: Script Hangs on History

**Solution:**
1. Limit history entries with a number (e.g., 50)
2. Windows Update history can be very large
3. Consider using smaller numbers for initial query

---

## 📊 Performance Tips

1. **Use PowerShell 7.5+** for best performance
2. **Enable verbose mode** only when troubleshooting
3. **Limit history queries** to recent entries
4. **Use local targeting** for faster operations
5. **Schedule off-peak** for large update installations

---

## 🔄 Upgrade from v1

If you're upgrading from the original script:

### What's New in v2
- Automatic dependency management
- Enhanced color-coded UI
- Pre-flight system checks
- System restore point creation
- KB exclusion list
- Compliance reporting
- Improved error handling and logging
- First-run configuration wizard
- PowerShell 7.5+ installation option

### Configuration Migration
- v2 will create a new config file
- Old settings can be manually re-entered via Settings menu
- No data loss - old config is not modified

---

## 📞 Support & Feedback

### Getting Help
- Run the script with `-Verbose` parameter for detailed output
- Check error logs in `%APPDATA%\WindowsUpdateManager\logs\`
- Review transcript logs if logging is enabled

### Documentation
- Built-in help: `Get-Help .\WindowsUpdate-Manager_v2.ps1 -Full`
- PSWindowsUpdate docs: https://www.powershellgallery.com/packages/PSWindowsUpdate

---

## 📜 License

This script is provided "as-is" without warranty of any kind. Use at your own risk.

---

## 🎓 Best Practices

1. **Always create restore points** before major changes
2. **Test on non-production systems** first
3. **Review updates** before installing (use scan first)
4. **Keep exclusion list** current and minimal
5. **Regular compliance reports** for audit trails
6. **Schedule updates** during maintenance windows
7. **Verify pre-flight checks** before mass deployments
8. **Enable logging** for troubleshooting and audits

---

## 🚀 Quick Start Checklist

- [ ] Run script (auto-elevates to Administrator)
- [ ] Complete first-run wizard
- [ ] Consider installing PowerShell 7.5+
- [ ] Verify dependencies are installed
- [ ] Configure settings to your preferences
- [ ] Run initial scan for updates
- [ ] Review and install updates as needed
- [ ] Set up KB exclusions if needed
- [ ] Enable logging for audit trail
- [ ] Generate compliance report

---

**Enjoy professional Windows Update management with WindowsUpdate-Manager v2!**

For questions or issues, please refer to the error logs or enable verbose mode for detailed diagnostics.
```
