# Windows Update Manager v3

A **production-ready, menu-driven PowerShell tool** that wraps **PSWindowsUpdate** to make Windows Update management safer and easier: scan, install, hide/unhide, uninstall, export history, and run remote update jobs — with **sane defaults** and **audit-friendly logging**.

> ⭐ If this saved you time, starring the repo helps prioritize maintenance and improvements.

## 30-second quick start
1) Run PowerShell as Admin (the script can auto-elevate):
```powershell
.\WindowsUpdate-Manager.ps1
```

2) First run will guide you through setup (module/dependency checks + preferences).

## What's new in v3
- **Visual Enhancements:** ASCII art titles with multiple styles (Standard, Slant, Cyberpunk)
- **Vibrant UI:** Color gradients and 24-bit RGB support for PowerShell 7+
- **Immersive Feedback:** Audio notifications and sound effects for key actions
- **Improved Layout:** Responsive layout adapting to console size with professional borders

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

## Safety
Windows Updates can reboot systems and remove updates. Test in non-production first. Use maintenance windows in server environments.

## Contributing
PRs welcome. If you’re not sure where to start, look for **good first issue**.
