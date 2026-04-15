ADUserCreator
=============

PowerShell desktop application for bulk Active Directory user creation from Excel.

Features
--------

- Import users from Excel (`.xlsx`)
- Generate `SamAccountName`, `Email`, and `UPN`
- Select target OU from a GUI dialog
- Select AD groups
- Generate random passwords
- Create password PDF output
- Build desktop EXE and installer

Requirements
------------

- Windows 10/11 x64
- PowerShell 5.1 or newer
- RSAT ActiveDirectory module
- Network access to the domain

Optional:

- `ImportExcel` PowerShell module

Run
---

Use only these 2 launch options:

Docker UI mode:

```powershell
.\scripts\start.cmd
```

Local UI mode:

```powershell
.\scripts\start.cmd -UiMode Local
```

Desktop Build
-------------

Build EXE and installer:

```powershell
powershell -ExecutionPolicy Bypass -File .\build\build-exe.ps1
```

The installer is created in:

- `release\`

Common Problems
---------------

ActiveDirectory module missing:

```powershell
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

ImportExcel missing:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

PowerShell script execution blocked:

Use the launcher:

```powershell
.\scripts\start.cmd
```

Security
--------

- Application runs under the current Windows user
- Passwords are generated locally
- No external data transfer is required for the desktop tool
