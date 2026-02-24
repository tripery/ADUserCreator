ADUserCreator
=============
PowerShell application for bulk Active Directory user creation from Excel
из WinForms graphical interface.
------------------------------------------------------------
FEATURES
--------
- Import users от Excel (.xlsx)
- автоматический UA → Latin transliteration
- Generates:
- SamAccountName
- Email
- UPN
- OU selection via GUI
- AD group selection
- Random password generation
- Splash screen
- Can be built into EXE
- Installer support
------------------------------------------------------------
REQUIREMENTS
------------
Требуется:
- Windows 10/(x64)
- PowerShell 5.1 или higher
- RSAT ActiveDirectory module
- Network access to domain (LAN или VPN)

Optional:

- ImportExcel PowerShell module (installer can install automatically)
------------------------------------------------------------
Installer is created using Inno Setup.
User workflow:
- Download Setup
- Next → Install
- Desktop shortcut создан
- Application ready to use
Installer может автоматически install:
- RSAT ActiveDirectory
- ImportExcel
(Admin rights required)
------------------------------------------------------------
COMMON PROBLEMS
---------------
ActiveDirectory module missing:
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
ImportExcel missing:
Install-Module ImportExcel -Scope CurrentUser
Cannot connect to domain:
- Check VPN
- Check DNS
- Check permissions
------------------------------------------------------------
SECURITY
- Application runs under current Windows user
- Passwords generated locally
- No external data transfer
------------------------------------------------------------
AUTOR
Internal Active Directory — автоматический инструмент.
https://www.linkedin.com/in/ruslan-kucher-27bb3736a/
------------------------------------------------------------
If something breaks:
Check RSAT, VPN and permissions first.
