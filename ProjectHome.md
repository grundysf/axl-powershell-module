Powershell Module for interacting with Cisco's AXL API

Some of the commands in this module work with Unified Communications Manager, while others work with Unified Presence Server.

## Installation ##
  1. Examine your PsModulePath environment variable to see which directories are searched for Powershell modules.  You can run `$env:psmodulepath` from a PS prompt.
  1. Create a subdirectory named Axl in one of the PsModulePath directories
  1. Copy the source files (Axl.psm1, etc...) to the new subdirectory
  1. Run `get-module -listavailable` and make sure that "Axl" shows up in the list

## Usage Examples ##
To view a the contact list of a user
```
PS C:\> import-module Axl
PS C:\> $cups = New-AxlConnection cups1.example.com admin
Password for admin@cups1.example.com: ********
PS C:\> Get-UcBuddy $cups jdoe

Owner           Buddy                          Nickname               State Group
-----           -----                          --------               ----- -----
jdoe            asmith@example.com             Alison Smith           1     Friends
jdoe            mx@example.com                 Amy Taylor             1     Friends
jdoe            fbloggs@example.com            Fred Bloggs            3     Friends
```

To view available commands in this module
```
PS C:\> import-module axl
PS C:\> get-command -module axl

CommandType     Name
-----------     ----
Function        Add-UcBuddy
Function        Add-UcUserAcl
Function        Get-StoredProc
Function        Get-Table
Function        Get-UcAssignedUsers
Function        Get-UcBuddy
Function        Get-UcCupNode
Function        Get-UcCupUser
Function        Get-UcDialPlans
Function        Get-UcDN
Function        Get-UcEm
Function        Get-UcLicenseCapabilities
Function        Get-UcLineAppearance
Function        Get-UcSqlQuery
Function        Get-UcSqlUpdate
Function        Get-UcUser
Function        Get-UcUserAcl
Function        Get-UcWatcher
Function        New-AxlConnection
Function        Remove-UcBuddy
Function        Rename-UcBuddy
Function        Set-AddAssignedSubClusterUsersByNode
Function        Set-UcCupUser
Function        Set-UcEm
Function        Set-UcLicenseCapabilities
Function        Set-UcPhone
```

Remove a particular contact from everyone's contact list
```
Get-UcBuddy -axl $cups -owner % -buddy jdoe@example.com | Remove-UcBuddy -axl $cups
```
Same thing but using positional parameters
```
Get-UcBuddy $cups % jdoe@example.com | Remove-UcBuddy $cups
```

If you wish to contribute to this project, please contact me at yttep.giarc@gmail.com