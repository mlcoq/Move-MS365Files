# Office365FileMover

PowerShell script to recursively copy or move files and folders between SharePoint Online and/or OneDrive for Business locations using PnP PowerShell.

## Requirements

- PowerShell 5.1 or later
- PnP.PowerShell module (or legacy `SharePointPnPPowerShellOnline`)

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

## Usage

```powershell
# Dry run — shows what would happen, no changes made
.\Move-Office365Files.ps1 -WhatIf

# Copy only — keeps source files intact
.\Move-Office365Files.ps1 -CopyOnly

# Full move — copies to destination, then deletes source
.\Move-Office365Files.ps1
```

## Configuration

Edit the variables at the top of the script:

| Variable | Description |
|---|---|
| `$SourceSiteUrl` | Full URL of the source SharePoint / OneDrive site |
| `$SourceFolderRel` | Site-relative path of the source folder |
| `$DestSiteUrl` | Full URL of the destination site |
| `$DestFolderAbs` | Server-relative path of the destination folder |

### SharePoint Online example

```powershell
$SourceSiteUrl   = "https://<tenant>.sharepoint.com/sites/<SourceSite>"
$SourceFolderRel = "Shared Documents/Reports"

$DestSiteUrl     = "https://<tenant>.sharepoint.com/sites/<DestSite>"
$DestFolderAbs   = "/sites/<DestSite>/Shared Documents/Reports"
```

### OneDrive for Business example

```powershell
$SourceSiteUrl   = "https://<tenant>-my.sharepoint.com/personal/<user>_<domain>_com"
$SourceFolderRel = "Documents/Reports"

$DestSiteUrl     = "https://<tenant>.sharepoint.com/sites/<DestSite>"
$DestFolderAbs   = "/sites/<DestSite>/Shared Documents/Reports"
```

> **Note:** OneDrive for Business is supported because it runs on SharePoint Online.  
> Personal OneDrive (consumer Microsoft accounts) is **not** supported.

## Features

- Recursive copy of all files and subfolders
- Dry run mode (`-WhatIf`) — no changes are made
- Copy-only mode — source files are not deleted
- Full move mode — copies then deletes source
- Timestamped log file saved next to the script
- Confirmation prompt before any changes
- Works across different sites and between SharePoint ↔ OneDrive for Business
- Auto-installs PnP module if not present
