# Office365FileMover

GUI-based PowerShell tool to copy or move files between SharePoint Online and OneDrive for Business.

The script asks for:
- Tenant name
- Source type: SharePoint or OneDrive
- Destination type: SharePoint or OneDrive
- Library name
- Optional subfolder (empty = root of library)
- OneDrive email when OneDrive is selected
- Action: Kopieren or Verplaatsen
- Optional WhatIf mode (no writes)

## Requirements

- PowerShell 5.1 or later
- PnP.PowerShell module

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

## Run

```powershell
.\Move-Office365Files.ps1
```

## Workflow

1. Vul tenant, bron en doel in.
2. Klik `Ophalen overzicht` om structuur op te halen van de bron:
	 - totaal aantal files
	 - totaal aantal folders
	 - totale grootte in GB
3. (Optioneel) Klik `Export CSV` om dit overzicht op te slaan.
4. Kies `Kopieren` of `Verplaatsen`.
5. Klik `Start`.

## Input notes

- For SharePoint:
	- `Site pad` is relative to tenant root.
	- `Library` should be the target document library name.
- For OneDrive:
	- Enter the OneDrive owner email.
	- Script converts email to the correct `/personal/...` URL segment.
	- `Library` should be the target document library name.
- Empty subfolder means library root.

## Supported combinations

- SharePoint -> SharePoint
- SharePoint -> OneDrive for Business
- OneDrive for Business -> SharePoint
- OneDrive for Business -> OneDrive for Business

## Important

- OneDrive for Business is supported.
- Personal OneDrive (consumer Microsoft accounts) is not supported.
- Authentication uses interactive login and may be done with an admin account.
- Source and destination library are validated before migration starts.
- Log file is created next to the script with timestamp.
