# ============================================================
#  Move-Office365Files.ps1  (PowerShell 5.1 Compatible)
#
#  Recursively copies or moves files and folders between:
#    - SharePoint Online document libraries
#    - OneDrive for Business libraries
#    - Or any combination of the two
#
#  REQUIREMENTS:
#    PowerShell 5.1
#    Install-Module PnP.PowerShell -Scope CurrentUser
#
#  RUN:
#    .\Move-Office365Files.ps1              # Move (copy + delete source)
#    .\Move-Office365Files.ps1 -WhatIf     # Dry run (no changes made)
#    .\Move-Office365Files.ps1 -CopyOnly   # Copy without deleting source
#
#  CONFIG EXAMPLES:
#
#    SharePoint site:
#      $SourceSiteUrl   = "https://<tenant>.sharepoint.com/sites/<SiteName>"
#      $SourceFolderRel = "Shared Documents/SomeFolder"
#      $DestFolderAbs   = "/sites/<SiteName>/Shared Documents/SomeFolder"
#
#    OneDrive for Business:
#      $SourceSiteUrl   = "https://<tenant>-my.sharepoint.com/personal/<user>_<domain>_com"
#      $SourceFolderRel = "Documents/SomeFolder"
#      $DestFolderAbs   = "/personal/<user>_<domain>_com/Documents/SomeFolder"
#
# ============================================================

param(
    [switch]$WhatIf,
    [switch]$CopyOnly
)

# ─── CONFIG ────────────────────────────────────────────────
# Source location
$SourceSiteUrl      = "https://<tenant>.sharepoint.com/sites/<SourceSite>"
$SourceFolderRel    = "Shared Documents/<FolderName>"   # Site-relative path

# Destination location (can be a different site or OneDrive)
$DestSiteUrl        = "https://<tenant>.sharepoint.com/sites/<DestSite>"
$DestFolderAbs      = "/sites/<DestSite>/Shared Documents/<FolderName>"  # Server-relative path

# Log file
$LogFile            = "$PSScriptRoot\MoveLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
# ───────────────────────────────────────────────────────────

$TotalCopied  = 0
$TotalFailed  = 0
$TotalDeleted = 0

# ─── LOGGING ───────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line      = "[$timestamp] [$Level] $Message"
    $color     = switch ($Level) {
        "ERROR"   { "Red" }
        "WARN"    { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "Cyan" }
    }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $LogFile -Value $line
}
# ───────────────────────────────────────────────────────────

# ─── RECURSIVE COPY FUNCTION ───────────────────────────────
# Walks every folder level and copies files + subfolders
function Copy-FolderRecursive {
    param(
        [string]$SrcFolderRelUrl,   # e.g. "Shared Documents/MyFolder/SubFolder"
        [string]$DstFolderAbsUrl    # e.g. "/sites/Dest/Shared Documents/MyFolder/SubFolder"
    )

    # ── Copy files in current folder ──────────────────────
    try {
        $files = Get-PnPFolderItem -FolderSiteRelativeUrl $SrcFolderRelUrl -ItemType File -ErrorAction Stop
    } catch {
        Write-Log "  Could not list files in '$SrcFolderRelUrl': $_" "ERROR"
        $files = @()
    }

    foreach ($file in $files) {
        $fileName = $file.Name
        $srcPath  = "$SrcFolderRelUrl/$fileName"
        $dstPath  = "$DstFolderAbsUrl/$fileName"

        Write-Log "  FILE: $srcPath"

        if ($WhatIf) {
            Write-Log "    [WHATIF] Would copy file -> $dstPath" "WARN"
            $script:TotalCopied++
            continue
        }

        try {
            Copy-PnPFile `
                -SourceUrl               $srcPath `
                -TargetUrl               $dstPath `
                -Force `
                -OverwriteIfAlreadyExists `
                -ErrorAction Stop

            Write-Log "    Copied '$fileName'" "SUCCESS"
            $script:TotalCopied++

        } catch {
            Write-Log "    FAILED to copy '$fileName': $_" "ERROR"
            $script:TotalFailed++
        }
    }

    # ── Recurse into subfolders ────────────────────────────
    try {
        $subFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $SrcFolderRelUrl -ItemType Folder -ErrorAction Stop
    } catch {
        Write-Log "  Could not list subfolders in '$SrcFolderRelUrl': $_" "ERROR"
        $subFolders = @()
    }

    foreach ($sub in $subFolders) {
        $subName      = $sub.Name
        $subSrcRelUrl = "$SrcFolderRelUrl/$subName"
        $subDstAbsUrl = "$DstFolderAbsUrl/$subName"

        Write-Log "  FOLDER: $subSrcRelUrl"

        if ($WhatIf) {
            Write-Log "    [WHATIF] Would create folder -> $subDstAbsUrl" "WARN"
            Copy-FolderRecursive -SrcFolderRelUrl $subSrcRelUrl -DstFolderAbsUrl $subDstAbsUrl
            continue
        }

        # Ensure destination subfolder exists
        try {
            Resolve-PnPFolder -SiteRelativePath ($subDstAbsUrl.TrimStart('/')) -ErrorAction Stop | Out-Null
        } catch {
            Write-Log "    Could not ensure destination folder '$subName': $_" "WARN"
        }

        Copy-FolderRecursive -SrcFolderRelUrl $subSrcRelUrl -DstFolderAbsUrl $subDstAbsUrl
    }
}
# ───────────────────────────────────────────────────────────

# ─── CHECK / INSTALL MODULE ────────────────────────────────
# Supports both the modern PnP.PowerShell and the legacy SharePointPnPPowerShellOnline module
$modulePreference = @("PnP.PowerShell", "SharePointPnPPowerShellOnline")
$moduleLoaded     = $false

foreach ($moduleName in $modulePreference) {
    if (Get-Module -ListAvailable -Name $moduleName) {
        Import-Module $moduleName -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Log "Module '$moduleName' loaded." "SUCCESS"
        $moduleLoaded = $true
        break
    }
}

if (-not $moduleLoaded) {
    Write-Log "No PnP module found. Installing PnP.PowerShell..." "WARN"
    Install-Module "PnP.PowerShell" -Scope CurrentUser -Force -AllowClobber
    Import-Module "PnP.PowerShell" -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Log "Module 'PnP.PowerShell' installed and loaded." "SUCCESS"
}
# ───────────────────────────────────────────────────────────

Write-Log "========================================"
Write-Log "  Office 365 File Mover"
Write-Log "  PS Version  : $($PSVersionTable.PSVersion)"
Write-Log "  Source      : $SourceSiteUrl / $SourceFolderRel"
Write-Log "  Destination : $DestSiteUrl$DestFolderAbs"
if ($WhatIf)   { Write-Log "  MODE: DRY RUN (no changes will be made)" "WARN" }
if ($CopyOnly) { Write-Log "  MODE: COPY ONLY (source files kept)"     "WARN" }
Write-Log "========================================"

# ─── CONNECT TO SOURCE ─────────────────────────────────────
Write-Log "Connecting to SOURCE site (browser login will open)..."
try {
    Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin -ErrorAction Stop
    Write-Log "Connected to SOURCE site." "SUCCESS"
} catch {
    Write-Log "Failed to connect to SOURCE: $_" "ERROR"
    exit 1
}

# If source and destination are on the same site, no second connection needed.
# If different sites, the copy cmdlet handles cross-site via the source connection.
$crossSite = ($SourceSiteUrl.TrimEnd('/') -ne $DestSiteUrl.TrimEnd('/'))

if ($crossSite) {
    Write-Log "Cross-site operation detected. Source and destination are on different sites."
}

# ─── SCAN SOURCE ───────────────────────────────────────────
Write-Log "Scanning source folder: $SourceFolderRel"

try {
    $topFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $SourceFolderRel -ItemType Folder -ErrorAction Stop
    $topFiles   = Get-PnPFolderItem -FolderSiteRelativeUrl $SourceFolderRel -ItemType File   -ErrorAction Stop
} catch {
    Write-Log "Failed to scan source folder: $_" "ERROR"
    Disconnect-PnPOnline
    exit 1
}

$totalItems = $topFolders.Count + $topFiles.Count
Write-Log "Found $($topFolders.Count) folder(s) and $($topFiles.Count) file(s) at root level."

if ($totalItems -eq 0) {
    Write-Log "Nothing found in source folder. Exiting." "WARN"
    Disconnect-PnPOnline
    exit 0
}

Write-Log "Top-level folders:"
foreach ($f in $topFolders) { Write-Log "   [FOLDER] $($f.Name)" }
Write-Log "Top-level files:"
foreach ($f in $topFiles)   { Write-Log "   [FILE]   $($f.Name)" }

# ─── CONFIRM ───────────────────────────────────────────────
if (-not $WhatIf) {
    Write-Log ""
    $action  = if ($CopyOnly) { "COPY" } else { "MOVE (copy + delete source)" }
    $confirm = Read-Host "Proceed with $action of ALL items? (yes/no)"
    if ($confirm -ne "yes") {
        Write-Log "Cancelled by user." "WARN"
        Disconnect-PnPOnline
        exit 0
    }
}

# ─── COPY TOP-LEVEL FILES ──────────────────────────────────
foreach ($file in $topFiles) {
    $fileName = $file.Name
    $srcPath  = "$SourceFolderRel/$fileName"
    $dstPath  = "$DestFolderAbs/$fileName"

    Write-Log "Copying root file: '$fileName'"

    if ($WhatIf) {
        Write-Log "  [WHATIF] Would copy -> $dstPath" "WARN"
        $TotalCopied++
        continue
    }

    try {
        Copy-PnPFile `
            -SourceUrl               $srcPath `
            -TargetUrl               $dstPath `
            -Force `
            -OverwriteIfAlreadyExists `
            -ErrorAction Stop

        Write-Log "  Copied '$fileName'" "SUCCESS"
        $TotalCopied++

    } catch {
        Write-Log "  FAILED for '$fileName': $_" "ERROR"
        $TotalFailed++
    }
}

# ─── RECURSIVELY COPY ALL FOLDERS ──────────────────────────
foreach ($folder in $topFolders) {
    $folderName   = $folder.Name
    $srcFolderRel = "$SourceFolderRel/$folderName"
    $dstFolderAbs = "$DestFolderAbs/$folderName"

    Write-Log ""
    Write-Log "==== [ FOLDER: $folderName ] ===="

    if (-not $WhatIf) {
        try {
            Resolve-PnPFolder -SiteRelativePath ($dstFolderAbs.TrimStart('/')) -ErrorAction Stop | Out-Null
            Write-Log "  Destination folder ready." "SUCCESS"
        } catch {
            Write-Log "  Could not create destination folder '$folderName': $_" "WARN"
        }
    }

    # Recursively copy everything inside
    Copy-FolderRecursive -SrcFolderRelUrl $srcFolderRel -DstFolderAbsUrl $dstFolderAbs

    # Delete source folder after copy (if not CopyOnly)
    if (-not $CopyOnly -and -not $WhatIf) {
        try {
            Remove-PnPFolder -Name $folderName -Folder $SourceFolderRel -Force -ErrorAction Stop
            Write-Log "  Deleted source folder '$folderName'." "SUCCESS"
            $TotalDeleted++
        } catch {
            Write-Log "  Could not delete source folder '$folderName': $_" "WARN"
        }
    }
}

# ─── DELETE REMAINING ROOT-LEVEL SOURCE FILES (MOVE MODE) ──
if (-not $CopyOnly -and -not $WhatIf) {
    foreach ($file in $topFiles) {
        $fileName = $file.Name
        $srcPath  = "$SourceFolderRel/$fileName"
        try {
            Remove-PnPFile -ServerRelativeUrl $srcPath -Force -ErrorAction Stop
            Write-Log "Deleted source file '$fileName'." "SUCCESS"
            $TotalDeleted++
        } catch {
            Write-Log "Could not delete source file '$fileName': $_" "WARN"
        }
    }
}

# ─── DISCONNECT ────────────────────────────────────────────
Disconnect-PnPOnline
Write-Log ""
Write-Log "Disconnected from SharePoint."

# ─── SUMMARY ───────────────────────────────────────────────
Write-Log "========================================"
Write-Log "  OPERATION COMPLETE"
Write-Log "  Items Copied          : $TotalCopied"
Write-Log "  Items Failed          : $TotalFailed"
Write-Log "  Items Deleted (source): $TotalDeleted"
Write-Log "  Log saved             : $LogFile"
Write-Log "========================================"

if ($TotalFailed -gt 0) {
    Write-Log "Some items failed. Check the log: $LogFile" "WARN"
    exit 1
} else {
    Write-Log "All done!" "SUCCESS"
    exit 0
}
