# ============================================================
#  MS365Mover (GUI)
#
#  Migrate files between SharePoint Online and OneDrive for Business.
#  Input is collected via a WinForms GUI:
#   - Tenant name
#   - Source type: SharePoint or OneDrive
#   - Destination type: SharePoint or OneDrive
#   - Library + optional subfolder (root allowed)
#   - OneDrive email (converted to personal URL segment)
#   - Action: Copy or Move
#
#  Requires:
#   - PowerShell 5.1+
#   - PnP.PowerShell module
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:LogFile = $null
$script:FileLoggingEnabled = $false
$script:LogTextBox = $null
$script:LastOverview = $null

function Ensure-RequiredPowerShell {
    $required = [version]"7.4.0"
    if ($PSVersionTable.PSVersion -ge $required) {
        return
    }

    $pwshCmd = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($null -ne $pwshCmd) {
        [System.Windows.Forms.MessageBox]::Show(
            "This script needs PowerShell 7.4+ for PnP.PowerShell. It will now restart in PowerShell 7.",
            "Restarting in PowerShell 7",
            "OK",
            "Information"
        ) | Out-Null

        Start-Process -FilePath $pwshCmd.Source -ArgumentList @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$PSCommandPath`""
        ) -WorkingDirectory $PSScriptRoot

        exit
    }

    throw "PowerShell 7.4+ is required for PnP.PowerShell, but 'pwsh' was not found. Install PowerShell 7 and run the script again."
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line
    if ($script:FileLoggingEnabled -and -not [string]::IsNullOrWhiteSpace($script:LogFile)) {
        Add-Content -Path $script:LogFile -Value $line
    }

    if ($null -ne $script:LogTextBox) {
        $script:LogTextBox.AppendText($line + [Environment]::NewLine)
        $script:LogTextBox.SelectionStart = $script:LogTextBox.TextLength
        $script:LogTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Ensure-PnPModule {
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-Log "PnP.PowerShell not found. Installing..." "WARN"
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-Log "PnP.PowerShell module loaded." "SUCCESS"
}

function Join-UrlPath {
    param(
        [string[]]$Segments
    )

    $clean = @()
    foreach ($s in $Segments) {
        if (-not [string]::IsNullOrWhiteSpace($s)) {
            $clean += ($s.Trim("/"))
        }
    }

    if ($clean.Count -eq 0) {
        return ""
    }

    return ($clean -join "/")
}

function Convert-EmailToOneDriveSegment {
    param([string]$Email)

    $mail = $Email.Trim().ToLowerInvariant()
    $segment = $mail -replace "[^a-z0-9]", "_"
    return $segment
}

function Build-Endpoint {
    param(
        [string]$Tenant,
        [string]$Type,
        [string]$SitePath,
        [string]$Library,
        [string]$SubPath,
        [string]$OneDriveEmail
    )

    if ([string]::IsNullOrWhiteSpace($Tenant)) {
        throw "Tenant name is required."
    }

    if ([string]::IsNullOrWhiteSpace($Library)) {
        throw "Library name is required."
    }

    $tenantClean = $Tenant.Trim().ToLowerInvariant()
    $libraryClean = $Library.Trim("/")
    $subPathClean = if ([string]::IsNullOrWhiteSpace($SubPath)) { "" } else { $SubPath.Trim("/") }

    if ($Type -eq "SharePoint") {
        $sitePathClean = if ([string]::IsNullOrWhiteSpace($SitePath)) { "" } else { $SitePath.Trim("/") }
        $siteUrl = if ([string]::IsNullOrWhiteSpace($sitePathClean)) {
            "https://$tenantClean.sharepoint.com"
        } else {
            "https://$tenantClean.sharepoint.com/$sitePathClean"
        }

        $siteRelRoot = if ([string]::IsNullOrWhiteSpace($sitePathClean)) { "" } else { "/$sitePathClean" }
        $folderRel = Join-UrlPath -Segments @($libraryClean, $subPathClean)
        $folderAbs = "/" + (Join-UrlPath -Segments @($siteRelRoot, $libraryClean, $subPathClean))

        return [pscustomobject]@{
            Type = "SharePoint"
            SiteUrl = $siteUrl
            Library = $libraryClean
            FolderRel = $folderRel
            FolderAbs = $folderAbs
            Identity = if ([string]::IsNullOrWhiteSpace($sitePathClean)) { "root" } else { $sitePathClean }
        }
    }

    if ($Type -eq "OneDrive") {
        if ([string]::IsNullOrWhiteSpace($OneDriveEmail)) {
            throw "OneDrive email is required."
        }

        $segment = Convert-EmailToOneDriveSegment -Email $OneDriveEmail
        $siteUrl = "https://$tenantClean-my.sharepoint.com/personal/$segment"
        $folderRel = Join-UrlPath -Segments @($libraryClean, $subPathClean)
        $folderAbs = "/personal/$segment/" + (Join-UrlPath -Segments @($libraryClean, $subPathClean))

        return [pscustomobject]@{
            Type = "OneDrive"
            SiteUrl = $siteUrl
            Library = $libraryClean
            FolderRel = $folderRel
            FolderAbs = $folderAbs
            Identity = $OneDriveEmail.Trim()
        }
    }

    throw "Unknown type '$Type'."
}

function Get-FolderStatsRecursive {
    param(
        [Parameter(Mandatory=$true)][string]$FolderRel,
        [Parameter(Mandatory=$true)]$Connection
    )

    $result = [pscustomobject]@{
        FileCount = 0
        FolderCount = 0
        TotalBytes = [int64]0
    }

    $files = @()
    $folders = @()

    try {
        $files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderRel -ItemType File -Connection $Connection -ErrorAction Stop
    } catch {
        Write-Log "Could not read files in '$FolderRel': $_" "ERROR"
    }

    foreach ($f in $files) {
        $result.FileCount++
        $len = 0
        if ($null -ne $f.Length) {
            $len = [int64]$f.Length
        }
        $result.TotalBytes += $len
    }

    try {
        $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderRel -ItemType Folder -Connection $Connection -ErrorAction Stop
    } catch {
        Write-Log "Could not read subfolders in '$FolderRel': $_" "ERROR"
    }

    foreach ($sub in $folders) {
        $result.FolderCount++
        $subRel = "$FolderRel/$($sub.Name)"
        $subStats = Get-FolderStatsRecursive -FolderRel $subRel -Connection $Connection
        $result.FileCount += $subStats.FileCount
        $result.FolderCount += $subStats.FolderCount
        $result.TotalBytes += $subStats.TotalBytes
    }

    return $result
}

function Ensure-DestinationFolder {
    param(
        [Parameter(Mandatory=$true)][string]$FolderAbs,
        [Parameter(Mandatory=$true)]$Connection
    )

    $siteRelative = $FolderAbs.TrimStart("/")
    Resolve-PnPFolder -SiteRelativePath $siteRelative -Connection $Connection -ErrorAction Stop | Out-Null
}

function Assert-LibraryExists {
    param(
        [Parameter(Mandatory=$true)][string]$LibraryName,
        [Parameter(Mandatory=$true)]$Connection,
        [Parameter(Mandatory=$true)][string]$RoleLabel
    )

    try {
        $null = Get-PnPList -Identity $LibraryName -Connection $Connection -ErrorAction Stop
        Write-Log "$RoleLabel library found: '$LibraryName'" "SUCCESS"
    } catch {
        throw "$RoleLabel library '$LibraryName' does not exist or is not accessible."
    }
}

function Copy-FolderRecursive {
    param(
        [Parameter(Mandatory=$true)][string]$SrcFolderRelUrl,
        [Parameter(Mandatory=$true)][string]$DstFolderAbsUrl,
        [Parameter(Mandatory=$true)]$SourceConnection,
        [Parameter(Mandatory=$true)]$DestinationConnection,
        [Parameter(Mandatory=$true)][string]$DestinationSiteUrl,
        [Parameter(Mandatory=$true)][bool]$CrossSite,
        [Parameter(Mandatory=$true)][bool]$WhatIfMode,
        [Parameter(Mandatory=$true)][ref]$Counters
    )

    $files = @()
    try {
        $files = Get-PnPFolderItem -FolderSiteRelativeUrl $SrcFolderRelUrl -ItemType File -Connection $SourceConnection -ErrorAction Stop
    } catch {
        Write-Log "Could not read files in '$SrcFolderRelUrl': $_" "ERROR"
    }

    foreach ($file in $files) {
        $fileName = $file.Name
        $srcPath = "$SrcFolderRelUrl/$fileName"
        $dstPath = "$DstFolderAbsUrl/$fileName"

        if ($WhatIfMode) {
            Write-Log "[WHATIF] Copy file: $srcPath -> $dstPath" "WARN"
            $Counters.Value.Copied++
            continue
        }

        try {
            $copyParams = @{
                SourceUrl = $srcPath
                TargetUrl = $dstPath
                Force = $true
                OverwriteIfAlreadyExists = $true
                Connection = $SourceConnection
                ErrorAction = "Stop"
            }

            if ($CrossSite) {
                $copyParams["TargetWebUrl"] = $DestinationSiteUrl
            }

            Copy-PnPFile @copyParams
            Write-Log "Copied: $srcPath" "SUCCESS"
            $Counters.Value.Copied++
        } catch {
            Write-Log "Copy failed: $srcPath - $_" "ERROR"
            $Counters.Value.Failed++
        }
    }

    $subFolders = @()
    try {
        $subFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $SrcFolderRelUrl -ItemType Folder -Connection $SourceConnection -ErrorAction Stop
    } catch {
        Write-Log "Could not read subfolders in '$SrcFolderRelUrl': $_" "ERROR"
    }

    foreach ($sub in $subFolders) {
        $subSrc = "$SrcFolderRelUrl/$($sub.Name)"
        $subDst = "$DstFolderAbsUrl/$($sub.Name)"

        if ($WhatIfMode) {
            Write-Log "[WHATIF] Create/check folder: $subDst" "WARN"
        } else {
            try {
                Ensure-DestinationFolder -FolderAbs $subDst -Connection $DestinationConnection
            } catch {
                Write-Log "Could not create/check destination folder '$subDst': $_" "WARN"
            }
        }

        Copy-FolderRecursive `
            -SrcFolderRelUrl $subSrc `
            -DstFolderAbsUrl $subDst `
            -SourceConnection $SourceConnection `
            -DestinationConnection $DestinationConnection `
            -DestinationSiteUrl $DestinationSiteUrl `
            -CrossSite $CrossSite `
            -WhatIfMode $WhatIfMode `
            -Counters $Counters
    }
}

function Invoke-Migration {
    param(
        [Parameter(Mandatory=$true)]$Source,
        [Parameter(Mandatory=$true)]$Destination,
        [Parameter(Mandatory=$true)][bool]$MoveMode,
        [Parameter(Mandatory=$true)][bool]$WhatIfMode
    )

    $script:LogFile = "$PSScriptRoot\MS365Mover_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $script:FileLoggingEnabled = $true

    Ensure-PnPModule

    Write-Log "Starting migration"
    Write-Log "Source: $($Source.SiteUrl) | $($Source.FolderRel)"
    Write-Log "Destination: $($Destination.SiteUrl) | $($Destination.FolderAbs)"
    Write-Log ("Action: " + (if ($MoveMode) { "Move" } else { "Copy" }))
    if ($WhatIfMode) { Write-Log "WhatIf mode is active." "WARN" }

    $srcConn = $null
    $dstConn = $null
    $sameSite = ($Source.SiteUrl.TrimEnd("/") -ieq $Destination.SiteUrl.TrimEnd("/"))

    try {
        Write-Log "Signing in to SOURCE..."
        $srcConn = Connect-PnPOnline -Url $Source.SiteUrl -Interactive -ReturnConnection -ErrorAction Stop
        Write-Log "Signed in to SOURCE." "SUCCESS"

        if ($sameSite) {
            $dstConn = $srcConn
            Write-Log "Source and destination are on the same site."
        } else {
            Write-Log "Signing in to DESTINATION..."
            $dstConn = Connect-PnPOnline -Url $Destination.SiteUrl -Interactive -ReturnConnection -ErrorAction Stop
            Write-Log "Signed in to DESTINATION." "SUCCESS"
        }

        Assert-LibraryExists -LibraryName $Source.Library -Connection $srcConn -RoleLabel "Source"
        Assert-LibraryExists -LibraryName $Destination.Library -Connection $dstConn -RoleLabel "Destination"

        $stats = Get-FolderStatsRecursive -FolderRel $Source.FolderRel -Connection $srcConn
        $sizeGb = [math]::Round($stats.TotalBytes / 1GB, 3)
        Write-Log "Source overview: $($stats.FileCount) files, $($stats.FolderCount) folders, $sizeGb GB"

        if (-not $WhatIfMode) {
            Ensure-DestinationFolder -FolderAbs $Destination.FolderAbs -Connection $dstConn
        }

        $topFiles = @()
        $topFolders = @()

        try {
            $topFiles = Get-PnPFolderItem -FolderSiteRelativeUrl $Source.FolderRel -ItemType File -Connection $srcConn -ErrorAction Stop
            $topFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $Source.FolderRel -ItemType Folder -Connection $srcConn -ErrorAction Stop
        } catch {
            throw "Could not retrieve top-level content: $_"
        }

        $counters = [pscustomobject]@{
            Copied = 0
            Failed = 0
            Deleted = 0
        }
        $counterRef = [ref]$counters

        foreach ($file in $topFiles) {
            $srcPath = "$($Source.FolderRel)/$($file.Name)"
            $dstPath = "$($Destination.FolderAbs)/$($file.Name)"

            if ($WhatIfMode) {
                Write-Log "[WHATIF] Copy file: $srcPath -> $dstPath" "WARN"
                $counterRef.Value.Copied++
            } else {
                try {
                    $copyParams = @{
                        SourceUrl = $srcPath
                        TargetUrl = $dstPath
                        Force = $true
                        OverwriteIfAlreadyExists = $true
                        Connection = $srcConn
                        ErrorAction = "Stop"
                    }

                    if (-not $sameSite) {
                        $copyParams["TargetWebUrl"] = $Destination.SiteUrl
                    }

                    Copy-PnPFile @copyParams
                    Write-Log "Copied: $srcPath" "SUCCESS"
                    $counterRef.Value.Copied++
                } catch {
                    Write-Log "Copy failed: $srcPath - $_" "ERROR"
                    $counterRef.Value.Failed++
                }
            }
        }

        foreach ($folder in $topFolders) {
            $srcSub = "$($Source.FolderRel)/$($folder.Name)"
            $dstSub = "$($Destination.FolderAbs)/$($folder.Name)"

            if ($WhatIfMode) {
                Write-Log "[WHATIF] Create/check folder: $dstSub" "WARN"
            } else {
                try {
                    Ensure-DestinationFolder -FolderAbs $dstSub -Connection $dstConn
                } catch {
                    Write-Log "Could not create/check destination folder '$dstSub': $_" "WARN"
                }
            }

            Copy-FolderRecursive `
                -SrcFolderRelUrl $srcSub `
                -DstFolderAbsUrl $dstSub `
                -SourceConnection $srcConn `
                -DestinationConnection $dstConn `
                -DestinationSiteUrl $Destination.SiteUrl `
                -CrossSite (-not $sameSite) `
                -WhatIfMode $WhatIfMode `
                -Counters $counterRef

            if ($MoveMode -and -not $WhatIfMode) {
                try {
                    Remove-PnPFolder -Name $folder.Name -Folder $Source.FolderRel -Force -Connection $srcConn -ErrorAction Stop
                    Write-Log "Source folder removed: $($folder.Name)" "SUCCESS"
                    $counterRef.Value.Deleted++
                } catch {
                    Write-Log "Could not remove source folder '$($folder.Name)': $_" "WARN"
                }
            }
        }

        if ($MoveMode -and -not $WhatIfMode) {
            foreach ($file in $topFiles) {
                $srcPath = "$($Source.FolderRel)/$($file.Name)"
                try {
                    Remove-PnPFile -ServerRelativeUrl $srcPath -Force -Connection $srcConn -ErrorAction Stop
                    Write-Log "Source file removed: $srcPath" "SUCCESS"
                    $counterRef.Value.Deleted++
                } catch {
                    Write-Log "Could not remove source file '$srcPath': $_" "WARN"
                }
            }
        }

        Write-Log "========================================"
        Write-Log "Migration completed"
        Write-Log "Copied    : $($counterRef.Value.Copied)"
        Write-Log "Failed    : $($counterRef.Value.Failed)"
        Write-Log "Removed   : $($counterRef.Value.Deleted)"
        Write-Log "Log file  : $script:LogFile"
        Write-Log "========================================"

        return [pscustomobject]@{
            Success = ($counterRef.Value.Failed -eq 0)
            Copied = $counterRef.Value.Copied
            Failed = $counterRef.Value.Failed
            Deleted = $counterRef.Value.Deleted
        }
    } finally {
        $script:FileLoggingEnabled = $false

        if ($null -ne $srcConn) {
            Disconnect-PnPOnline -Connection $srcConn -ErrorAction SilentlyContinue
        }
        if ($null -ne $dstConn -and (-not $sameSite)) {
            Disconnect-PnPOnline -Connection $dstConn -ErrorAction SilentlyContinue
        }
    }
}

function Get-EndpointFromUi {
    param(
        [string]$Tenant,
        [System.Windows.Forms.ComboBox]$TypeCombo,
        [System.Windows.Forms.TextBox]$SitePathBox,
        [System.Windows.Forms.TextBox]$LibraryBox,
        [System.Windows.Forms.TextBox]$SubPathBox,
        [System.Windows.Forms.TextBox]$EmailBox
    )

    $type = [string]$TypeCombo.SelectedItem
    return Build-Endpoint `
        -Tenant $Tenant `
        -Type $type `
        -SitePath $SitePathBox.Text `
        -Library $LibraryBox.Text `
        -SubPath $SubPathBox.Text `
        -OneDriveEmail $EmailBox.Text
}

function Update-TypeUi {
    param(
        [System.Windows.Forms.ComboBox]$TypeCombo,
        [System.Windows.Forms.TextBox]$SitePathBox,
        [System.Windows.Forms.TextBox]$EmailBox,
        [System.Windows.Forms.TextBox]$LibraryBox
    )

    $isOneDrive = ([string]$TypeCombo.SelectedItem -eq "OneDrive")
    $SitePathBox.Enabled = -not $isOneDrive
    $EmailBox.Enabled = $isOneDrive

    if ($isOneDrive) {
        if ([string]::IsNullOrWhiteSpace($LibraryBox.Text) -or $LibraryBox.Text -eq "Shared Documents") {
            $LibraryBox.Text = "Documents"
        }
    } else {
        if ([string]::IsNullOrWhiteSpace($LibraryBox.Text) -or $LibraryBox.Text -eq "Documents") {
            $LibraryBox.Text = "Shared Documents"
        }
    }
}

function Update-TenantExamples {
    param(
        [System.Windows.Forms.TextBox]$TenantBox,
        [System.Windows.Forms.Label]$ExampleLabel
    )

    $tenant = $TenantBox.Text.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($tenant)) {
        $tenant = "<tenant>"
    }

    $ExampleLabel.Text = "Examples: SharePoint URL https://$tenant.sharepoint.com/sites/Finance | OneDrive URL https://$tenant-my.sharepoint.com/personal/john_doe_contoso_com"
}

function Update-SitePathExample {
    param(
        [System.Windows.Forms.TextBox]$TenantBox,
        [System.Windows.Forms.ComboBox]$TypeCombo,
        [System.Windows.Forms.TextBox]$SitePathBox,
        [System.Windows.Forms.Label]$ExampleLabel
    )

    $tenant = $TenantBox.Text.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($tenant)) {
        $tenant = "<tenant>"
    }

    if ([string]$TypeCombo.SelectedItem -eq "OneDrive") {
        $ExampleLabel.Text = "Example URL: https://$tenant-my.sharepoint.com/personal/john_doe_contoso_com"
        return
    }

    $enteredPath = $SitePathBox.Text.Trim().Trim("/")
    $defaultPath = "sites/Finance"
    if ([string]::IsNullOrWhiteSpace($enteredPath)) {
        $enteredPath = $defaultPath
    }

    $ExampleLabel.Text = "Example URL: https://$tenant.sharepoint.com/$enteredPath"
}

function Update-LibrarySubfolderExamples {
    param(
        [System.Windows.Forms.TextBox]$TenantBox,
        [System.Windows.Forms.ComboBox]$TypeCombo,
        [System.Windows.Forms.TextBox]$SitePathBox,
        [System.Windows.Forms.TextBox]$LibraryBox,
        [System.Windows.Forms.TextBox]$SubPathBox,
        [System.Windows.Forms.Label]$LibraryExampleLabel,
        [System.Windows.Forms.Label]$SubfolderExampleLabel
    )

    $tenant = $TenantBox.Text.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($tenant)) {
        $tenant = "<tenant>"
    }

    $library = $LibraryBox.Text.Trim().Trim("/")
    if ([string]::IsNullOrWhiteSpace($library)) {
        if ([string]$TypeCombo.SelectedItem -eq "OneDrive") {
            $library = "Documents"
        } else {
            $library = "Shared Documents"
        }
    }

    $subPath = $SubPathBox.Text.Trim().Trim("/")
    if ([string]::IsNullOrWhiteSpace($subPath)) {
        $subPath = "ProjectFolder"
    }

    if ([string]$TypeCombo.SelectedItem -eq "OneDrive") {
        $baseUrl = "https://$tenant-my.sharepoint.com/personal/john_doe_contoso_com"
        $LibraryExampleLabel.Text = "Example URL: $baseUrl/$library"
        $SubfolderExampleLabel.Text = "Example URL: $baseUrl/$library/$subPath"
        return
    }

    $sitePath = $SitePathBox.Text.Trim().Trim("/")
    if ([string]::IsNullOrWhiteSpace($sitePath)) {
        $sitePath = "sites/Finance"
    }

    $baseUrl = "https://$tenant.sharepoint.com/$sitePath"
    $LibraryExampleLabel.Text = "Example URL: $baseUrl/$library"
    $SubfolderExampleLabel.Text = "Example URL: $baseUrl/$library/$subPath"
}

# ---------------- GUI ----------------
Ensure-RequiredPowerShell

$form = New-Object System.Windows.Forms.Form
$form.Text = "MS365 File Mover"
$form.Size = New-Object System.Drawing.Size(980, 760)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $true

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Font = $font

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Location = New-Object System.Drawing.Point(20, 20)
$lblTenant.Size = New-Object System.Drawing.Size(220, 25)
$lblTenant.Text = "Tenant:"
$form.Controls.Add($lblTenant)

$txtTenant = New-Object System.Windows.Forms.TextBox
$txtTenant.Location = New-Object System.Drawing.Point(250, 18)
$txtTenant.Size = New-Object System.Drawing.Size(260, 25)
$form.Controls.Add($txtTenant)

$lblTenantExample = New-Object System.Windows.Forms.Label
$lblTenantExample.Location = New-Object System.Drawing.Point(20, 48)
$lblTenantExample.Size = New-Object System.Drawing.Size(930, 30)
$lblTenantExample.Text = "Examples: SharePoint URL https://<tenant>.sharepoint.com/sites/Finance | OneDrive URL https://<tenant>-my.sharepoint.com/personal/john_doe_contoso_com"
$form.Controls.Add($lblTenantExample)

$grpSource = New-Object System.Windows.Forms.GroupBox
$grpSource.Text = "Source"
$grpSource.Location = New-Object System.Drawing.Point(20, 85)
$grpSource.Size = New-Object System.Drawing.Size(450, 295)
$form.Controls.Add($grpSource)

$srcLblType = New-Object System.Windows.Forms.Label
$srcLblType.Location = New-Object System.Drawing.Point(20, 30)
$srcLblType.Size = New-Object System.Drawing.Size(120, 25)
$srcLblType.Text = "Type:"
$grpSource.Controls.Add($srcLblType)

$srcType = New-Object System.Windows.Forms.ComboBox
$srcType.Location = New-Object System.Drawing.Point(150, 28)
$srcType.Size = New-Object System.Drawing.Size(260, 25)
$srcType.DropDownStyle = "DropDownList"
[void]$srcType.Items.AddRange(@("SharePoint", "OneDrive"))
$srcType.SelectedIndex = 0
$grpSource.Controls.Add($srcType)

$srcLblSite = New-Object System.Windows.Forms.Label
$srcLblSite.Location = New-Object System.Drawing.Point(20, 65)
$srcLblSite.Size = New-Object System.Drawing.Size(120, 25)
$srcLblSite.Text = "Site path:"
$grpSource.Controls.Add($srcLblSite)

$srcSitePath = New-Object System.Windows.Forms.TextBox
$srcSitePath.Location = New-Object System.Drawing.Point(150, 63)
$srcSitePath.Size = New-Object System.Drawing.Size(260, 25)
$srcSitePath.Text = ""
$grpSource.Controls.Add($srcSitePath)

$srcSiteExample = New-Object System.Windows.Forms.Label
$srcSiteExample.Location = New-Object System.Drawing.Point(150, 88)
$srcSiteExample.Size = New-Object System.Drawing.Size(280, 25)
$srcSiteExample.Text = "Example URL:"
$grpSource.Controls.Add($srcSiteExample)

$srcLblEmail = New-Object System.Windows.Forms.Label
$srcLblEmail.Location = New-Object System.Drawing.Point(20, 120)
$srcLblEmail.Size = New-Object System.Drawing.Size(120, 25)
$srcLblEmail.Text = "OneDrive email:"
$grpSource.Controls.Add($srcLblEmail)

$srcEmail = New-Object System.Windows.Forms.TextBox
$srcEmail.Location = New-Object System.Drawing.Point(150, 118)
$srcEmail.Size = New-Object System.Drawing.Size(260, 25)
$srcEmail.Enabled = $false
$grpSource.Controls.Add($srcEmail)

$srcLblLibrary = New-Object System.Windows.Forms.Label
$srcLblLibrary.Location = New-Object System.Drawing.Point(20, 155)
$srcLblLibrary.Size = New-Object System.Drawing.Size(120, 25)
$srcLblLibrary.Text = "Library:"
$grpSource.Controls.Add($srcLblLibrary)

$srcLibrary = New-Object System.Windows.Forms.TextBox
$srcLibrary.Location = New-Object System.Drawing.Point(150, 153)
$srcLibrary.Size = New-Object System.Drawing.Size(260, 25)
$srcLibrary.Text = "Shared Documents"
$grpSource.Controls.Add($srcLibrary)

$srcLibraryExample = New-Object System.Windows.Forms.Label
$srcLibraryExample.Location = New-Object System.Drawing.Point(150, 178)
$srcLibraryExample.Size = New-Object System.Drawing.Size(280, 25)
$srcLibraryExample.Text = "Example URL:"
$grpSource.Controls.Add($srcLibraryExample)

$srcLblSub = New-Object System.Windows.Forms.Label
$srcLblSub.Location = New-Object System.Drawing.Point(20, 215)
$srcLblSub.Size = New-Object System.Drawing.Size(120, 25)
$srcLblSub.Text = "Subfolder (opt.):"
$grpSource.Controls.Add($srcLblSub)

$srcSubPath = New-Object System.Windows.Forms.TextBox
$srcSubPath.Location = New-Object System.Drawing.Point(150, 213)
$srcSubPath.Size = New-Object System.Drawing.Size(260, 25)
$srcSubPath.Text = ""
$grpSource.Controls.Add($srcSubPath)

$srcSubExample = New-Object System.Windows.Forms.Label
$srcSubExample.Location = New-Object System.Drawing.Point(150, 238)
$srcSubExample.Size = New-Object System.Drawing.Size(280, 45)
$srcSubExample.Text = "Example URL:"
$grpSource.Controls.Add($srcSubExample)

$grpDest = New-Object System.Windows.Forms.GroupBox
$grpDest.Text = "Destination"
$grpDest.Location = New-Object System.Drawing.Point(500, 85)
$grpDest.Size = New-Object System.Drawing.Size(450, 295)
$form.Controls.Add($grpDest)

$dstLblType = New-Object System.Windows.Forms.Label
$dstLblType.Location = New-Object System.Drawing.Point(20, 30)
$dstLblType.Size = New-Object System.Drawing.Size(120, 25)
$dstLblType.Text = "Type:"
$grpDest.Controls.Add($dstLblType)

$dstType = New-Object System.Windows.Forms.ComboBox
$dstType.Location = New-Object System.Drawing.Point(150, 28)
$dstType.Size = New-Object System.Drawing.Size(260, 25)
$dstType.DropDownStyle = "DropDownList"
[void]$dstType.Items.AddRange(@("SharePoint", "OneDrive"))
$dstType.SelectedIndex = 0
$grpDest.Controls.Add($dstType)

$dstLblSite = New-Object System.Windows.Forms.Label
$dstLblSite.Location = New-Object System.Drawing.Point(20, 65)
$dstLblSite.Size = New-Object System.Drawing.Size(120, 25)
$dstLblSite.Text = "Site path:"
$grpDest.Controls.Add($dstLblSite)

$dstSitePath = New-Object System.Windows.Forms.TextBox
$dstSitePath.Location = New-Object System.Drawing.Point(150, 63)
$dstSitePath.Size = New-Object System.Drawing.Size(260, 25)
$dstSitePath.Text = ""
$grpDest.Controls.Add($dstSitePath)

$dstSiteExample = New-Object System.Windows.Forms.Label
$dstSiteExample.Location = New-Object System.Drawing.Point(150, 88)
$dstSiteExample.Size = New-Object System.Drawing.Size(280, 25)
$dstSiteExample.Text = "Example URL:"
$grpDest.Controls.Add($dstSiteExample)

$dstLblEmail = New-Object System.Windows.Forms.Label
$dstLblEmail.Location = New-Object System.Drawing.Point(20, 120)
$dstLblEmail.Size = New-Object System.Drawing.Size(120, 25)
$dstLblEmail.Text = "OneDrive email:"
$grpDest.Controls.Add($dstLblEmail)

$dstEmail = New-Object System.Windows.Forms.TextBox
$dstEmail.Location = New-Object System.Drawing.Point(150, 118)
$dstEmail.Size = New-Object System.Drawing.Size(260, 25)
$dstEmail.Enabled = $false
$grpDest.Controls.Add($dstEmail)

$dstLblLibrary = New-Object System.Windows.Forms.Label
$dstLblLibrary.Location = New-Object System.Drawing.Point(20, 155)
$dstLblLibrary.Size = New-Object System.Drawing.Size(120, 25)
$dstLblLibrary.Text = "Library:"
$grpDest.Controls.Add($dstLblLibrary)

$dstLibrary = New-Object System.Windows.Forms.TextBox
$dstLibrary.Location = New-Object System.Drawing.Point(150, 153)
$dstLibrary.Size = New-Object System.Drawing.Size(260, 25)
$dstLibrary.Text = "Shared Documents"
$grpDest.Controls.Add($dstLibrary)

$dstLibraryExample = New-Object System.Windows.Forms.Label
$dstLibraryExample.Location = New-Object System.Drawing.Point(150, 178)
$dstLibraryExample.Size = New-Object System.Drawing.Size(280, 25)
$dstLibraryExample.Text = "Example URL:"
$grpDest.Controls.Add($dstLibraryExample)

$dstLblSub = New-Object System.Windows.Forms.Label
$dstLblSub.Location = New-Object System.Drawing.Point(20, 215)
$dstLblSub.Size = New-Object System.Drawing.Size(120, 25)
$dstLblSub.Text = "Subfolder (opt.):"
$grpDest.Controls.Add($dstLblSub)

$dstSubPath = New-Object System.Windows.Forms.TextBox
$dstSubPath.Location = New-Object System.Drawing.Point(150, 213)
$dstSubPath.Size = New-Object System.Drawing.Size(260, 25)
$dstSubPath.Text = ""
$grpDest.Controls.Add($dstSubPath)

$dstSubExample = New-Object System.Windows.Forms.Label
$dstSubExample.Location = New-Object System.Drawing.Point(150, 238)
$dstSubExample.Size = New-Object System.Drawing.Size(280, 45)
$dstSubExample.Text = "Example URL:"
$grpDest.Controls.Add($dstSubExample)

$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text = "Options"
$grpOptions.Location = New-Object System.Drawing.Point(20, 400)
$grpOptions.Size = New-Object System.Drawing.Size(930, 90)
$form.Controls.Add($grpOptions)

$lblAction = New-Object System.Windows.Forms.Label
$lblAction.Location = New-Object System.Drawing.Point(20, 38)
$lblAction.Size = New-Object System.Drawing.Size(60, 25)
$lblAction.Text = "Action:"
$grpOptions.Controls.Add($lblAction)

$cmbAction = New-Object System.Windows.Forms.ComboBox
$cmbAction.Location = New-Object System.Drawing.Point(85, 35)
$cmbAction.Size = New-Object System.Drawing.Size(140, 25)
$cmbAction.DropDownStyle = "DropDownList"
[void]$cmbAction.Items.AddRange(@("Copy", "Move"))
$cmbAction.SelectedIndex = 0
$grpOptions.Controls.Add($cmbAction)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Location = New-Object System.Drawing.Point(250, 37)
$chkWhatIf.Size = New-Object System.Drawing.Size(160, 25)
$chkWhatIf.Text = "WhatIf (no changes)"
$grpOptions.Controls.Add($chkWhatIf)

$btnFetch = New-Object System.Windows.Forms.Button
$btnFetch.Location = New-Object System.Drawing.Point(470, 30)
$btnFetch.Size = New-Object System.Drawing.Size(140, 35)
$btnFetch.Text = "Get overview"
$grpOptions.Controls.Add($btnFetch)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Location = New-Object System.Drawing.Point(620, 30)
$btnExportCsv.Size = New-Object System.Drawing.Size(130, 35)
$btnExportCsv.Text = "Export CSV"
$grpOptions.Controls.Add($btnExportCsv)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(770, 30)
$btnRun.Size = New-Object System.Drawing.Size(130, 35)
$btnRun.Text = "Start"
$grpOptions.Controls.Add($btnRun)

$txtSummary = New-Object System.Windows.Forms.TextBox
$txtSummary.Location = New-Object System.Drawing.Point(20, 500)
$txtSummary.Size = New-Object System.Drawing.Size(930, 50)
$txtSummary.Multiline = $true
$txtSummary.ReadOnly = $true
$txtSummary.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($txtSummary)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 560)
$txtLog.Size = New-Object System.Drawing.Size(930, 190)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($txtLog)

$script:LogTextBox = $txtLog

$txtTenant.Add_TextChanged({
    Update-TenantExamples -TenantBox $txtTenant -ExampleLabel $lblTenantExample
    Update-SitePathExample -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -ExampleLabel $srcSiteExample
    Update-SitePathExample -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -ExampleLabel $dstSiteExample
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -LibraryBox $srcLibrary -SubPathBox $srcSubPath -LibraryExampleLabel $srcLibraryExample -SubfolderExampleLabel $srcSubExample
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -LibraryBox $dstLibrary -SubPathBox $dstSubPath -LibraryExampleLabel $dstLibraryExample -SubfolderExampleLabel $dstSubExample
})

$srcSitePath.Add_TextChanged({
    Update-SitePathExample -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -ExampleLabel $srcSiteExample
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -LibraryBox $srcLibrary -SubPathBox $srcSubPath -LibraryExampleLabel $srcLibraryExample -SubfolderExampleLabel $srcSubExample
})

$dstSitePath.Add_TextChanged({
    Update-SitePathExample -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -ExampleLabel $dstSiteExample
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -LibraryBox $dstLibrary -SubPathBox $dstSubPath -LibraryExampleLabel $dstLibraryExample -SubfolderExampleLabel $dstSubExample
})

$srcLibrary.Add_TextChanged({
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -LibraryBox $srcLibrary -SubPathBox $srcSubPath -LibraryExampleLabel $srcLibraryExample -SubfolderExampleLabel $srcSubExample
})

$dstLibrary.Add_TextChanged({
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -LibraryBox $dstLibrary -SubPathBox $dstSubPath -LibraryExampleLabel $dstLibraryExample -SubfolderExampleLabel $dstSubExample
})

$srcSubPath.Add_TextChanged({
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -LibraryBox $srcLibrary -SubPathBox $srcSubPath -LibraryExampleLabel $srcLibraryExample -SubfolderExampleLabel $srcSubExample
})

$dstSubPath.Add_TextChanged({
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -LibraryBox $dstLibrary -SubPathBox $dstSubPath -LibraryExampleLabel $dstLibraryExample -SubfolderExampleLabel $dstSubExample
})

$srcType.Add_SelectedIndexChanged({
    Update-TypeUi -TypeCombo $srcType -SitePathBox $srcSitePath -EmailBox $srcEmail -LibraryBox $srcLibrary
    Update-SitePathExample -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -ExampleLabel $srcSiteExample
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -LibraryBox $srcLibrary -SubPathBox $srcSubPath -LibraryExampleLabel $srcLibraryExample -SubfolderExampleLabel $srcSubExample
})

$dstType.Add_SelectedIndexChanged({
    Update-TypeUi -TypeCombo $dstType -SitePathBox $dstSitePath -EmailBox $dstEmail -LibraryBox $dstLibrary
    Update-SitePathExample -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -ExampleLabel $dstSiteExample
    Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -LibraryBox $dstLibrary -SubPathBox $dstSubPath -LibraryExampleLabel $dstLibraryExample -SubfolderExampleLabel $dstSubExample
})

$btnFetch.Add_Click({
    try {
        $tenant = $txtTenant.Text.Trim()
        $source = Get-EndpointFromUi -Tenant $tenant -TypeCombo $srcType -SitePathBox $srcSitePath -LibraryBox $srcLibrary -SubPathBox $srcSubPath -EmailBox $srcEmail

        Ensure-PnPModule
        Write-Log "Fetching source overview..."
        $srcConn = Connect-PnPOnline -Url $source.SiteUrl -Interactive -ReturnConnection -ErrorAction Stop

        try {
            Assert-LibraryExists -LibraryName $source.Library -Connection $srcConn -RoleLabel "Source"
            $stats = Get-FolderStatsRecursive -FolderRel $source.FolderRel -Connection $srcConn
            $gb = [math]::Round($stats.TotalBytes / 1GB, 3)
            $txtSummary.Text = "Source: $($source.SiteUrl) | Path: $($source.FolderRel) | Files: $($stats.FileCount) | Folders: $($stats.FolderCount) | Size: $gb GB"

            $script:LastOverview = [pscustomobject]@{
                Timestamp = Get-Date
                Tenant = $tenant
                SourceType = $source.Type
                SourceSiteUrl = $source.SiteUrl
                SourceLibrary = $source.Library
                SourceFolderRel = $source.FolderRel
                FileCount = $stats.FileCount
                FolderCount = $stats.FolderCount
                TotalBytes = $stats.TotalBytes
                TotalGB = $gb
            }

            Write-Log "Overview ready: $($stats.FileCount) files, $($stats.FolderCount) folders, $gb GB" "SUCCESS"
        } finally {
            Disconnect-PnPOnline -Connection $srcConn -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Log "Failed to fetch overview: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Failed to fetch overview: $_", "Error", "OK", "Error") | Out-Null
    }
})

$btnExportCsv.Add_Click({
    try {
        if ($null -eq $script:LastOverview) {
            throw "No overview has been fetched yet. Click 'Get overview' first."
        }

        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Title = "Save as CSV"
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "MS365Mover_Overview_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $saveDialog.InitialDirectory = $PSScriptRoot

        $dialogResult = $saveDialog.ShowDialog()
        if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        $script:LastOverview | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
        Write-Log "Overview exported to: $($saveDialog.FileName)" "SUCCESS"
        [System.Windows.Forms.MessageBox]::Show("CSV saved:\n$($saveDialog.FileName)", "Export complete", "OK", "Information") | Out-Null
    } catch {
        Write-Log "CSV export failed: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("CSV export failed: $_", "Error", "OK", "Error") | Out-Null
    }
})

$btnRun.Add_Click({
    try {
        $tenant = $txtTenant.Text.Trim()
        $source = Get-EndpointFromUi -Tenant $tenant -TypeCombo $srcType -SitePathBox $srcSitePath -LibraryBox $srcLibrary -SubPathBox $srcSubPath -EmailBox $srcEmail
        $dest = Get-EndpointFromUi -Tenant $tenant -TypeCombo $dstType -SitePathBox $dstSitePath -LibraryBox $dstLibrary -SubPathBox $dstSubPath -EmailBox $dstEmail

        $moveMode = ([string]$cmbAction.SelectedItem -eq "Move")
        $whatIfMode = $chkWhatIf.Checked

        if ($source.SiteUrl.TrimEnd("/") -ieq $dest.SiteUrl.TrimEnd("/") -and $source.FolderAbs.TrimEnd("/") -ieq $dest.FolderAbs.TrimEnd("/")) {
            throw "Source and destination are identical. Choose a different destination path."
        }

        $confirmText = "Action: " + (if ($moveMode) { "Move" } else { "Copy" }) + "`n`nSource:`n$($source.SiteUrl)`n$($source.FolderRel)`n`nDestination:`n$($dest.SiteUrl)`n$($dest.FolderAbs)"
        $confirm = [System.Windows.Forms.MessageBox]::Show($confirmText, "Confirm", "OKCancel", "Question")
        if ($confirm -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "Action canceled by user." "WARN"
            return
        }

        $result = Invoke-Migration -Source $source -Destination $dest -MoveMode $moveMode -WhatIfMode $whatIfMode

        if ($result.Success) {
            [System.Windows.Forms.MessageBox]::Show("Done. Everything processed.", "Completed", "OK", "Information") | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("Done with errors. Check the log file.", "Completed with warnings", "OK", "Warning") | Out-Null
        }
    } catch {
        Write-Log "Migration failed: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Migration failed: $_", "Error", "OK", "Error") | Out-Null
    }
})

Update-TypeUi -TypeCombo $srcType -SitePathBox $srcSitePath -EmailBox $srcEmail -LibraryBox $srcLibrary
Update-TypeUi -TypeCombo $dstType -SitePathBox $dstSitePath -EmailBox $dstEmail -LibraryBox $dstLibrary
Update-TenantExamples -TenantBox $txtTenant -ExampleLabel $lblTenantExample
Update-SitePathExample -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -ExampleLabel $srcSiteExample
Update-SitePathExample -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -ExampleLabel $dstSiteExample
Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $srcType -SitePathBox $srcSitePath -LibraryBox $srcLibrary -SubPathBox $srcSubPath -LibraryExampleLabel $srcLibraryExample -SubfolderExampleLabel $srcSubExample
Update-LibrarySubfolderExamples -TenantBox $txtTenant -TypeCombo $dstType -SitePathBox $dstSitePath -LibraryBox $dstLibrary -SubPathBox $dstSubPath -LibraryExampleLabel $dstLibraryExample -SubfolderExampleLabel $dstSubExample

Write-Log "GUI started. Fill tenant + source/destination and click 'Get overview' first."
[void]$form.ShowDialog()
