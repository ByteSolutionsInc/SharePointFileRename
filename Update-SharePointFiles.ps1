<#
    ==========================
    Configuration (edit here)
    ==========================

    Prerequisites (App-only auth):
    - PowerShell 7+ (recommended).
    - PnP.PowerShell module installed:
        Install-Module PnP.PowerShell -Scope CurrentUser
    - Create an app registration in Microsoft Entra ID.
    - Upload the public cert (or use thumbprint from local store).
    - Grant admin consent for SharePoint permissions (application):
        * Sites.FullControl.All (or least-privilege equivalent).
    - In SharePoint Admin Center, ensure the app is allowed (if required).

        Required values:
        - TenantId: Entra tenant GUID (example only)
        - ClientId: App (client) ID (example only)
        - CertThumbprint: thumbprint of cert in CurrentUser\My (example only)
            OR CertPath + CertPassword (PFX)

        Site list creation and review:
        - Use Get-TeamsSiteUrls to build a list of sites.
        - Review and edit the list before processing.
        - Example flow:
                $siteUrls = Get-TeamsSiteUrls -AllSites -IncludeNoAccess:$false
                # Review and edit $siteUrls before running
                Invoke-RemoveAkiraSuffix -SiteUrls $siteUrls -Libraries $LibrariesToUse -Connection $conn -Verbose

    Certificate creation (PowerShell, local personal store):
    # Create a self-signed cert and export PFX + CER
    $cert = New-SelfSignedCertificate `
        -Subject "CN=PnP-AppOnly" `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -NotAfter (Get-Date).AddYears(2)

    $pfxPath = Join-Path $env:USERPROFILE "Desktop\PnP-AppOnly.pfx"
    $cerPath = Join-Path $env:USERPROFILE "Desktop\PnP-AppOnly.cer"
    $pfxPassword = Read-Host -AsSecureString "PFX Password"

    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword | Out-Null
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

    # Upload the CER to the app registration (Certificates & secrets).
    # Use the thumbprint from $cert.Thumbprint for CertThumbprint below.

    Notes:
    - Keep the PFX secure; treat as a secret.
    - If you change the cert, update the app registration and thumbprint.

    Disclaimer:
    - Provided “as is” without warranties or guarantees of any kind.
    - You assume all risk and liability for use, including data loss or service impact.
#>
$TenantId = "00000000-0000-0000-0000-000000000000"   # TODO: replace with your tenant ID
$ClientId = "00000000-0000-0000-0000-000000000000"   # TODO: replace with your app (client) ID
$CertThumbprint = "0000000000000000000000000000000000000000" # TODO: replace with your cert thumbprint
$CertPath = ""
$CertPassword = $null

# Logging
$RunStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$LogPath = Join-Path -Path (Get-Location) -ChildPath "AkiraCleanup-$RunStamp.log"

# Optional: limit to a smaller test library list (leave empty to process all libraries)
$LibrariesToUse = @("Shared Documents")

# Admin Center URL
$AdminUrl = "https://contoso-admin.sharepoint.com"    # TODO: replace with your admin center URL

# Site URL lists
$siteUrls = @(
    "https://contoso.sharepoint.com/sites/Finance",
    "https://contoso.sharepoint.com/sites/Marketing",
    "https://contoso.sharepoint.com/sites/Operations"
)

# Single-site test list
$siteUrl = @('https://contoso.sharepoint.com/sites/Marketing')

function Test-PnPConnection {
    <#
        .SYNOPSIS
            Validates that a PnP connection/context exists.
        .OUTPUTS
            [bool] $true if connected; $false otherwise.
    #>
    param(
        $Connection
    )
    try {
        # Lightweight call that requires a context; throws if missing
        if ($Connection) {
            $null = Get-PnPContext -Connection $Connection -ErrorAction Stop
        } else {
            $null = Get-PnPContext -ErrorAction Stop
        }
        return $true
    } catch {
        return $false
    }
}

function Connect-PnPTargetSite {
    <#
        .SYNOPSIS
            Connects to a SharePoint site using app-only auth, with a fallback for older PnP.PowerShell parameter sets.
    #>
    param(
        [Parameter(Mandatory)][string] $Url,
        [Parameter(Mandatory)][string] $ClientId,
        [string] $TenantId,
        [string] $CertThumbprint,
        [string] $CertPath,
        [securestring] $CertPassword
    )

    try {
        if ($CertThumbprint) {
            $conn = Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -Thumbprint $CertThumbprint -ReturnConnection -ErrorAction Stop
        } elseif ($CertPath) {
            if (-not $CertPassword) {
                $CertPassword = Read-Host -Prompt "Enter certificate password" -AsSecureString
            }
            $conn = Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertPath -CertificatePassword $CertPassword -ReturnConnection -ErrorAction Stop
        } else {
            throw "Certificate not provided. Set CertThumbprint or CertPath."
        }
        return $conn
    } catch {
        Write-Warning "Connect failed for '$Url': $($_.Exception.Message)"
        return $null
    }
}


function Get-TeamsSiteUrls {
    <#
        .SYNOPSIS
            Returns SharePoint site URLs for Microsoft Teams (including Private Channel sites),
            or ALL site URLs in the tenant, and merges with any KnownSiteUrls.

        .DESCRIPTION
            When -AllSites is specified, enumerates tenant sites via Get-PnPTenantSite -Detailed and
            applies optional filters (exclude OneDrive, admin/system, locked/archived/no-access).
            When Teams are provided, resolves Team->GroupId and queries SPO Search for sites where
            DepartmentId == GroupId (parent + private channel sites). Results are deduped and sorted.

        .PARAMETER Teams
            One or more Team display names (exact or unique partials). Optional.

        .PARAMETER KnownSiteUrls
            One or more fully-qualified SharePoint site URLs to include as-is. Optional.

        .PARAMETER AllSites
            When present, returns all SPO site URLs in the tenant (subject to filters).

        .PARAMETER IncludeOneDrive
            Include OneDrive (personal) sites when -AllSites is used. Default: False.

        .PARAMETER IncludeAdminAndSystem
            Include admin center, mysite host, and other system URLs when -AllSites is used. Default: False.

        .PARAMETER IncludeArchivedOrLocked
            Include archived or locked sites when -AllSites is used. Default: False.

        .PARAMETER IncludeNoAccess
            Include sites where the current identity has no access (URL only) when -AllSites is used. Default: True.

        .OUTPUTS
            [string[]] Sorted unique list of site URLs.
    #>
    [CmdletBinding()]
    param(
        [string[]] $Teams,
        [string[]] $KnownSiteUrls,
        [switch]   $AllSites,
        [switch]   $IncludeOneDrive,
        [switch]   $IncludeAdminAndSystem,
        [switch]   $IncludeArchivedOrLocked,
        [switch]   $IncludeNoAccess = $true
    )

    if (-not (Test-PnPConnection)) {
        throw "No active PnP connection. Connect to the admin center first with Connect-PnPOnline."
    }

    $urls = New-Object System.Collections.Generic.List[string]

    # -------- Teams-based discovery (parent + private channel sites) --------
    if ($Teams -and $Teams.Count -gt 0) {
        # Pull all groups once
        $allGroups = Get-PnPMicrosoft365Group -IncludeSiteUrl:$false

        foreach ($team in $Teams) {
            $grp = $allGroups | Where-Object { $_.DisplayName -eq $team } | Select-Object -First 1
            if (-not $grp) { $grp = $allGroups | Where-Object { $_.DisplayName -like "*$team*" } | Select-Object -First 1 }
            if (-not $grp) {
                Write-Warning "Team '$team' could not be resolved to a Microsoft 365 Group."
                continue
            }

            $groupId = $grp.Id
            # KQL: DepartmentId = GroupId captures Team root + Private Channel sites
            $kql = "contentclass:STS_Site DepartmentId:$groupId"
            $results = Submit-PnPSearchQuery -Query $kql -SelectProperties "SPWebUrl","SPSiteUrl","Title","SiteTemplate","DepartmentId" -All

            foreach ($row in $results.ResultRows) {
                $u = if ($row.SPWebUrl) { $row.SPWebUrl } elseif ($row.SPSiteUrl) { $row.SPSiteUrl } else { $null }
                if ($u -and $u -match '^https://') { [void]$urls.Add($u) }
            }
        }
    }

    # ------------------------ All-sites enumeration -------------------------
    if ($AllSites) {
        # Requires admin center connection and SPO admin permission
        $tenantSites = Get-PnPTenantSite -Detailed

        foreach ($s in $tenantSites) {
            $u = $s.Url
            if (-not $u) { continue }

            # Exclude admin/system unless requested
            if (-not $IncludeAdminAndSystem) {
                if ($u -match '-admin\.sharepoint\.com')        { continue }
                if ($u -match '\.mysite\.sharepoint\.com')      { continue } # mysite host (rare in modern tenants)
                if ($u -match '/portals/hub' )                  { continue } # managed metadata hub (legacy)
                if ($u -match '/search' )                       { continue }
                if ($u -match '/sites/appcatalog' )             { continue }
            }

            # Exclude OneDrive personal unless requested
            if (-not $IncludeOneDrive) {
                # OneDrive personal sites typically under https://<tenant>-my.sharepoint.com/personal/...
                if ($u -match '-my\.sharepoint\.com') { continue }
            }

            # Exclude locked/archived if not requested
            if (-not $IncludeArchivedOrLocked) {
                if ($s.LockState -and ($s.LockState -ne "Unlock")) { continue }
                if ($s.DenyAddAndCustomizePages -and ($s.DenyAddAndCustomizePages -ne "Disabled")) {
                    # informational; typically we still include
                }
                if ($s.RelatedGroupId -eq [Guid]::Empty) { # not directly used
                    # fine to include; teamsless collaboration sites
                }
                if ($s.IsTeamsConnected -and $s.TeamsChannelType -eq "Archived") { continue }
            }

            # Include even if no access (URL only) unless explicitly excluded
            if (-not $IncludeNoAccess) {
                # Attempt a soft check: skip site collections marked as NoAccess (where known)
                # PnP doesn't expose a simple "HasAccess" flag; we keep the URL unless explicitly disabled.
            }

            [void]$urls.Add($u)
        }
    }

    # ----------------------------- Known URLs -------------------------------
    if ($KnownSiteUrls) {
        foreach ($ku in $KnownSiteUrls) {
            if ($ku -and $ku -match '^https://') { [void]$urls.Add($ku) }
        }
    }

    # --------------------------- Deduplicate/sort ---------------------------
    ($urls.ToArray() | Sort-Object -Unique)
}

function Invoke-RemoveAkiraSuffix {

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string[]]$SiteUrls,

        [Parameter(Mandatory)]
        [string[]]$Libraries,

        [string]$FolderServerRelativeUrl,

        [string]$LogPath,

        [Parameter(Mandatory)]
        $Connection
    )

    if (-not $LogPath) {
        $LogPath = Join-Path -Path (Get-Location) -ChildPath ("AkiraCleanup-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
    }

    foreach ($siteUrl in $SiteUrls) {

        Write-Host ""
        Write-Host "Site: $siteUrl" -ForegroundColor Cyan

        $resolvedFolder = $null
        if ($FolderServerRelativeUrl) {
            if ($FolderServerRelativeUrl -match '^https?://') {
                try {
                    $uri = [System.Uri]$FolderServerRelativeUrl
                    if ($uri.Query -match 'id=([^&]+)') {
                        $resolvedFolder = [System.Net.WebUtility]::UrlDecode($Matches[1])
                    } else {
                        $resolvedFolder = [System.Uri]::UnescapeDataString($uri.AbsolutePath)
                    }
                } catch {
                    $resolvedFolder = $FolderServerRelativeUrl
                }
            } else {
                $resolvedFolder = $FolderServerRelativeUrl
            }
        }

        foreach ($lib in $Libraries) {

            Write-Host "  Library: $lib" -ForegroundColor Yellow
            if ($resolvedFolder) {
                Write-Host "    Folder scope: $resolvedFolder" -ForegroundColor DarkYellow
            }

            $list = Get-PnPList -Identity $lib -Connection $Connection -ErrorAction SilentlyContinue
            if (-not $list) {
                Write-Warning "  Library not found: $lib"
                continue
            }

            # Find .akira files
            $query = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <Contains>
        <FieldRef Name='FileLeafRef'/>
        <Value Type='Text'>.akira</Value>
      </Contains>
    </Where>
  </Query>
</View>
"@

            $akira = if ($resolvedFolder) {
                Get-PnPListItem -List $lib -Query $query -PageSize 2000 -FolderServerRelativeUrl $resolvedFolder -Connection $Connection
            } else {
                Get-PnPListItem -List $lib -Query $query -PageSize 2000 -Connection $Connection
            }

            if (-not $akira) {
                Write-Host "    Found 0 '.akira' files."
                continue
            }

            # Filter safely
            $akira = $akira | Where-Object { $_["FileLeafRef"] -and $_["FileLeafRef"].ToLower().EndsWith(".akira") }

            Write-Host "    Found $($akira.Count) '.akira' files."

            foreach ($item in $akira) {

                try {
                    $currentName  = $item["FileLeafRef"]
                    $sourceUrl    = $item["FileRef"]

                    if (-not $sourceUrl) {
                        Write-Warning "      Skipping $currentName : Missing FileRef"
                        continue
                    }

                    $newName      = ($currentName -replace '\.akira$','').TrimEnd()
                    $folderPath   = Split-Path $sourceUrl -Parent
                    $targetUrl    = "$folderPath/$newName" -replace '\\','/'  # ensure forward slashes

                    if ($PSCmdlet.ShouldProcess($sourceUrl,"Rename to $newName")) {

                        Write-Verbose "Moving '$sourceUrl' -> '$targetUrl'"

                        $previousVerbosePreference = $VerbosePreference
                        try {
                            $VerbosePreference = 'SilentlyContinue'
                            Move-PnPFile `
                                -ServerRelativeUrl $sourceUrl `
                                -TargetUrl $targetUrl `
                                -Force `
                                -Connection $Connection `
                                -ErrorAction Stop
                        } finally {
                            $VerbosePreference = $previousVerbosePreference
                        }

                        Write-Host "      SUCCESS: $currentName -> $newName" -ForegroundColor Green
                        Add-Content -Path $LogPath -Value ("{0}\tSUCCESS\t{1}\t{2}\t{3}\t{4}" -f (Get-Date -Format "o"), $siteUrl, $lib, $sourceUrl, $targetUrl)
                    }

                }
                catch {
                    Write-Warning "      FAILED: $currentName : $($_.Exception.Message)"
                    Add-Content -Path $LogPath -Value ("{0}\tFAILED\t{1}\t{2}\t{3}\t{4}" -f (Get-Date -Format "o"), $siteUrl, $lib, $sourceUrl, $_.Exception.Message)
                }

            }
        }
    }
}


<#
    App-only auth (certificate). Use EITHER thumbprint (cert in CurrentUser\My)
    OR certificate path (PFX) + password.
#>
# 1) Connect ONCE to the Admin Center (outside functions)
$adminConn = Connect-PnPTargetSite -Url $AdminUrl -ClientId $ClientId -TenantId $TenantId -CertThumbprint $CertThumbprint -CertPath $CertPath -CertPassword $CertPassword
if (-not $adminConn) {
    throw "Failed to connect to admin center: $AdminUrl"
}

# 2) Build the URL array (Teams + known URLs)
$siteUrls = Get-TeamsSiteUrls -AllSites -IncludeNoAccess:$false

<#
    Examples (choose one block at a time):
    1) Single site + folder scope, dry-run (-WhatIf)
    2) Single site + folder scope, execute
    3) Single site + full library, execute
    4) Multiple sites, dry-run (-WhatIf)
    5) Multiple sites, execute
    6) One-off folder scan (no rename)
#>

# 1) Single site + folder scope, dry-run
foreach ($u in $siteUrl) {
    $conn = Connect-PnPTargetSite -Url $u -ClientId $ClientId -TenantId $TenantId -CertThumbprint $CertThumbprint -CertPath $CertPath -CertPassword $CertPassword
    if (-not $conn) { continue }
    try { $null = Get-PnPWeb -Includes Url -Connection $conn -ErrorAction Stop } catch { Write-Warning "Skipping '$u': $($_.Exception.Message)"; continue }
    Invoke-RemoveAkiraSuffix -SiteUrls $u -Libraries $LibrariesToUse -FolderServerRelativeUrl "/sites/Marketing/Shared Documents/Brand/Assets/Artwork/Sample Folder" -Connection $conn -Verbose -WhatIf
}

# 2) Single site + folder scope, execute
foreach ($u in $siteUrl) {
    $conn = Connect-PnPTargetSite -Url $u -ClientId $ClientId -TenantId $TenantId -CertThumbprint $CertThumbprint -CertPath $CertPath -CertPassword $CertPassword
    if (-not $conn) { continue }
    try { $null = Get-PnPWeb -Includes Url -Connection $conn -ErrorAction Stop } catch { Write-Warning "Skipping '$u': $($_.Exception.Message)"; continue }
    Invoke-RemoveAkiraSuffix -SiteUrls $u -Libraries $LibrariesToUse -FolderServerRelativeUrl "/sites/Marketing/Shared Documents/Brand/Assets/Artwork/Sample Folder" -Connection $conn -Verbose
}

# 3) Single site + full library, execute
foreach ($u in $siteUrl) {
    $conn = Connect-PnPTargetSite -Url $u -ClientId $ClientId -TenantId $TenantId -CertThumbprint $CertThumbprint -CertPath $CertPath -CertPassword $CertPassword
    if (-not $conn) { continue }
    try { $null = Get-PnPWeb -Includes Url -Connection $conn -ErrorAction Stop } catch { Write-Warning "Skipping '$u': $($_.Exception.Message)"; continue }
    Invoke-RemoveAkiraSuffix -SiteUrls $u -Libraries $LibrariesToUse -Connection $conn -Verbose
}

# 4) Multiple sites, dry-run
foreach ($u in $siteUrls) {
    $conn = Connect-PnPTargetSite -Url $u -ClientId $ClientId -TenantId $TenantId -CertThumbprint $CertThumbprint -CertPath $CertPath -CertPassword $CertPassword
    if (-not $conn) { continue }
    try { $null = Get-PnPWeb -Includes Url -Connection $conn -ErrorAction Stop } catch { Write-Warning "Skipping '$u': $($_.Exception.Message)"; continue }
    Invoke-RemoveAkiraSuffix -SiteUrls $u -Libraries $LibrariesToUse -Connection $conn -LogPath $LogPath -WhatIf -Verbose
}

# 5) Multiple sites, execute
foreach ($u in $siteUrls) {
    $conn = Connect-PnPTargetSite -Url $u -ClientId $ClientId -TenantId $TenantId -CertThumbprint $CertThumbprint -CertPath $CertPath -CertPassword $CertPassword
    if (-not $conn) { continue }
    try { $null = Get-PnPWeb -Includes Url -Connection $conn -ErrorAction Stop } catch { Write-Warning "Skipping '$u': $($_.Exception.Message)"; continue }
    Invoke-RemoveAkiraSuffix -SiteUrls $u -Libraries $LibrariesToUse -Connection $conn -LogPath $LogPath -Verbose
}

# 6) One-off folder scan (no rename)
Get-PnPListItem -List "Shared Documents" -FolderServerRelativeUrl "/sites/Marketing/Shared Documents/Brand/Assets/Artwork/Sample Folder" -PageSize 2000 -Connection $conn | Select-Object Id, @{n="Name";e={$_.FieldValues.FileLeafRef}}, @{n="FileRef";e={$_.FieldValues.FileRef}}