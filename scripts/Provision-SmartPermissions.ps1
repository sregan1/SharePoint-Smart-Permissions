<#
.SYNOPSIS
    Provisions a "Permissions.aspx" page with the Smart Permissions web part on
    one site, a list of sites, or every site you can find through search --
    without PnP.PowerShell or any other module.

.DESCRIPTION
    Same outcome as Provision-SmartPermissions-PnP.ps1, using only built-in
    PowerShell (5.1 or 7+), OAuth device-code sign-in and the SharePoint REST API.
    For each site it:
      - Creates "Permissions.aspx" if it does not already exist
      - Adds the Smart Permissions web part (skipped if already on the page)
      - Publishes the page
      - Breaks inheritance on the page and grants only the Owners group Full Control
      - Adds the page to Quick Launch
      - Blanks the byline so the page doesn't show the account that ran the script

    Sites that fail are logged as warnings and skipped.

.PREREQUISITES
    No app registration is needed. Sign-in uses Microsoft's own "SharePoint Online
    Management Shell" public client (the same one Connect-SPOService uses), which
    exists in every tenant. If your tenant's Conditional Access blocks the device
    code flow for it, or you prefer your own app, create a public-client app
    registration with SharePoint delegated AllSites.FullControl and set $clientId.
    The signed-in account needs rights on each target site (a SharePoint
    Administrator needs to be a site collection admin there, or use
    Set-SPOUser/Add-SPOSiteCollectionAdmin first).

    The Smart Permissions app must already be deployed to the tenant App Catalog
    and approved.

.NOTES
    Site selection, in priority order:
      1. $targetSiteUrl           one site
      2. $siteListPath            text file, one site URL per line
      3. (neither)                SharePoint search for all site collections the
                                  signed-in account can see. Search is security-trimmed
                                  and can lag behind newly created sites, so for a
                                  complete tenant-wide run prefer $siteListPath
                                  (e.g. exported from the SharePoint admin center).

    Try it on one test site ($targetSiteUrl) first.
#>

# --- Configuration ---
$tenantName     = ""   # Required: e.g. "contoso" (from contoso.sharepoint.com)
$clientId       = "9bc3ab49-b65d-410a-85ad-de819febfddc"  # Microsoft's "SharePoint Online Management Shell" public client,
                       # already present in every tenant - no app registration needed.
                       # Override only if you prefer your own app registration.
$targetSiteUrl  = ""   # Optional: single site, e.g. "https://contoso.sharepoint.com/sites/mysite"
$siteListPath   = ""   # Optional: text file with one site URL per line
$componentId    = "c2f2de50-4e8b-4f3c-9abc-ef0123456789"  # Smart Permissions web part ID - do not change
$rememberSignIn = $true   # Cache the sign-in (encrypted for your Windows user) so later runs don't prompt.
                          # Delete the file named in the "Using saved sign-in" message to sign out.
$pageName       = "Permissions"
$pageTitle      = "Permissions"
$webPartProps   = @{ defaultView = "home" }

if (-not $tenantName -or -not $clientId) {
    throw "tenantName is required. Set it at the top of the script."
}

$ErrorActionPreference = 'Stop'
# Accept "contoso", "contoso-admin", "contoso.sharepoint.com" or a full URL of either.
$tenantName = ($tenantName.Trim() -replace '^https?://', '' -replace '/.*$', '' -replace '\.sharepoint\.com$', '' -replace '-admin$', '')
$spHost   = "https://$tenantName.sharepoint.com"

# Resolve the tenant ID from SharePoint itself: an unauthenticated call returns a 401
# whose WWW-Authenticate header carries realm="<tenant guid>". This avoids guessing
# the tenant's *.onmicrosoft.com domain, which doesn't always match the SharePoint name.
function Get-TenantId {
    try {
        Invoke-WebRequest -Uri "$spHost/_vti_bin/client.svc" -Headers @{ Authorization = 'Bearer' } -UseBasicParsing | Out-Null
    } catch {
        $resp = $_.Exception.Response
        $hdr = $null
        if ($resp) {
            try { $hdr = $resp.Headers.GetValues('WWW-Authenticate') -join ' ' } catch { $hdr = [string]$resp.Headers['WWW-Authenticate'] }
        }
        if ($hdr -match 'realm="([0-9a-fA-F-]{36})"') { return $Matches[1] }
    }
    throw "Could not determine the tenant ID from $spHost. Check `$tenantName (the part before .sharepoint.com)."
}
$tenantId = Get-TenantId

# --- Authentication (device code flow) ---
$script:token = $null
$script:refreshToken = $null
$script:tokenExpires = [datetime]::MinValue

function Connect-SharePoint {
    $scope = "$spHost/.default offline_access"
    $dc = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode" `
        -Body @{ client_id = $clientId; scope = $scope }
    Write-Host $dc.message -ForegroundColor Yellow
    $deadline = (Get-Date).AddSeconds($dc.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $dc.interval
        try {
            $t = Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
                -Body @{
                    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                    client_id   = $clientId
                    device_code = $dc.device_code
                }
            Set-Token $t
            return
        } catch {
            $body = $null
            if ($_.ErrorDetails.Message) { $body = $_.ErrorDetails.Message | ConvertFrom-Json }
            if ($body.error -ne 'authorization_pending') { throw "Sign-in failed: $($_.ErrorDetails.Message)" }
        }
    }
    throw "Sign-in timed out."
}

# The refresh token is stored encrypted with Windows DPAPI (ConvertFrom-SecureString), so only
# the same Windows user on the same machine can read it. Where DPAPI isn't available the save
# fails quietly and you simply sign in each run.
$tokenCachePath = Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { $HOME }) "SmartPermissions-signin-$tenantId.txt"

function Save-SignIn {
    if (-not $rememberSignIn -or -not $script:refreshToken) { return }
    try {
        ConvertTo-SecureString $script:refreshToken -AsPlainText -Force | ConvertFrom-SecureString |
            Set-Content -Path $tokenCachePath -Encoding ASCII
    } catch { }
}

function Restore-SignIn {
    if (-not $rememberSignIn -or -not (Test-Path $tokenCachePath)) { return $false }
    try {
        $secure = Get-Content -Path $tokenCachePath | ConvertTo-SecureString
        $script:refreshToken = [System.Net.NetworkCredential]::new('', $secure).Password
        $script:tokenExpires = [datetime]::MinValue
        Get-AccessToken | Out-Null   # proves the saved token still works
        Write-Host "Using saved sign-in ($tokenCachePath)" -ForegroundColor DarkGray
        return $true
    } catch {
        Remove-Item $tokenCachePath -ErrorAction SilentlyContinue
        $script:refreshToken = $null
        return $false
    }
}

function Set-Token($t) {
    $script:token = $t.access_token
    if ($t.refresh_token) { $script:refreshToken = $t.refresh_token; Save-SignIn }
    $script:tokenExpires = (Get-Date).AddSeconds([int]$t.expires_in - 120)
}

function Get-AccessToken {
    if ((Get-Date) -ge $script:tokenExpires) {
        $t = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
            -Body @{
                grant_type    = "refresh_token"
                client_id     = $clientId
                refresh_token = $script:refreshToken
                scope         = "$spHost/.default offline_access"
            }
        Set-Token $t
    }
    return $script:token
}

# --- REST helper (retries on throttling) ---
function Invoke-Sp {
    param(
        [string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Url,
        $Body = $null,
        [hashtable]$Headers = @{},
        [string]$ContentType = 'application/json;odata=nometadata'
    )
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $h = @{
            Authorization = "Bearer $(Get-AccessToken)"
            Accept        = 'application/json;odata=nometadata'
        }
        foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] }
        $req = @{ Method = $Method; Uri = $Url; Headers = $h }
        if ($null -ne $Body) {
            $req.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 -Compress }
            $req.ContentType = $ContentType
        } elseif ($Method -ne 'GET') {
            $req.Body = ''
            $req.ContentType = $ContentType
        }
        try {
            return Invoke-RestMethod @req
        } catch {
            $code = [int]$_.Exception.Response.StatusCode
            if (($code -eq 429 -or $code -eq 503) -and $attempt -lt 5) {
                $wait = 5 * $attempt
                try { $ra = $_.Exception.Response.Headers['Retry-After']; if ($ra) { $wait = [int]$ra } } catch { }
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Test-SpExists([string]$Url) {
    try { Invoke-Sp -Url $Url | Out-Null; return $true }
    catch {
        if ([int]$_.Exception.Response.StatusCode -eq 404) { return $false }
        throw
    }
}

# --- Site list ---
function Get-SitesFromSearch {
    $urls = New-Object System.Collections.Generic.List[string]
    $start = 0
    do {
        $q = [uri]::EscapeDataString("contentclass:STS_Site -path:`"$tenantName-my.sharepoint.com`"")
        $r = Invoke-Sp -Url "$spHost/_api/search/query?querytext='$q'&selectproperties='Path'&trimduplicates=false&rowlimit=500&startrow=$start"
        $rows = @($r.PrimaryQueryResult.RelevantResults.Table.Rows)
        foreach ($row in $rows) {
            $p = ($row.Cells | Where-Object { $_.Key -eq 'Path' }).Value
            if ($p) { $urls.Add($p.TrimEnd('/')) }
        }
        $start += 500
    } while ($rows.Count -eq 500)
    return $urls | Select-Object -Unique
}

Write-Host "IMPORTANT: the account you sign in with must be a site collection administrator on every site this script processes." -ForegroundColor Yellow
Write-Host "Sites where it isn't will fail with an access error and be skipped." -ForegroundColor Yellow

if (-not (Restore-SignIn)) { Connect-SharePoint }

if ($targetSiteUrl) {
    Write-Host "Single-site mode: $targetSiteUrl" -ForegroundColor Cyan
    $siteUrls = @($targetSiteUrl.TrimEnd('/'))
} elseif ($siteListPath) {
    $siteUrls = @(Get-Content $siteListPath | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim().TrimEnd('/') })
    Write-Host "Loaded $($siteUrls.Count) site(s) from $siteListPath" -ForegroundColor Cyan
} else {
    $siteUrls = @(Get-SitesFromSearch)
    Write-Host "Search found $($siteUrls.Count) site(s)." -ForegroundColor Cyan
}

foreach ($siteUrl in $siteUrls) {
    Write-Host "Processing: $siteUrl"
    try {
        $api = "$siteUrl/_api"
        $web = Invoke-Sp -Url "$api/web?`$select=ServerRelativeUrl,Url"
        $webPath = $web.ServerRelativeUrl.TrimEnd('/')
        $pageLeaf = "$pageName.aspx"
        $pageFileUrl = "$webPath/SitePages/$pageLeaf"
        $fileApi = "$api/web/GetFileByServerRelativeUrl('$pageFileUrl')"

        # ── Create the page if it doesn't exist ─────────────────────────────
        if (Test-SpExists "$fileApi`?`$select=Name") {
            Write-Host "  $pageLeaf already exists"
        } else {
            $created = Invoke-Sp -Method Post -Url "$api/sitepages/pages" -Body @{
                Title = $pageTitle
                PageLayoutType = 'Article'
                PromotedState = 0
            }
            # Use the name SharePoint actually gave the page (it can differ from $pageName).
            if ($created.Name) {
                $pageLeaf = $created.Name
                $pageFileUrl = "$webPath/SitePages/$pageLeaf"
                $fileApi = "$api/web/GetFileByServerRelativeUrl('$pageFileUrl')"
            }
            Write-Host "  Created $pageLeaf"
        }

        # A just-created page can take a moment to become readable.
        $item = $null
        for ($try = 1; $try -le 6 -and -not $item; $try++) {
            try { $item = Invoke-Sp -Url "$fileApi/ListItemAllFields?`$select=Id" }
            catch {
                if ($try -eq 6) { throw }
                Start-Sleep -Seconds 3
            }
        }
        $pageApi = "$api/sitepages/pages($($item.Id))"

        # ── Add the web part unless it's already on the page ────────────────
        $current = Invoke-Sp -Url "$pageApi`?`$select=CanvasContent1"
        # CanvasContent1 is HTML, not JSON: each control is a div carrying HTML-encoded JSON
        # in data attributes. (An earlier run of this script wrote a raw JSON array, which
        # SharePoint stores but never renders, so only treat HTML canvases as "already there".)
        $hasWebPart = $current.CanvasContent1 -and
            $current.CanvasContent1 -match 'data-sp-canvascontrol' -and
            $current.CanvasContent1 -match [regex]::Escape($componentId)
        if ($hasWebPart) {
            Write-Host "  Web part already present - skipping add"
            $canvas = $current.CanvasContent1
        } else {
            function ConvertTo-Attr($obj) {
                [System.Net.WebUtility]::HtmlEncode(($obj | ConvertTo-Json -Depth 20 -Compress))
            }
            $controlData = ConvertTo-Attr @{
                controlType = 3
                id          = [guid]::NewGuid().ToString()
                webPartId   = $componentId
                position    = @{ zoneIndex = 1; sectionIndex = 1; controlIndex = 1; sectionFactor = 12; layoutIndex = 1 }
                emphasis    = @{}
            }
            $webPartData = ConvertTo-Attr @{
                id          = $componentId
                instanceId  = [guid]::NewGuid().ToString()
                title       = 'Smart Permissions'
                description = 'Smart Permissions'
                dataVersion = '1.0'
                properties  = $webPartProps
                serverProcessedContent = @{
                    htmlStrings = @{}; searchablePlainTexts = @{}; imageSources = @{}; links = @{}
                }
            }
            $settingsData = ConvertTo-Attr @{
                controlType = 0
                pageSettingsSlice = @{ isDefaultDescription = $true; isDefaultThumbnail = $true }
            }
            $canvas = '<div>' +
                "<div data-sp-canvascontrol=`"`" data-sp-canvasdataversion=`"1.0`" data-sp-controldata=`"$controlData`">" +
                "<div data-sp-webpart=`"`" data-sp-webpartdataversion=`"1.0`" data-sp-webpartdata=`"$webPartData`">" +
                "<div data-sp-componentid=`"`">$componentId</div><div data-sp-htmlproperties=`"`"></div></div></div>" +
                "<div data-sp-canvascontrol=`"`" data-sp-canvasdataversion=`"1.0`" data-sp-controldata=`"$settingsData`"></div>" +
                '</div>'
            Write-Host "  Web part added"
        }

        # Title region with no author byline (so the page doesn't show who ran this script)
        $layout = @(@{
            id = 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788'
            instanceId = 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788'
            title = 'Title Region'
            description = 'Title Region Description'
            serverProcessedContent = @{ htmlStrings = @{}; searchablePlainTexts = @{ title = $pageTitle }; imageSources = @{}; links = @{} }
            dataVersion = '1.4'
            properties = @{
                title = $pageTitle; imageSourceType = 4; layoutType = 'NoImage'; textAlignment = 'Left'
                showTopicHeader = $false; showPublishDate = $false; topicHeader = ''
                authors = @(); authorByline = @()
            }
        }) | ConvertTo-Json -Depth 20 -Compress

        # Write the page content straight to the page's list item. (ValidateUpdateListItem
        # caps these fields at 255 characters and the sitepages savepage route isn't
        # available everywhere, but a plain item update has neither limit.)
        $listItemApi = "$fileApi/ListItemAllFields"
        $entityType = (Invoke-Sp -Url "$listItemApi/ParentList?`$select=ListItemEntityTypeFullName").ListItemEntityTypeFullName
        Invoke-Sp -Method Post -Url $listItemApi `
            -ContentType 'application/json;odata=verbose' `
            -Headers @{ 'X-HTTP-Method' = 'MERGE'; 'IF-MATCH' = '*' } `
            -Body (@{
                __metadata            = @{ type = $entityType }
                Title                 = $pageTitle
                CanvasContent1        = $canvas
                LayoutWebpartsContent = $layout
            } | ConvertTo-Json -Depth 5 -Compress) | Out-Null

        # Publish first so navigation can link to it (a never-published draft can't be linked).
        try { Invoke-Sp -Method Post -Url "$fileApi/CheckIn(comment='Automated provisioning',checkintype=1)" | Out-Null } catch { }
        Invoke-Sp -Method Post -Url "$fileApi/Publish(comment='Automated provisioning')" | Out-Null

        # ── Permissions: Owners only ────────────────────────────────────────
        $owners = Invoke-Sp -Url "$api/web/AssociatedOwnerGroup?`$select=Id"
        $fullControl = Invoke-Sp -Url "$api/web/roledefinitions/getbytype(5)?`$select=Id"
        Invoke-Sp -Method Post -Url "$listItemApi/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)" | Out-Null
        Invoke-Sp -Method Post -Url "$listItemApi/roleassignments/addroleassignment(principalid=$($owners.Id),roledefid=$($fullControl.Id))" | Out-Null
        Write-Host "  Permissions set: Owners only"

        # ── Navigation: Quick Launch ────────────────────────────────────────
        $pageUrl = "$siteUrl/SitePages/$pageLeaf"
        $nodes = (Invoke-Sp -Url "$api/web/navigation/quicklaunch?`$select=Id,Title,Url").value
        # SharePoint can store the link as server-relative or absolute, so match on the page's
        # path ending rather than the full URL; otherwise every run adds another copy.
        $pageUrlPattern = "(^|/)SitePages/$([regex]::Escape($pageLeaf))$"
        foreach ($n in @($nodes | Where-Object { $_.Title -eq $pageTitle -and $_.Url -match $pageUrlPattern })) {
            Invoke-Sp -Method Post -Url "$api/web/navigation/getnodebyid($($n.Id))" `
                -Headers @{ 'X-HTTP-Method' = 'DELETE'; 'IF-MATCH' = '*' } | Out-Null
        }

        Invoke-Sp -Method Post -Url "$api/web/navigation/quicklaunch" -Body @{ Title = $pageTitle; Url = $pageUrl; IsExternal = $false } | Out-Null
        Write-Host "  Navigation node added"

        # The permission updates can leave the page as a draft; publish for good.
        try { Invoke-Sp -Method Post -Url "$fileApi/CheckIn(comment='Automated provisioning',checkintype=1)" | Out-Null } catch { }
        Invoke-Sp -Method Post -Url "$fileApi/Publish(comment='Automated provisioning')" | Out-Null
        Write-Host "  Page published"

        Write-Host "  Done." -ForegroundColor Green
    }
    catch {
        Write-Warning "  Failed on ${siteUrl}: $_ $($_.ErrorDetails.Message)"
        Write-Warning "  At line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
    }
}

Write-Host "Provisioning complete." -ForegroundColor Cyan
