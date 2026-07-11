<#
.SYNOPSIS
    Provisions a "Permissions.aspx" page with the Smart Permissions web part
    on one site collection or every site collection in the tenant.

.DESCRIPTION
    Connects to the SharePoint admin center, enumerates all site collections
    (excluding OneDrive), and for each site:
      - Creates "Permissions.aspx" if it does not already exist
      - Adds the Smart Permissions web part to the page
      - Publishes the page
      - Breaks permission inheritance on the page and grants only the site
        Owners group access (removing Members and Visitors)
      - Adds the page to the Quick Launch navigation
      - Enables audience targeting on the Site Pages library and sets the
        Owners group as the audience on the page (M365-connected sites only)

    Sites that fail (e.g. locked, no-script, insufficient access) are logged
    as warnings and skipped so the rest of the loop completes.

.PARAMETER adminUrl
    Set the $adminUrl variable at the top of the script before running.
    Format: https://<tenant>-admin.sharepoint.com

.PREREQUISITES
    PnP.PowerShell module must be installed:
        Install-Module PnP.PowerShell

    The account used must have SharePoint Administrator or Global Administrator
    permissions in the tenant.

.EXAMPLE
    # 1. Set $adminUrl at the top of the script
    # 2. Run the script:
    .\Provision-SmartPermissions.ps1

.NOTES
    Set $targetSiteUrl to provision a single site without touching the rest of
    the tenant. Leave it empty to run across all site collections.

    The Smart Permissions app must already be deployed (and approved, if API
    permissions require it) to the tenant App Catalog. The script installs the
    app on each site automatically if it isn't already there, so ticking
    "Make this solution available to all sites automatically" during upload
    is not required.

    Audience targeting (the -AudienceIds parameter on Add-PnPNavigationNode
    and the page audience field) is only available on Microsoft 365
    Group-connected sites. Classic team sites will get the nav node added
    without audience filtering and will skip the page audience step.
#>

# --- Configuration ---
$adminUrl       = ""   # Required in tenant-wide mode: e.g. "https://contoso-admin.sharepoint.com"
                       # Not used in single-site mode (when $targetSiteUrl is set).
$targetSiteUrl  = ""   # Optional: set to provision a single site, e.g. "https://contoso.sharepoint.com/sites/mysite"
                       # Leave empty to provision all site collections in the tenant.
$clientId       = ""   # Optional: Entra ID App client id from Register-PnPEntraIDAppForInteractiveLogin.
                       # Only needed on tenants that reject the default PnP Management Shell app —
                       # a growing number of tenants require this for any interactive PnP sign-in.
$solutionId     = "a2f2de50-4e8b-4f3c-9abc-ef0123456789"  # Smart Permissions solution/app ID (package-solution.json) — do not change
$componentId    = "c2f2de50-4e8b-4f3c-9abc-ef0123456789"  # Smart Permissions web part ID — do not change
$pageName       = "Permissions"   # File name (without .aspx)
$pageTitle      = "Permissions"   # Display title shown at top of page
$webPartProps   = @{
    defaultView = "home"          # Starting view for the web part ("home", "report", etc.)
}

# --- Validate config ---
# Only required in tenant-wide mode — single-site mode connects directly to
# $targetSiteUrl and never touches $adminUrl.
if (-not $targetSiteUrl -and -not $adminUrl) {
    throw "adminUrl is required for tenant-wide mode. Set it at the top of the script, or set targetSiteUrl to provision a single site instead."
}

# Splat so -ClientId is only passed when set, without an empty-string arg
$connectParams = @{ Interactive = $true }
if ($clientId) {
    $connectParams["ClientId"] = $clientId
}

# --- Build the list of sites to provision ---
if ($targetSiteUrl) {
    # Single-site mode — connect directly, no admin center enumeration needed
    Write-Host "Single-site mode: $targetSiteUrl" -ForegroundColor Cyan
    $sites = @([PSCustomObject]@{ Url = $targetSiteUrl.TrimEnd('/') })
} else {
    # Tenant-wide mode
    Connect-PnPOnline -Url $adminUrl @connectParams
    $sites = Get-PnPTenantSite -IncludeOneDriveSites:$false
    Write-Host "Found $($sites.Count) site(s). Starting provisioning..." -ForegroundColor Cyan
}

foreach ($site in $sites) {
    Write-Host "Processing: $($site.Url)"
    try {
        Connect-PnPOnline -Url $site.Url @connectParams

        # ── Ensure the Smart Permissions app is installed on this site ───────
        # Scope defaults to Tenant, which is where the app is deployed. Passing
        # -Scope Site here would look for a site-collection app catalog instead,
        # which most sites don't have, causing a sitecollectionappcatalog 404.
        $app = Get-PnPApp -Identity $solutionId -ErrorAction SilentlyContinue
        if (-not $app) {
            Write-Warning "  Smart Permissions app not found in this site's available apps — attempting install from tenant App Catalog..."
        }
        if (-not $app -or -not $app.InstalledVersion) {
            Write-Host "  Installing Smart Permissions app on this site..."
            # -Wait blocks until the (asynchronous) install actually finishes, instead of
            # guessing with a fixed sleep — a too-short sleep left the component unresolved
            # and Add-PnPPageWebPart silently added a control the page couldn't render.
            Install-PnPApp -Identity $solutionId -Wait
        }

        # Create the page only if it doesn't already exist
        $page = Get-PnPPage -Identity "$pageName.aspx" -ErrorAction SilentlyContinue
        if (-not $page) {
            $page = Add-PnPPage -Name $pageName -Title $pageTitle `
                                -LayoutType Article -Publish
            Write-Host "  Created $pageName.aspx"
        } else {
            Write-Host "  $pageName.aspx already exists — adding web part"
        }

        # Resolve the actual component object instead of passing a raw GUID string —
        # passing $componentId directly to -Component silently produced an empty
        # placeholder control (no WebPartId/JsonWebPartData bound) even though the
        # component is genuinely deployed and available on the site.
        $targetComponent = Get-PnPAvailablePageComponents -Page $page | Where-Object { [Guid]$_.Id -eq [Guid]$componentId }

        # Skip if the web part is already on the page (re-running the script on an
        # existing page would otherwise stack up a new duplicate copy each time).
        # Cast both sides to [Guid] — a plain string compare silently fails (and the
        # guard never triggers) if WebPartId ever surfaces braced or upper-case.
        $existingWebPart = $page.Controls | Where-Object { $_.WebPartId -and ([Guid]$_.WebPartId -eq [Guid]$componentId) }
        if (-not $existingWebPart) {
            if ($targetComponent) {
                Add-PnPPageWebPart -Page $page `
                    -Component $targetComponent `
                    -WebPartProperties $webPartProps
            } else {
                Write-Warning "  Could not resolve the Smart Permissions component on this page — skipping add"
            }
        } else {
            Write-Host "  Web part already on page — skipping duplicate add"
        }

        # Publish the SAME in-memory $page object that just had the web part added to
        # it — re-resolving by filename here fetched a stale copy without the new
        # control and published that over the real change.
        Set-PnPPage -Identity $page -Publish -CommentsEnabled:$false

        # Resolve the pages library once by its stable, non-localized root-folder
        # URL segment ("SitePages") rather than the display title "Site Pages" —
        # the title is localized (e.g. German "Websiteseiten"), which would make
        # every -List "Site Pages" lookup below fail silently on non-English sites,
        # leaving the page provisioned but still inheriting Members/Visitors access.
        $pagesList = Get-PnPList -Identity "SitePages"

        # ── Permissions: Owners only ──────────────────────────────────────────
        $pageItem = Get-PnPListItem -List $pagesList `
            -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$pageName.aspx</Value></Eq></Where></Query></View>"
        if ($pageItem) {
            $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup
            # -ClearExisting is only valid alongside -User/-Group (not -InheritPermissions),
            # so break inheritance, clear existing role assignments, and grant Owners in one call.
            Set-PnPListItemPermission -List $pagesList -Identity $pageItem.Id `
                -Group $ownersGroup -AddRole "Full Control" -ClearExisting
            Write-Host "  Permissions set: Owners only"
        } else {
            Write-Warning "  Could not find page list item — permissions not updated"
        }

        # ── Navigation: add to Quick Launch ──────────────────────────────────
        $pageUrl = "$($site.Url)/SitePages/$pageName.aspx"
        # Remove any existing node for this exact page URL (not just matching Title)
        # so an unrelated node that happens to share the page's title is left alone.
        $existingNode = Get-PnPNavigationNode -Location QuickLaunch |
            Where-Object { $_.Title -eq $pageTitle -and $_.Url -eq $pageUrl }
        if ($existingNode) {
            Remove-PnPNavigationNode -Identity $existingNode.Id -Force
        }

        # On M365-connected sites, scope the nav node to the Owners group so
        # only Owners see the link in the left navigation.
        $siteData = Get-PnPSite -Includes GroupId
        if ($siteData.GroupId -ne [Guid]::Empty) {
            # M365 group GUID — owners of the site are owners of this group
            $ownersGroupId = $siteData.GroupId
            Add-PnPNavigationNode -Location QuickLaunch -Title $pageTitle `
                -Url $pageUrl -AudienceIds @($ownersGroupId)
            Write-Host "  Navigation node added (audience: Owners group)"
        } else {
            Add-PnPNavigationNode -Location QuickLaunch -Title $pageTitle -Url $pageUrl
            Write-Host "  Navigation node added (no audience — non-M365 site)"
        }

        # ── Page audience targeting ───────────────────────────────────────────
        if ($pageItem -and $siteData.GroupId -ne [Guid]::Empty) {
            # Enable audience targeting on the Site Pages library (idempotent)
            Set-PnPList -Identity $pagesList -EnableAudienceTargeting $true
            # Set the M365 Owners group as the page audience
            $ownersClaim = "c:0o.c|federateddirectoryclaimprovider|$($siteData.GroupId)"
            Set-PnPListItem -List $pagesList -Identity $pageItem.Id -Values @{
                "_ModernAudienceTargetUserField" = $ownersClaim
            } | Out-Null
            Write-Host "  Page audience set: Owners group"
        } elseif ($pageItem) {
            Write-Host "  Page audience skipped — non-M365 site (SharePoint groups not supported for audience targeting)"
        }

        Write-Host "  Done." -ForegroundColor Green
    }
    catch {
        # Log and continue — don't abort the entire run for one site
        Write-Warning "  Failed on $($site.Url): $_"
    }
}

Write-Host "Provisioning complete." -ForegroundColor Cyan
