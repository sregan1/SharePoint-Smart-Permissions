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

    The Smart Permissions app must already be deployed to the tenant App Catalog
    and approved before this script will successfully add the web part.

    Audience targeting (the -AudienceIds parameter on Add-PnPNavigationNode
    and the page audience field) is only available on Microsoft 365
    Group-connected sites. Classic team sites will get the nav node added
    without audience filtering and will skip the page audience step.
#>

# --- Configuration ---
$adminUrl       = ""   # Required: e.g. "https://contoso-admin.sharepoint.com"
$targetSiteUrl  = ""   # Optional: set to provision a single site, e.g. "https://contoso.sharepoint.com/sites/mysite"
                       # Leave empty to provision all site collections in the tenant.
$componentId    = "c2f2de50-4e8b-4f3c-9abc-ef0123456789"  # Smart Permissions web part ID — do not change
$pageName       = "Permissions"   # File name (without .aspx)
$pageTitle      = "Permissions"   # Display title shown at top of page
$webPartProps   = @{
    defaultView = "home"          # Starting view for the web part ("home", "report", etc.)
}

# --- Validate config ---
if (-not $adminUrl) {
    throw "adminUrl is required. Set it at the top of the script before running."
}

# --- Build the list of sites to provision ---
if ($targetSiteUrl) {
    # Single-site mode — connect directly, no admin center enumeration needed
    Write-Host "Single-site mode: $targetSiteUrl" -ForegroundColor Cyan
    $sites = @([PSCustomObject]@{ Url = $targetSiteUrl.TrimEnd('/') })
} else {
    # Tenant-wide mode
    Connect-PnPOnline -Url $adminUrl -Interactive
    $sites = Get-PnPTenantSite -IncludeOneDriveSites:$false
    Write-Host "Found $($sites.Count) site(s). Starting provisioning..." -ForegroundColor Cyan
}

foreach ($site in $sites) {
    Write-Host "Processing: $($site.Url)"
    try {
        Connect-PnPOnline -Url $site.Url -Interactive

        # Create the page only if it doesn't already exist
        $page = Get-PnPPage -Identity "$pageName.aspx" -ErrorAction SilentlyContinue
        if (-not $page) {
            $page = Add-PnPPage -Name $pageName -Title $pageTitle `
                                -LayoutType Article
            Write-Host "  Created $pageName.aspx"
        } else {
            Write-Host "  $pageName.aspx already exists"
        }

        $web = Get-PnPWeb
        $pageFileUrl = "$($web.ServerRelativeUrl.TrimEnd('/'))/SitePages/$pageName.aspx"

        # Resolve the pages library once by its stable, non-localized root-folder
        # URL segment ("SitePages") rather than the display title "Site Pages" —
        # the title is localized (e.g. German "Websiteseiten"), which would make
        # every -List "Site Pages" lookup below fail silently on non-English sites,
        # leaving the page provisioned but still inheriting Members/Visitors access.
        $pagesList = Get-PnPList -Identity "SitePages"

        # Skip if the web part is already on the page - otherwise re-running
        # against an existing page adds a second copy of it every time.
        $existingWebPart = Get-PnPPageComponent -Page $page |
            Where-Object { $_.Id -and [Guid]$_.Id -eq [Guid]$componentId }
        if ($existingWebPart) {
            Write-Host "  Web part already present - skipping add"
        } else {
            # Resolve to a full component object rather than passing the bare GUID -
            # passing just the string leaves the control's component reference
            # unbound (null id / empty data-sp-componentid in the saved canvas),
            # which breaks manifest resolution when the page renders.
            $component = Get-PnPPageComponent -Page $page -ListAvailable |
                Where-Object { [Guid]$_.Id -eq [Guid]$componentId }
            if (-not $component) {
                throw "Component $componentId is not available on this site (not installed/approved?)"
            }

            Add-PnPPageWebPart -Page $page `
                -Component $component `
                -WebPartProperties $webPartProps
            Write-Host "  Web part added"
        }

        # Publish now, establishing the first major/published version. Nav
        # nodes can't link to a page that has never been published at all
        # ("Cannot add the file ... because it is a draft item") — a minor
        # draft bump on top of an already-published page is fine, but a page
        # with zero major versions isn't. The permissions/audience updates
        # below will bump it back to an unpublished draft; the REST call
        # further down re-publishes it for good afterward.
        Set-PnPPage -Identity $page -Publish -CommentsEnabled:$false -HeaderLayoutType NoImage

        # ── Permissions: Owners only ──────────────────────────────────────────
        $pageItem = Get-PnPListItem -List $pagesList `
            -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$pageName.aspx</Value></Eq></Where></Query></View>"
        if ($pageItem) {
            $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup
            # Breaks inheritance, clears existing role assignments, and grants
            # the Owners group Full Control — all in one call, since
            # -ClearExisting is only valid alongside -User/-Group.
            Set-PnPListItemPermission -List $pagesList -Identity $pageItem.Id `
                -Group $ownersGroup -AddRole "Full Control" -ClearExisting
            Write-Host "  Permissions set: Owners only"
        } else {
            Write-Warning "  Could not find page list item — permissions not updated"
        }

        # ── Navigation: add to Quick Launch ──────────────────────────────────
        $pageUrl = "$($site.Url)/SitePages/$pageName.aspx"
        # Remove any existing node for this exact page (title AND url) so an
        # unrelated node that happens to share the page's title is left alone.
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
            Set-PnPList -Identity $pagesList -EnableModernAudienceTargeting $true
            # Set the M365 Owners group as the page audience
            $ownersClaim = "c:0o.c|federateddirectoryclaimprovider|$($siteData.GroupId)"
            Set-PnPListItem -List $pagesList -Identity $pageItem.Id -Values @{
                "_ModernAudienceTargetUserField" = $ownersClaim
            } | Out-Null
            Write-Host "  Page audience set: Owners group"
        } elseif ($pageItem) {
            Write-Host "  Page audience skipped — non-M365 site (SharePoint groups not supported for audience targeting)"
        }

        # Clear the byline so the page doesn't show whichever account ran
        # this provisioning script as the author. Article's page header
        # always renders "By <name>" from PageHeader.AuthorByLine unless
        # it's blanked out here; done last so the permissions/audience
        # list-item updates above don't get overwritten by this Save().
        $page = Get-PnPPage -Identity "$pageName.aspx"
        $page.PageHeader.AuthorByLine = ""
        $page.PageHeader.Authors = ""
        $page.Save() | Out-Null

        # Publish directly via SharePoint's native REST endpoint (the same one
        # the browser's "Publish" button calls) rather than relying on
        # PnP.Core's Page.Publish(), which has repeatedly left the page in an
        # unpublished "Needs Publishing" state after the permissions/audience
        # updates above touch the list item.
        $publishUrl = "$($web.Url)/_api/web/getfilebyserverrelativeurl('$pageFileUrl')/Publish(comment='Automated provisioning')"
        Invoke-PnPSPRestMethod -Url $publishUrl -Method Post -Content "{}" -ContentType "application/json" | Out-Null
        Write-Host "  Page published"

        Write-Host "  Done." -ForegroundColor Green
    }
    catch {
        # Log and continue — don't abort the entire run for one site
        Write-Warning "  Failed on $($site.Url): $_"
        Write-Warning "  At line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
    }
}

Write-Host "Provisioning complete." -ForegroundColor Cyan
