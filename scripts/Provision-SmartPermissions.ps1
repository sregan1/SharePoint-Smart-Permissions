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
                                -LayoutType Article -SkipPublish
            Write-Host "  Created $pageName.aspx"
        } else {
            Write-Host "  $pageName.aspx already exists — adding web part"
        }

        Add-PnPPageWebPart -Page $page `
            -DefaultWebPartType ThirdParty `
            -WebPartId $componentId `
            -WebPartProperties $webPartProps

        Publish-PnPPage -Identity "$pageName.aspx"

        # ── Permissions: Owners only ──────────────────────────────────────────
        $pageItem = Get-PnPListItem -List "Site Pages" `
            -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$pageName.aspx</Value></Eq></Where></Query></View>"
        if ($pageItem) {
            $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup
            # Break inheritance and remove all existing role assignments
            Set-PnPListItemPermission -List "Site Pages" -Identity $pageItem.Id `
                -InheritPermissions:$false -ClearExisting
            # Grant the Owners group Full Control
            Set-PnPListItemPermission -List "Site Pages" -Identity $pageItem.Id `
                -Group $ownersGroup -AddRole "Full Control"
            Write-Host "  Permissions set: Owners only"
        } else {
            Write-Warning "  Could not find page list item — permissions not updated"
        }

        # ── Navigation: add to Quick Launch ──────────────────────────────────
        $pageUrl = "$($site.Url)/SitePages/$pageName.aspx"
        # Remove any existing node with the same title to avoid duplicates
        $existingNode = Get-PnPNavigationNode -Location QuickLaunch |
            Where-Object { $_.Title -eq $pageTitle }
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
            Set-PnPList -Identity "Site Pages" -EnableAudienceTargeting $true
            # Set the M365 Owners group as the page audience
            $ownersClaim = "c:0o.c|federateddirectoryclaimprovider|$($siteData.GroupId)"
            Set-PnPListItem -List "Site Pages" -Identity $pageItem.Id -Values @{
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
