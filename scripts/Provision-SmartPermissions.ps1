<#
.SYNOPSIS
    Provisions a "Permissions.aspx" page with the Smart Permissions web part
    on every site collection in the tenant.

.DESCRIPTION
    Connects to the SharePoint admin center, enumerates all site collections
    (excluding OneDrive), and for each site:
      - Creates "Permissions.aspx" if it does not already exist
      - Adds the Smart Permissions web part to the page
      - Publishes the page

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
    To test on a single site before running tenant-wide, comment out the
    foreach loop and replace with a direct Connect-PnPOnline to one site URL.

    The Smart Permissions app must already be deployed to the tenant App Catalog
    and approved before this script will successfully add the web part.
#>

# --- Configuration ---
$adminUrl     = ""   # Required: e.g. "https://contoso-admin.sharepoint.com"
$componentId  = "c2f2de50-4e8b-4f3c-9abc-ef0123456789"  # Smart Permissions web part ID — do not change
$pageName     = "Permissions"   # File name (without .aspx)
$pageTitle    = "Permissions"   # Display title shown at top of page
$webPartProps = @{
    defaultView = "home"        # Starting view for the web part ("home", "report", etc.)
}

# --- Validate config ---
if (-not $adminUrl) {
    throw "adminUrl is required. Set it at the top of the script before running."
}

# --- Connect to admin center and enumerate all site collections ---
Connect-PnPOnline -Url $adminUrl -Interactive
$sites = Get-PnPTenantSite -IncludeOneDriveSites:$false

Write-Host "Found $($sites.Count) site(s). Starting provisioning..." -ForegroundColor Cyan

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
        Write-Host "  Done." -ForegroundColor Green
    }
    catch {
        # Log and continue — don't abort the entire run for one site
        Write-Warning "  Failed on $($site.Url): $_"
    }
}

Write-Host "Provisioning complete." -ForegroundColor Cyan
