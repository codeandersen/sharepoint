<#
        .SYNOPSIS
        Creates Sharepoint Teams Sites and folders.

        .DESCRIPTION
        Script is used to Creates Sharepoint Teams Sites and folders in document library from a csv file.

        CSV file format example
        SiteTitle;Alias;Folders;LCID;TimeZone;SiteUrl
        Finance Department;Finance_Department;Finance Suppliers,Finance Invoices;1030;UTCPLUS0100_BRUSSELS_COPENHAGEN_MADRID_PARIS;https://contoso.sharepoint.com/sites/Finance_Department
        
        SiteTitle; (Title for Sharepoint Site)
        Alias; (Sharepoint site alias can't contain blank spaces)
        Folders; (Folders to create in Sharepoint Site Document library. When more than one separatet with a comma ",")
        LCID; (Sharepoint language for Sharepoint site "https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-csom/jj167546(v=office.15)")
        TimeZone; (Sharepoint Site TimeZone https://pnp.github.io/pnpcore/api/PnP.Core.Admin.Model.SharePoint.TimeZone.html)
        SiteUrl (Sharepoint URL for the site. This is used to check if Sharepoint Site already exists)


        .PARAMETER Name
        -csvFilePath
            Path to CSV file
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false   

        -AdminSiteUrl
            Sharepoint admin site url
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false   

        -PnPOnlineAppId
            PnPOnline Application ID
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false   

        .PARAMETER Extension

        .EXAMPLE
        C:\PS> 2.Create-SharepointSites.ps1 -csvFilePath "C:\script\Sharepoint Teams.csv" -AdminSiteUrl contoso-admin.sharepoint.com -PnPOnlineAppId a4339632-4336-4337-833a-63f0332833e

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        BLOG: http://www.apento.com
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>

param (
    [string]$csvFilePath,
    [string]$AdminSiteUrl,
    [string]$PnPOnlineAppId
)

# Check if CSV file path is provided
if (-not (Test-Path $csvFilePath)) {
    Write-Host "CSV file not found: $csvFilePath" -ForegroundColor Red
    exit
}

# Import the CSV file
$sites = Import-Csv -Path $csvFilePath -Delimiter ";"

# Connect to SharePoint Online using PnP PowerShell with app-only authentication
Connect-PnPOnline -Url $AdminSiteUrl -ClientId $PnPOnlineAppId -Interactive -Verbose

# Loop through each row in the CSV file and create a SharePoint site
foreach ($site in $sites) {
    $siteTitle = $site.SiteTitle
    $siteAlias = $site.Alias
    $folders = $site.Folders -split ","
    $lcid = $site.LCID
    $timezone = $site.Timezone
    $siteUrl = $site.SiteUrl

    # Check if the site already exists
    $existingSite = Get-PnPTenantSite -Identity $siteUrl -ErrorAction SilentlyContinue -Verbose
    if ($existingSite -ne $null) {
        Write-Host "Site already exists: $siteUrl" -ForegroundColor Yellow
        continue
    }

    # Try-Catch block for site creation
    try {
        # Create the SharePoint site
        Write-Host "Creating site: $siteTitle" -ForegroundColor Green
        New-PnPSite -Type TeamSite -Title "$siteTitle" -Alias $siteAlias -Lcid $lcid -TimeZone $timezone -Wait -Verbose
        
        # Reconnect to the newly created site using the same interactive method
        Connect-PnPOnline -Url $siteUrl -ClientId $PnPOnlineAppId -Interactive

        # Create folders in the default document library
        foreach ($folder in $folders) {
            $SiteDocFolder = "$siteUrl/Delte Dokumenter"
            Write-Host "Creating folder: $folder in $SiteDocFolder" -ForegroundColor Cyan
            # Create folder in the Documents library
            Add-PnPFolder -Name "$folder" -Folder "$SiteDocFolder" -Verbose
        }
    }
    catch {
        # If there's an error, display the error message and continue with the next site
        Write-Host "Error creating site: $siteTitle ($siteUrl)" -ForegroundColor Red
        Write-Host "Error message: $_" -ForegroundColor Red
        # Disconnect PnPOnline
        Disconnect-PnPOnline
    }

    
}

# Disconnect from the admin site
Disconnect-PnPOnline
