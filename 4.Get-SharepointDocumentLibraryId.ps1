<#
        .SYNOPSIS
        Get Library ID's for folders in Sharepoint Site document library

        .DESCRIPTION
        Script is used to Get Library ID's for folders in Sharepoint Site document library from a csv file and export to CSV so it can be used for mapping in Intune.

        CSV file format example
        SiteUrl;LibraryName
        https://contoso.sharepoint.com/sites/Finance_Department;Dokumenter
        
        SiteUrl; (EntraID Security Group to be set as contributor)
        LibraryName; (Sharepoint Library name in danish dokumenter translates into Documents)

        .PARAMETER Name
        -csvFilePath
            Path to CSV file
            
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

        -PnPOnlineAppId
            PnPOnline Application ID
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false  

        .PARAMETER Extension

        .EXAMPLE
        C:\PS> 4.Get-SharepointDocumentLibraryId.ps1 -csvFilePath "C:\script\Sharepoint permissions.csv" -PnPOnlineAppId a4339632-4336-4337-833a-63f0332833e -outputCsvPath "C:\script\4.SharepointDocumentLibraryIDs.csv"

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
    [string]$PnPOnlineAppId,
    [string]$outputCsvPath  # Path where the output CSV will be saved
)

# Import PnP PowerShell module
Import-Module PnP.PowerShell

# Read the CSV file
$sitesData = Import-Csv -Path $csvFilePath -Delimiter ";"

# Function to URL encode and format the values
function CustomUrlEncode {
    param ($value)
    
    # Perform standard URL encoding
    $encodedValue = [System.Web.HttpUtility]::UrlEncode($value)
    
    # Replace specific characters to match the desired format
    $encodedValue = $encodedValue -replace '%7b', '%7B'  
    $encodedValue = $encodedValue -replace '%7d', '%7D'
    $encodedValue = $encodedValue -replace '%2d', '%2D'  
    $encodedValue = $encodedValue -replace '%3a', '%3A'  
    $encodedValue = $encodedValue -replace '%2f', '%2F'  
    $encodedValue = $encodedValue -replace '-', '%2D'  
    
    return $encodedValue
}

# Function to encode web URLs with dot (.) replacement
function CustomUrlEncodeForWebUrl {
    param ($value)
    
    # Perform standard URL encoding
    $encodedValue = [System.Web.HttpUtility]::UrlEncode($value)
    
    # Replace specific characters to match the desired format
    $encodedValue = $encodedValue -replace '%7b', '%7B'  
    $encodedValue = $encodedValue -replace '%7d', '%7D'
    $encodedValue = $encodedValue -replace '%2d', '%2D' 
    $encodedValue = $encodedValue -replace '%3a', '%3A'  
    $encodedValue = $encodedValue -replace '%2f', '%2F'  
    $encodedValue = $encodedValue -replace '\.', '%2E'   
    $encodedValue = $encodedValue -replace '_', '%5F'   
    
    return $encodedValue
}

# Initialize an array to store the output data
$outputData = @()

# Loop through each site/library combination in the CSV
foreach ($site in $sitesData) {
    $siteUrl = $site.SiteUrl
    $libraryName = $site.LibraryName

    # Connect to the SharePoint site using app credentials
    try {
        Connect-PnPOnline -Url $siteUrl -ClientId $PnPOnlineAppId -Interactive
        Write-Host "Successfully connected to $siteUrl" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect to $siteUrl" -ForegroundColor Red
        continue
    }

    # Get the document library
    $library = Get-PnPList -Identity $libraryName

    # Check if the library exists
    if ($null -eq $library) {
        Write-Host "The library '$libraryName' does not exist on site '$siteUrl'." -ForegroundColor Red
        Disconnect-PnPOnline
        continue
    }

    # Get the top-level folders from the document library
    $folders = Get-PnPListItem -List $libraryName | Where-Object { $_.FileSystemObjectType -eq "Folder" }

    # Output folder details with SiteId
    foreach ($folder in $folders) {
        # Get Site ID, Web ID, List ID, Folder ID, and other necessary details
        $site = Get-PnPSite
        $tenantId = Get-PnPTenantId 
        $siteId = (Get-PnPSite -Includes ID).Id
        $webId = (Get-PnPWeb).Id
        $listId = (Get-PnPList -Identity $libraryName).Id
        $webUrl = (Get-PnPWeb).Url
        $folderId = $folder.FieldValues["UniqueId"]

        # URL encode and apply custom encoding
        $encodedTenantId = CustomUrlEncode($tenantId)
        $encodedSiteId = CustomUrlEncode("{$siteId}")
        $encodedWebId = CustomUrlEncode("{$webId}")
        $encodedListId = CustomUrlEncode($listId)
        $encodedFolderId = CustomUrlEncode($folderId)
        
        # Use the special encoding for webUrl to handle dots
        $encodedWebUrl = CustomUrlEncodeForWebUrl($webUrl)

        # Construct the final LibraryID string
        $libraryIDString = "tenantId=$encodedTenantId&siteId=$encodedSiteId&webId=$encodedWebId&listId=$encodedListId&folderId=$encodedFolderId&webUrl=$encodedWebUrl&version=1"

        # Store the output data in the array with a single column `LibraryID`
        $outputData += [pscustomobject]@{
            SharepointUrl = $webUrl
            Foldername = $($folder.FieldValues.Title)
            LibraryID = $libraryIDString
        }

        Write-Host "Processed data for folder: $($folder.FieldValues.Title)" -ForegroundColor Cyan
    }

    
}

# Export the output data to CSV with the column name `LibraryID`
$outputData | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding unicode

Write-Host "Data successfully exported to $outputCsvPath" -ForegroundColor Green
# Disconnect PnPOnline
Disconnect-PnPOnline