# Define script parameters
param (
    [Parameter(Mandatory=$true)][string]$DisplayName,  # Display name for the group
    [Parameter(Mandatory=$true)][string]$Alias,        # Alias for the group
    [Parameter(Mandatory=$false)][string]$Owners,      # Comma-separated list of owners
    [Parameter(Mandatory=$false)][string]$AccessType = "Public",  # AccessType (Public/Private)
    [Parameter(Mandatory=$true)][string]$TenantName,   # Tenant name to construct SharePoint site URL
    [Parameter(Mandatory=$false)][string[]]$Folders,    # Array of folder names
    [Parameter(Mandatory=$false)]$PnPOnlineAppId    # PnP Online app id
)

# Ensure the necessary modules are installed and imported
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Verbose
}
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Verbose
}
Import-Module ExchangeOnlineManagement -Verbose
Import-Module PnP.PowerShell -Verbose

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..."
$UserCredential = Get-Credential
Connect-ExchangeOnline -UserPrincipalName $UserCredential.UserName -ShowProgress $true

# Create the Microsoft 365 Group (Unified Group)
try {
    Write-Host "Creating Microsoft 365 Group..."
    $group = New-UnifiedGroup -DisplayName $DisplayName `
                              -Alias $Alias `
                              -AccessType $AccessType

    Write-Host "Microsoft 365 Group '$DisplayName' created successfully."

    # Add Owners to the Group
    if ($Owners) {
        $OwnerArray = $Owners -split ","
        foreach ($Owner in $OwnerArray) {
            Add-UnifiedGroupLinks -Identity $Alias -LinkType Owner -Links $Owner
            Write-Host "Added owner: $Owner"
        }
    }

    # Connect to SharePoint Online (PnP PowerShell) to create folders
    $siteUrl = "https://$TenantName.sharepoint.com/sites/$Alias"
    Write-Host "Connecting to SharePoint site: $siteUrl $PnPOnlineAppId"
    #Connect-PnPOnline -url $siteUrl -Interactive -ClientId $PnPOnlineAppId  -Verbose
    Connect-PnPOnline $siteUrl -ClientId $PnPOnlineAppId -Interactive
    #Connect-PnPOnline https://hcconsultantcy.sharepoint.com/sites/st_eftermarked -ClientId 36125c42-b99a-44d3-a829-2c5cd7449787 -Interactive


    # Create folders from the provided array
    if ($Folders) {
        Write-Host "Creating folders in the SharePoint site..."
        foreach ($folderName in $Folders) {
            Add-PnPFolder -Name $folderName -Folder "Shared Documents"
            Write-Host "Created folder: $folderName"
        }
    }

    Write-Host "Folders created successfully in the SharePoint site."

} catch {
    Write-Error "Error creating Microsoft 365 Group or folders: $_"
} finally {
    # Disconnect from SharePoint Online
    Disconnect-PnPOnline
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
