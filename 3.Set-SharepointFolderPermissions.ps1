<#
        .SYNOPSIS
        Set permissions on folders in a document library on a Sharepoint Site

        .DESCRIPTION
        Script is used to Set permissions on folders in a document library on a Sharepoint Site from a csv file

        CSV file format example
        EntraIdSecGroupName;FolderName;SiteURL;ListName;ShortURL;SharepointDocumentShareName;Sharepoint_Permission
        SP_Finance_Department_Finance_Suppliers;Finance Suppliers;Dokumenter;https://contoso.sharepoint.com/sites/Finance_Department;/sites/Finance_Department;Delte Dokumenter;Må bidrage
        SP_Finance_Department_Invoices;Finance Invoices;Dokumenter;https://contoso.sharepoint.com/sites/Finance_Department;/sites/Finance_Department;Delte Dokumenter;Må bidrage
        
        EntraIdSecGroupName; (EntraID Security Group to be set as contributor)
        FolderName; (Folder to have the security group added to)
        SiteURL; (Sharepoint URL for Sharepoint Site)
        ListName; (Sharepoint List name example is for a danish setup for other language change this)
        ShortURL; (Used in section for setting the permisson on the correct document folder.)
        SharepointDocumentShareName; (The default document library in sharepoint in danish Delte Dokumenter in english Shared Documents)
        Sharepoint_Permission; (Set the permission for the group in danish Må bidrag in english translate into Contributor)

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

        .PARAMETER Extension

        .EXAMPLE
        C:\PS> 3.Set-SharepointFolderPermissions.ps1 -csvFilePath "C:\script\Sharepoint permissions.csv" -PnPOnlineAppId a4339632-4336-4337-833a-63f0332833e

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
    [string]$PnPOnlineAppId
)

# Import the CSV file
$folders = Import-Csv $csvFilePath -Delimiter ";"

#Connect to Azure AD with Admin account
Import-Module AzureAD -UseWindowsPowerShell
Connect-AzureAD

# Loop through each row in the CSV
foreach ($folder in $folders) {
    $SiteURL = $folder.SiteURL
    $ListName = $folder.ListName
    $FolderName = $folder.FolderName
    $ShortURL = $folder.ShortURL
    $EntraIdSecGroupName = $folder.EntraIdSecGroupName
    $SharepointDocumentShareName =  $folder.SharepointDocumentShareName
    $Sharepoint_Permission =  $folder.Sharepoint_Permission
    
    try {
        
        # Get the Object ID for the Azure AD Group
        $EntraIDGroupId = Get-AzureADGroup -Filter "DisplayName eq '$EntraIdSecGroupName'" | Select-Object -ExpandProperty ObjectId
        
        # Connect to the SharePoint Online site
        Connect-PnPOnline -Url $SiteURL -Interactive -ClientID $PnPOnlineAppId

        # Get the folder from the server-relative URL
        $FolderServerRelativeURL = "$ShortURL/$SharepointDocumentShareName/$FolderName"
        $Folder = Get-PnPFolder -Url $FolderServerRelativeURL

        # Grant permission to the folder
        Write-Host "Setting permission on $FolderServerRelativeURL" -ForegroundColor Cyan
        Set-PnPListItemPermission -List $ListName -Identity $Folder.ListItemAllFields -User "c:0t.c|tenant|$EntraIDGroupId" -AddRole "$Sharepoint_Permission"

        Write-Host "Permissions successfully set for folder: $FolderServerRelativeURL" -ForegroundColor Green

    } catch {
        Write-Host "Error setting permissions for folder: $FolderServerRelativeURL" -ForegroundColor Red
        Write-Host "Error details: $_" -ForegroundColor Red

        Disconnect-AzureAD
        Disconnect-PnPOnline 
    }
}
Disconnect-AzureAD
Disconnect-PnPOnline 