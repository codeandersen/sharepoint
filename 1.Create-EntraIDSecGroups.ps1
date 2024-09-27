<#
        .SYNOPSIS
        Creates EntraID Security Groups from a CSV file.

        .DESCRIPTION
        Script is used to create Entra ID Security groups from CSV. Input CSV file needs the following column Security_GroupName;MailNickName;Description
        
        .EXAMPLE
        CSV file format example
        Security_GroupName;MailNickName;Description
        SP_Finance_Department_;SP_HR;Gives access to a folder on Sharepoint site.

        .PARAMETER Name
        -csvFilePath
            Path to CSV file
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false   

        .PARAMETER Extension

        .EXAMPLE
        C:\PS> 1.Create-EntraIDSecGroups.ps1 -csvFilePath C:\script\securitygroups.csv

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        BLOG: http://www.apento.com
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>


param (
    [string]$csvFilePath
)

# Import the AzureAD module
Import-Module AzureAD 

# Connect to Azure AD (you will be prompted to sign in)
Connect-AzureAD

# Check if CSV file parameter is provided and if the file exists
if (-not $csvFilePath) {
    Write-Host "Please provide the path to the CSV file." -ForegroundColor Red
    exit
}

if (-Not (Test-Path -Path $csvFilePath)) {
    Write-Host "The specified CSV file does not exist. Please provide a valid file path." -ForegroundColor Red
    exit
}

# Import the CSV file, specifying that the delimiter is a semicolon (;)
$groups = Import-Csv -Path $csvFilePath -Delimiter ";"

# Loop through each row in the CSV file and create a security group
foreach ($group in $groups) {
    $groupName = $group.Security_GroupName
    $mailNickName = $group.MailNickName
    $description = $group.Description

    try {
        # Attempt to create a new Azure AD Security Group
        New-AzureADGroup -DisplayName "$groupName" -Description "$description" -MailEnabled $false -SecurityEnabled $true -MailNickName $mailNickName
        Write-Host "Successfully created group: $groupName" -ForegroundColor Green
    }
    catch {
        # Catch any errors and display a message
        Write-Host "Failed to create group: $groupName" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Disconnect from Azure AD after script execution
Disconnect-AzureAD

Write-Host "Group creation process completed." -ForegroundColor Cyan
