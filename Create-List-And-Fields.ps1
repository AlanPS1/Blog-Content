<#
.SYNOPSIS
Create-List-And_fields.ps1 is used to create a SharePoint List and poulate the 14 fields required for lab on AlanPs1.io

.DESCRIPTION 
This script creates a SharePoint List and adds the 14 fields that are a single line of text

.OUTPUTS
No Outputs that are not the columns headers

.EXAMPLE
.\Create-List-And_fields.ps1

.EXAMPLE
.\Create-List-And_fields.ps1 -URL "https://XXXXXX.sharepoint.com/sites/YourSiteCollection"

.EXAMPLE
.\Create-List-And_fields.ps1 -Verbose

.LINK
https://www.alanps1.io/power-platform/flow/flow-create-a-sharepoint-list-with-pnp-powershell-or-the-gui-part-3

.NOTES
Written by AlanPs1

Version history:
V1.00, 17/02/2019 - Initial version - AlanPs1

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
    [ValidateNotNullOrEmpty()]
    [string]$URL = "https://XXXXXX.sharepoint.com/sites/YourSiteCollection"
)

# Check if SharePointPnPPowerShellOnline module is available
if (-not(Get-Module -Name SharePointPnPPowerShellOnline)) {
    
}

else {
    Write-Host "We can't find SharePointPnPPowerShellOnline module so now we'll install it"
    Install-Module SharePointPnPPowerShellOnline
}

# Connect to SharePointPnPPowerShellOnline
$credentials = Get-Credential
Connect-PnPOnline -Url $URL -Credentials $credentials

# Set the $ListName variable
$ListName = 'Get Messages From Title'

# Create the list called "Get Messages From Title"
New-PnPList -Title $ListName -Template GenericList -Url "Lists/$ListName" -OnQuickLaunch -ErrorAction Continue

# Add the 14 fields listed in the blog
Add-PnPField -List $ListName -DisplayName "userPrincipalName" -InternalName "userPrincipalName" -Type Text -AddToDefaultView
Add-PnPField -List $ListName -DisplayName "subject" -InternalName "subject" -Type Text -AddToDefaultView
Add-PnPField -List $ListName -DisplayName "messageId" -InternalName "messageId" -Type Text -AddToDefaultView
Add-PnPField -List $ListName -DisplayName "hasAttachments" -InternalName "hasAttachments" -Type Text -AddToDefaultView
Add-PnPField -List $ListName -DisplayName "attachmentId" -InternalName "attachmentId" -Type Text -AddToDefaultView
Add-PnPField -List $ListName -DisplayName "status" -InternalName "status" -Type Text -AddToDefaultView
Add-PnPField -List $ListName -DisplayName "sender" -InternalName "sender" -Type Text -AddToDefaultView
Add-PnPField -List $ListName -DisplayName "createdDateTime" -InternalName "createdDateTime" -Type Text
Add-PnPField -List $ListName -DisplayName "receivedDateTime" -InternalName "TestColumn" -Type Text
Add-PnPField -List $ListName -DisplayName "sentDateTime" -InternalName "receivedDateTime" -Type Text
Add-PnPField -List $ListName -DisplayName "from" -InternalName "from" -Type Text
Add-PnPField -List $ListName -DisplayName "to" -InternalName "to" -Type Text
Add-PnPField -List $ListName -DisplayName "ccRecipients" -InternalName "ccRecipients" -Type Text
Add-PnPField -List $ListName -DisplayName "bccRecipients" -InternalName "bccRecipients" -Type Text

# Disconnect from SharePointPnPPowerShellOnline
Disconnect-PnPOnline