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
$List1Name = 'Get Messages From Title'
$List2Name = 'Get Attachment IDs'

# Create the list called "Get Messages From Title"
New-PnPList -Title $List1Name -Template GenericList -Url "Lists/$List1Name" -OnQuickLaunch -ErrorAction Continue

# Add the 13 fields listed in the blog
Add-PnPField -List $List1Name -DisplayName "userPrincipalName" -InternalName "userPrincipalName" -Type Text -AddToDefaultView
Add-PnPField -List $List1Name -DisplayName "subject" -InternalName "subject" -Type Text -AddToDefaultView
Add-PnPField -List $List1Name -DisplayName "messageId" -InternalName "messageId" -Type Text -AddToDefaultView
Add-PnPField -List $List1Name -DisplayName "hasAttachments" -InternalName "hasAttachments" -Type Text -AddToDefaultView
Add-PnPField -List $List1Name -DisplayName "status" -InternalName "status" -Type Text -AddToDefaultView
Add-PnPField -List $List1Name -DisplayName "sender" -InternalName "sender" -Type Text -AddToDefaultView
Add-PnPField -List $List1Name -DisplayName "createdDateTime" -InternalName "createdDateTime" -Type Text
Add-PnPField -List $List1Name -DisplayName "receivedDateTime" -InternalName "TestColumn" -Type Text
Add-PnPField -List $List1Name -DisplayName "sentDateTime" -InternalName "receivedDateTime" -Type Text
Add-PnPField -List $List1Name -DisplayName "from" -InternalName "from" -Type Text
Add-PnPField -List $List1Name -DisplayName "to" -InternalName "to" -Type Text
Add-PnPField -List $List1Name -DisplayName "ccRecipients" -InternalName "ccRecipients" -Type Text
Add-PnPField -List $List1Name -DisplayName "bccRecipients" -InternalName "bccRecipients" -Type Text

# Create the list called "Get Attachment IDs"
New-PnPList -Title $List2Name -Template GenericList -Url "Lists/$List2Name" -OnQuickLaunch -ErrorAction Continue

# Add the 14 fields listed in the blog
Add-PnPField -List $List2Name -DisplayName "userPrincipalName" -InternalName "userPrincipalName" -Type Text -AddToDefaultView
Add-PnPField -List $List2Name -DisplayName "subject" -InternalName "subject" -Type Text -AddToDefaultView
Add-PnPField -List $List2Name -DisplayName "messageId" -InternalName "messageId" -Type Text -AddToDefaultView
Add-PnPField -List $List2Name -DisplayName "hasAttachments" -InternalName "hasAttachments" -Type Text -AddToDefaultView
Add-PnPField -List $List2Name -DisplayName "attachmentId" -InternalName "attachmentId" -Type Text -AddToDefaultView
Add-PnPField -List $List2Name -DisplayName "status" -InternalName "status" -Type Text -AddToDefaultView
Add-PnPField -List $List2Name -DisplayName "sender" -InternalName "sender" -Type Text -AddToDefaultView
Add-PnPField -List $List2Name -DisplayName "createdDateTime" -InternalName "createdDateTime" -Type Text
Add-PnPField -List $List2Name -DisplayName "receivedDateTime" -InternalName "TestColumn" -Type Text
Add-PnPField -List $List2Name -DisplayName "sentDateTime" -InternalName "receivedDateTime" -Type Text
Add-PnPField -List $List2Name -DisplayName "from" -InternalName "from" -Type Text
Add-PnPField -List $List2Name -DisplayName "to" -InternalName "to" -Type Text
Add-PnPField -List $List2Name -DisplayName "ccRecipients" -InternalName "ccRecipients" -Type Text
Add-PnPField -List $List2Name -DisplayName "bccRecipients" -InternalName "bccRecipients" -Type Text

# Disconnect from SharePointPnPPowerShellOnline
Disconnect-PnPOnline