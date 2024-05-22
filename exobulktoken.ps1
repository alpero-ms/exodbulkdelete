<#

Script Name: EXO Mail Cleanup Tool
Author: Alper Özdemir - alper@cloudvision.com.tr
Date: 21.5.2024
Description: Deleting emails through EXO can be challenging in certain scenarios, especially since the new-complianceSearch feature can only delete a maximum of 10 emails at a time. 
You can use EWS to access and delete all emails older than a specific date across all folders.


1. Create an Azure App registration with the following API permissions:

	Office 365 Exchange Online
	EWS.AccessAsUser.All
	full_access_as_app
	Mail.ReadWrite.All
	Grant admin consent

2. Edit exobulk.ps1 xxx values:

	$TenantId = "xxx.onmicrosoft.com"
	$AppClientId="xxx"
	$ClientSecret = (ConvertTo-SecureString 'xxx' -AsPlainText -Force)
	$mailbox = @(alper@cloudvision.com.tr)

3. It only deletes mail items (not calendar or contact items) for a certain age, and this behavior can be controlled with the following lines:

	$FolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Note")	
	$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, [DateTime]"2024-6-6")

	https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/folders-and-items-in-ews-in-exchange

4. If you want to permanently delete the emails, do not forget to disable the Single Item Recovery (SIR) feature beforehand.

5. Do not apply this script in a production environment without first testing it in a demo environment. The author disclaims any responsibility for damages or data loss resulting from the use of this script.

#>

Import-Module C:\lib\net35\Microsoft.Exchange.WebServices.dll
Install-Module -Name MSAL.PS

$TenantId = "xxx.onmicrosoft.com"
$AppClientId="xxx"
$MsalParams = @{
    ClientId = $AppClientId
    TenantId = $TenantId   
    Scopes   = "https://outlook.office365.com/.default"
    ClientSecret = (ConvertTo-SecureString 'xxx' -AsPlainText -Force)   
}

$Logfile = '*.log'
$totalitem = 0
$TokenDateTime = Get-Date -DisplayHint Time
$MsalResponse = Get-MsalToken @MsalParams
$EWSAccessToken  = $MsalResponse.AccessToken

Function Clear-MsalTokenCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch] $FromDisk
    )

    if ($FromDisk) {
        $TokenCachePath = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "MSAL.PS\MSAL.PS.msalcache.bin3"
        if (Test-Path $TokenCachePath) { Remove-Item -LiteralPath $TokenCachePath -Force }
    }
    else {
        $script:PublicClientApplications = New-Object 'System.Collections.Generic.List[Microsoft.Identity.Client.IPublicClientApplication]'
        $script:ConfidentialClientApplications = New-Object 'System.Collections.Generic.List[Microsoft.Identity.Client.IConfidentialClientApplication]'
    }
}

Function Get-Token{
$MsalResponse = Get-MsalToken @MsalParams
$EWSAccessToken  = $MsalResponse.AccessToken
$TokenDateTime = Get-Date -DisplayHint Time
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EWSAccessToken
}

$Usr = get-mailbox alper@cloudvision.com.tr
$usr | % {
Write-Host $_.PrimarySMTPAddress -ForegroundColor Green
$LogFile = 'C:\Temp\' + $_.PrimarySMTPAddress. + 'log'
$before = (Get-MailboxFolderStatistics $_.PrimarySMTPAddress | where {$_.ContainerClass -eq 'IPF.Note'} | Measure-Object -Sum -Property ItemsInFolder).Sum
$Eversion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2015
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($Eversion)
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EWSAccessToken
$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$mailbox = @($_.PrimarySMTPAddress )
$Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $($mailbox)) 
$Folderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$Folderview.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.Webservices.Data.BasePropertySet]::FirstClassProperties)
$Folderview.PropertySet.Add([Microsoft.Exchange.Webservices.Data.FolderSchema]::DisplayName)
$Folderview.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep
$FolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Note")
$FoldersResult = $Service.FindFolders([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot,$FolderSearchFilter, $Folderview)
$count = 50000
$view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $count
$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, [DateTime]"2024-4-4")
$fcount = 0
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
$FoldersResult | % {
do
{
try
{
$findItemsResults = $FoldersResult.Folders[$fcount].FindItems($searchFilter,$view)
Write-Host "item sayisi : "$findItemsResults.Items.Count "Folder ismi : " $FoldersResult.Folders[$fcount].DisplayName
}
catch
{
Clear-MsalTokenCache
Get-Token
}
foreach ($item in $findItemsResults.Items) {
try
{
$message = [Microsoft.Exchange.WebServices.Data.Item]::Bind($Service, $item.Id, $propertyset)
$message.Delete('HardDelete')
$currentdate = Get-Date
$tokentime = $MsalResponse.ExpiresOn.DateTime.AddHours(+3)
$compare = ($tokentime - $currentdate)
if($compare.Minute -lt 20)
{
Clear-MsalTokenCache
Get-Token
}
}
catch
{
}
}
}
while ($findItemsResults.Items.Count -gt 0)
$fcount++
}
$after = (Get-MailboxFolderStatistics $_.PrimarySMTPAddress | where {$_.ContainerClass -eq 'IPF.Note'} | Measure-Object -Sum -Property ItemsInFolder).Sum
$logvalue = $TokenDateTime.ToString() + " - " + $_.PrimarySMTPAddress + " - Item count before deletion : (" + $before + ") after deletion (" + $after + ")"
Add-content $Logfile -value $logvalue
}

