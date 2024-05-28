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

$Logfile = ""
$UserTime = Get-Date
$MsalResponse = Get-MsalToken @MsalParams
$EWSAccessToken  = $MsalResponse.AccessToken

Function Get-Token{
$MsalResponse = Get-MsalToken @MsalParams -ForceRefresh
$EWSAccessToken  = $MsalResponse.AccessToken
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EWSAccessToken
}
Function Check-Token{
$currentdate = Get-Date
$tokentime = $MsalResponse.ExpiresOn.DateTime.AddHours(+3)
$compare = ($tokentime - $currentdate)
if($compare.Minutes -lt 20)
{
Get-Token
}
}

$query = "((CustomAttribute2 -eq 'Sales Turkey') -or (CustomAttribute2 -eq 'Sales))"
$LogFile = 'C:\Temp\deletednewitemcount.log'
$mbx = Get-Recipient -Filter $query -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox  -ResultSize Unlimited
$mbx = $mbx | Select-Object -Index @(1..200)
Do
{
$mbx | % {
Write-Host "Mailbox : " $_.PrimarySMTPAddress -ForegroundColor Green
Write-Host "--------------------------------------------------------" -ForegroundColor Green
$Eversion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2015
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($Eversion)
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EWSAccessToken
$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $($_.PrimarySMTPAddress)) 
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

$Folderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$Folderview.PropertySet = $propertySet
$Folderview.PropertySet.Add([Microsoft.Exchange.Webservices.Data.FolderSchema]::DisplayName)
$Folderview.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep
$FolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Note")
$FoldersResult = $Service.FindFolders([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot,$FolderSearchFilter, $Folderview)
$pageSize = 1000
$offset = 0
$fcount = 0

$view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList ($pageSize + 1), $offset
$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, [DateTime]"2024-3-3")

$viewcount = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView(1000000)
$totalitems = 0
$allItemsarray = @()
$starttime = get-date
$FoldersResult | % {
try
{
$allItems = @()
do
{
try
{
    $processTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $items = $FoldersResult.Folders[$fcount].FindItems($searchFilter,$view)

    if ($items.MoreAvailable)
    {
        $view.Offset += $pageSize
    }

    $items.Items | ForEach-Object {
        $allItems += $PSItem
        $allItemsarray += $PSItem
    }
    $processTimer.Stop()
    write-host $processTimer.Elapsed $allItemsarray.Count
    Check-Token
    }

catch
{
Get-Token
}
}
while ($items.MoreAvailable)
Write-Host "Folder Name : (" $FoldersResult.Folders[$fcount].DisplayName ") Items Count: " $allItems.Count
$fcount++
$totalitems += $allItems.Count
}
catch
{
Get-Token
}
}
$logtext = $_.PrimarySMTPAddress + " - Total item count: (" + $totalitems + ")" + " - Start time : " + $starttime
Write-Host "--------------------------------------------------------" -ForegroundColor Green
Write-Host $logtext -ForegroundColor Green
Write-Host "--------------------------------------------------------" -ForegroundColor Green
Add-content $LogFile -value $logtext
$silinen = 0
try
{
foreach ($item in $allItemsarray) 
{
try
{
Check-Token
$message = [Microsoft.Exchange.WebServices.Data.Item]::Bind($Service, $item.Id, $propertyset)
$message.Delete('HardDelete')
$silinen++
Write-Host $silinen " :|---|: " $totalitems -ForegroundColor Green
}
catch
{
Get-Token
}
}
}
catch
{
Get-Token
}
$endtime =  get-date
$logtext = $_.PrimarySMTPAddress + " - Total deleted item count : (" + $silinen + ") End time : " + $endtime
Write-Host "--------------------------------------------------------" -ForegroundColor Green
Write-Host $_.PrimarySMTPAddress " - Total deleted item count : ("  $silinen  ")"
Write-Host "--------------------------------------------------------" -ForegroundColor Green
Add-content $LogFile -value $logtext
Start-Sleep -Seconds 5
Get-Token
}
}
while (1 -gt 0)
