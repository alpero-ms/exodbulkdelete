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

5. To install EWS Managed API from powershell

  	Register-PackageSource -provider NuGet -name nugetRepository -location https://www.nuget.org/api/v2
  	Install-Package Exchange.WebServices.Managed.Api
   
7.  The default lifetime of SAML token is 1 hour, after 1 hour you will get an error "401 unauth". You can increaase SAML roken lifetime for Azure App.

     https://learn.microsoft.com/en-us/entra/identity-platform/configurable-token-lifetimes

9. Do not apply this script in a production environment without first testing it in a demo environment. The author disclaims any responsibility for damages or data loss resulting from the use of this script.

10. I've deleted more than 5M emails in 4 days with this script (I have used 2 Azure app with 12 PS ISE Screen) on my CX tenant.
11. For Archive mailbox change "MsgFolderRoot" as "ArchiveMsgFolderRoot"
    $FoldersResult = $Service.FindFolders([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$FolderSearchFilter, $Folderview)



