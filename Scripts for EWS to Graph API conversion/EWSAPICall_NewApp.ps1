#Provide your Office 365 Tenant Id or Tenant Domain Name
$tenantID = "<Tenant ID>"
    
#Used the Microsoft Graph PowerShell app id. You can create and use your own Azure AD App id if needed.
$clientID=<Client ID>
$ClientSecret = ConvertTo-SecureString "<Client Secret>" -AsPlainText -Force
$Scopes   = ”https://outlook.office365.com/.default”  

$MsalResponse = Get-MsalToken -clientID $clientID -clientSecret $clientSecret -tenantID $tenantID -Scopes $Scopes

$EWSAccessToken101  = $MsalResponse.AccessToken

Import-Module 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
 
#Provide the mailbox id (email address) to connect
$MailboxName101 ="PSMTP Address of the mailbox in question"

$Service101 = $null 
$Service101 = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
#$global:exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)
 
#Use Modern Authentication
$Service101.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EWSAccessToken101
#$exchangeService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $EWSAccessToken11
 
#Check EWS connection
$Service101.Url = "https://outlook.office365.com/EWS/Exchange.asmx"

$Service101.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName101)

$Service101.HttpHeaders.Add("X-AnchorMailbox", $MailboxName101);

$MailboxRootid101 = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName101)

$MailboxRoot101 = $null
$MailboxRoot101 = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service101,$MailboxRootid101)
$MailboxRoot101
#$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName) 
#$Inbox =  [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWS,$folderid)
#return $Inbox