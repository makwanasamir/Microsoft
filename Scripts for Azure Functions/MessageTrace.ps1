using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

#Message Trace Code
        $Status1 = $Null
        $ReceivedTime1 = $Null
        $Sender1 = $Null
        $Recipient1 = $Null
        $Subject1 = $Null
$tenantId = $env:TenantID
$tenant = $env:Tenant
Write-Host $tenant
$AdminUserName = $env:AutomationAdmin_UserName
$Password = $env:AutomationAdmin_Password
$SenderEmailAddress = $request.body.SenderEmail
Write-Host $SenderEmailAddress
$Sender1 = $SenderEmailAddress
$RecipientEmailAddress = $request.body.RecipientEmail
Write-Host $RecipientEmailAddress
$Recipient1 = $RecipientEmailAddress

[DateTime] $ExpectedDeliveryDate = $request.body.ExpectedDateofDelivery
Write-Host $ExpectedDeliveryDate
$Date = Get-Date $ExpectedDeliveryDate
Write-Host $Date

$EmailSubject = $request.body.EmailSubject
Write-Host $EmailSubject

$MessageTraceStartDate = (Get-Date $ExpectedDeliveryDate).AddDays(-2)
$MessageTraceEndDate = (Get-Date $ExpectedDeliveryDate).AddDays(2)

$securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($AdminUserName, $securePassword)

Try
{
    #Connect-ExchangeOnline -Credential $credential -TenantId $tenantId
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    $ConnectStatus = "ConnectEXOSuccess"
}
Catch
{
    $ConnectStatus = "ConnectEXOFailed"
}

Try
{
    $AllEmail = Get-MessageTrace -SenderAddress $SenderEmailAddress -RecipientAddress $RecipientEmailAddress -StartDate $MessageTraceStartDate -EndDate $MessageTraceEndDate | where{$_.Subject -like "*$EmailSubject*"}
    #$AllEmail = Get-MessageTrace -SenderAddress $SenderEmailAddress -RecipientAddress $RecipientEmailAddress -StartDate $Date -EndDate $Date | where{$_.Subject -like "*$EmailSubject*"}
    
    Write-Host $AllEmail
    Write-Host $AllEmail.count
    

If($AllEmail.count -ge 1)
{
    Foreach($Email in $AllEmail)
        {
        $Status1 = $AllEmail.Status
        $ReceivedTime1 = $AllEmail.Received
        $Sender1 = $AllEmail.SenderAddress
        $Recipient1 = $AllEmail.RecipientAddress
        $Subject1 = $AllEmail.Subject

        If($Status1 -like "Delivered")
        {
            Write-Host "Email from "+ $Sender1 +" to " $Recipient1 + " was " + $Status1 + " on " + $ReceivedTime1
            $MessageToUser =  "We found your email from "+ $Sender1 +" to " + $Recipient1 + " was " + $Status1 + " on " + $ReceivedTime1
        }
        else
        {
            Write-Host "Email from "+ $Sender1 +" to " $Recipient1 + " was " + $Status1 + " on " + $ReceivedTime1
            $MessageToUser =  "We found your email from "+ $Sender1 +" to " + $Recipient1 + " it was " + $Status1 + " on " + $ReceivedTime1 + "Please check your JunkEmail/Spam folder"
        }
        }
}
elseif($AllEmail.count -eq 1)
{
        $Status1 = $AllEmail.Status
        $ReceivedTime1 = $AllEmail.Received
        $Sender1 = $AllEmail.SenderAddress
        $Recipient1 = $AllEmail.RecipientAddress
        $Subject1 = $AllEmail.Subject

        If($Status1 -like "Delivered")
        {
            Write-Host "Email from "+ $Sender1 +" to " $Recipient1 + " was " + $Status1 + " on " + $ReceivedTime1
            $MessageToUser =  "We found your email from "+ $Sender1 +" to " + $Recipient1 + " was " + $Status1 + " on " + $ReceivedTime1
        }
        else
        {
            Write-Host "Email from "+ $Sender1 +" to " $Recipient1 + " was " + $Status1 + " on " + $ReceivedTime1
            $MessageToUser =  "We found your email from "+ $Sender1 +" to " + $Recipient1 + " it was " + $Status1 + " on " + $ReceivedTime1 + ". Please check your JunkEmail/Spam folder"
        }
}
else
{
    $MessageToUser = "We searched backend for your missing email. However, we did not find any email from "+ $Sender1 +" to the " + $Recipient1 + " on the expected delivery date: " + $ExpectedDeliveryDate
}

    $GetMessageTraceStatus = "GetMessageTraceSuccess"
}
Catch
{
    $GetMessageTraceStatus = "GetMessageTraceFailed"
    $MessageToUser = "MessageToUserFailed"
}


Try
{
    Disconnect-ExchangeOnline -Confirm:$false
    $DisconnectStatus = "DisconnectEXOSuccess"
}
Catch
{
    $DisconnectStatus = "DisconnectEXOFailed"
}

If($MessageToUser -eq $Null)
{
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    #Body = $(Get-Module -ListAvailable | Select-Object Name, Path)
    Body = @{“ConnectStatus” = $ConnectStatus; “GetMessageTraceStatus” = $GetMessageTraceStatus; “MessageToUser” = $MessageToUser; "DisconnectStatus" = $DisconnectStatus}
})
}
else
{
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    #Body = $(Get-Module -ListAvailable | Select-Object Name, Path)
    Body = @{“ConnectStatus” = $ConnectStatus; “GetMessageTraceStatus” = $GetMessageTraceStatus; “MessageToUser” = $MessageToUser; "DisconnectStatus" = $DisconnectStatus}
})
}