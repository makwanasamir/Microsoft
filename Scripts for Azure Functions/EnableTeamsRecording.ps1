using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

#Teams Recording Enablement Code
$tenantId = $env:TenantID
$AdminUserName = $env:AutomationAdmin_UserName
$Password = $env:AutomationAdmin_Password
$UserUPN = $request.body.UserUPN
$TeamsMeetingPolicy_RecordingEnabled = $env:TeamsMeetingPolicyName

$securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($AdminUserName, $securePassword)

Try
{
    Connect-MicrosoftTeams -Credential $credential -TenantId $tenantId
    $ConnectStatus = "ConnectTeamsPSSuccess"
}
Catch
{
    $ConnectStatus = "ConnectTeamsPSFailed"
}

Try
{
    $Output = (get-csonlineuser $UserUPN).TeamsMeetingPolicy
    Write-Host "Currently, TeamsMeetingPolicy is $Output for $UserUPN"
    $GetStatus = "GetCSOnlineUserSuccess"
}
Catch
{
    $GetStatus = "GetCSOnlineUserFailed"
}

Try
{
    If($Output -eq $null)
    {
    Write-Host "Currently, TeamsMeetingPolicy is set to '$NULL', hence granting $TeamsMeetingPolicy_RecordingEnabled"
    Grant-CsTeamsMeetingPolicy -PolicyName $TeamsMeetingPolicy_RecordingEnabled -Identity $UserUPN
    }
    else
    {
    Grant-CsTeamsMeetingPolicy -PolicyName $Null -Identity $UserUPN
    Write-Host "Currently, TeamsMeetingPolicy is set to $Output, hence setting policy to $Null before granting appropriate Teams Meeting policy"
    Grant-CsTeamsMeetingPolicy -PolicyName $TeamsMeetingPolicy_RecordingEnabled -Identity $UserUPN
    }
    $GrantStatus = "RecordingEnablementSuccess"
}
Catch
{
    $GrantStatus = "RecordingEnablementFailed"
}

Try
{
    Disconnect-MicrosoftTeams
    $DisconnectStatus = "DisconnectTeamsPSSuccess"
}
Catch
{
    $DisconnectStatus = "DisconnectTeamsPSFailed"
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    #Body = $(Get-Module -ListAvailable | Select-Object Name, Path)
    Body = @{“ConnectStatus” = $ConnectStatus; “GetStatus” = $GetStatus; “GrantStatus” = $GrantStatus; "DisconnectStatus" = $DisconnectStatus}
})
