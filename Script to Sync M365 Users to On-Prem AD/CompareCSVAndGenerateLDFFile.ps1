#######################################################################
# Name: CompareCSVAndGenerateLDFFiles.ps1
# This script to compare CSV files and generate LDF file to import data into AD
# Written By: Samir Makwana (makwanasamir@hotmail.com)
# Last Modified: 12-April-2020
#######################################################################

#To prompt explorer to choose folder where Attributes.csv file is located
$browse = New-Object System.Windows.Forms.FolderBrowserDialog
$browse.ShowDialog()
$FolderPath = $Browse.SelectedPath


#Defining Variables
$MSOLCSVFilePath = "$FolderPath\AADUsersData.csv"
$ADVSCFilePath = "$FolderPath\ADUsersData.csv"
$AttrListFilePath = "$FolderPath\Attributes.csv"
$MatchStats = "$FolderPath\MatchStats.txt"
$LDIFFile = $Null
$Set_AzureADUser_ImmutableIDScript = "$FolderPath\Set_AzureADUser_ImmutableIDScript.txt"
$RightNow = Get-Date -Format "dd_MM_yyyy_HH_mm"
$AADFName = $Null
$AADLName = $Null
$MatchFound = $Null
$AADObjectID = $Null
$ADFName = $Null
$ADLName = $Null
$User = $Null
$ADUser = $Null
$ADAttrCounter = 0
[String]$Data = $null
[string]$email = $null
$value = $null
$DN = $Null
$ChangeType = $Null

#Tenant 
$Tenant = "@O365Hybrid.onmicrosoft.com"
$AADVariableToExclude = @("ObjectID","DirSyncEnabled","ImmutableID","RecipientTypeDetails1","RecipientTypeDetails2")

#Importing CSV Data
$MSOLCSV = Import-Csv $MSOLCSVFilePath
$ADCSV = Import-Csv $ADVSCFilePath
$AttrToCompare = Import-Csv $AttrListFilePath
$ADAttributeCount = $AttrToCompare.Count

Write-Host "Azure AD Object Count:" $MSOLCSV.Count -ForegroundColor Magenta
Write-Host "AD Object Count:" $ADCSV.Count -ForegroundColor Magenta

#Match stats file header
"FirstName|LastName|ObjectID|MatchFound|DistingushedName" | Out-File $MatchStats -Append

#Foreach user in AAD user list
Foreach($AADUser in $MSOLCSV)
{
$MatchFound = $null

$AADFName = $Null
$AADLName = $Null
$AADObjectID = $Null

$AADFName = $AADUser.GivenName
$AADLName = $AADUser.Surname
$AADObjectID = $AADUser.ObjectID

    #If AADUser First and Last name is not null then continue otherwise go to ELSE section
    If($AADFName.length -ne 0 -and $AADLName.length -ne 0)
    {
        #Foreach user in onprem AD user list
        Foreach($ADUser in $ADCSV)
        {
        $ADFName = $Null
        $ADLName = $Null
        $ADDN = $Null

        $ADFName = $ADUser.GivenName
        $ADLName = $ADUser.Sn
        $MatchFound = $TRUE
        $ADDN = $ADUser.DistinguishedName
            
            #If azure AD User First Name and Last Name matches with GivenName and Surname in AD then match found and loop breaks
            #Else continue until end of the user list
            if($AADFName -like $ADFName -and $AADLName -like $ADLName)
            {
                $ADFName = $ADUser.GivenName
                $ADLName = $ADUser.Sn
                $MatchFound = $TRUE
                $ADDN = $ADUser.DistinguishedName

                #$AADFName
                #$AADLName
                #$ADFName
                #$ADLName

                "$AADFName|$AADLName|$AADObjectID|$MatchFound|$ADDN" | Out-File $MatchStats -Append

                Break;
            }
            else
            {
            $MatchFound = $FALSE
            }
        }

    
        If($MatchFound -eq $TRUE)
        {
                Write-Host "============================================="
                Write-Host "$AADFName $AADLName matches $ADFName $ADLName"
                
                $VariableName=$null
                [string]$email = $null
                $value = $null
                $ADAttrCounter = 0
                $DN = $Null
                $ChangeType = $Null
                
                $DN = $ADUser.DistinguishedName

                $UserMail = $AADUser.Mail
                #$UserMail = "samir@o365hybrid.ca"
                $LDIFFile = $FolderPath + "\" + $UserMail +"-"+ $RightNow + ".ldf"
                $LDIFFile

                While($ADAttrCounter -lt $ADAttributeCount)
                {
                    $ADVariableName=$null
                    $ADVariableName = $AttrToCompare[$ADAttrCounter].ADAttribute
                    $AADVariableName=$null
                    $AADVariableName = $AttrToCompare[$ADAttrCounter].O365Attribute
                    
                    #$AADVariableName
                    #$AADUser.$AADVariableName
                    #$ADVariableName
                    #$ADUser.$ADVariableName
                    
                    If(($AADUser.$AADVariableName).length -ne 0 -and $AADVariableToExclude -notcontains $AADVariableName)
                    {
                        If($AADVariableName -like "ProxyAddresses")
                        {
                            $AADProxyAddresses = $AADUser.$AADVariableName -split ";"
                            #Write-Host "AAD Proxy::" $AADProxyAddresses
                            
                            $ADProxyAddresses = $ADUser.$ADVariableName -split ";"
                            #Write-Host "AD Proxy::" $ADProxyAddresses

                            #Write-host $AADProxyAddresses -ForegroundColor Red
                            #Write-Host $ADProxyAddresses -ForegroundColor Green

                            $AllProxyAddresses = $AADProxyAddresses + $ADProxyAddresses
                            #Write-host $AllProxyAddresses.Count -ForegroundColor Magenta
                            #Write-host $AllProxyAddresses -ForegroundColor Cyan

                            $AllProxyAddresses = $AllProxyAddresses | sort -Unique
                            $AllProxyAddresses = $AllProxyAddresses | Where-Object {$_}
                            #Write-host $AllProxyAddresses.Count -ForegroundColor Magenta
                            #Write-host $AllProxyAddresses -ForegroundColor Cyan

                            "dn: $DN" | Out-File $LDIFFile -Append
                            "changetype: modify" | Out-File $LDIFFile -Append
                            "delete: ProxyAddresses" | Out-File $LDIFFile -Append
                            "-" | Out-File $LDIFFile -Append
                            "" | Out-File $LDIFFile -Append
                            "dn: $DN" | Out-File $LDIFFile -Append
                            "changetype: modify" | Out-File $LDIFFile -Append
                            "add: ProxyAddresses" | Out-File $LDIFFile -Append
                            
                            Foreach($Proxy in $AllProxyAddresses)
                            {
                                If($Proxy -notlike "$Tenant")
                                {
                                    "ProxyAddresses: $Proxy" | Out-File $LDIFFile -Append
                                }
                            }
                            "-" | Out-File $LDIFFile -Append
                            "" | Out-File $LDIFFile -Append
                        }
                        elseif($AADVariableName -like "RecipientTypeDetails")
                        {

                            switch ($AADUser.$AADVariableName) 
                            {
                                "UserMailbox"  
                                {
                                        Write-host "onprem AD $ADVariableName value needs to set to 4"
                                        
                                        "dn: $DN" | Out-File $LDIFFile -Append
                                        "changetype: modify" | Out-File $LDIFFile -Append
                                        "replace: msExchRemoteRecipientType" | Out-File $LDIFFile -Append
                                        "msExchRemoteRecipientType: 4" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "replace: msExchRecipientTypeDetails" | Out-File $LDIFFile -Append
                                        "msExchRecipientTypeDetails: 2147483648" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "replace: msExchRecipientDisplayType" | Out-File $LDIFFile -Append
                                        "msExchRecipientDisplayType: -2147483642" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "" | Out-File $LDIFFile -Append
                                    
                                    break
                                }
                                "SharedMailbox" 
                                {
                                        Write-host "onprem AD $ADVariableName value needs to set to 100"
                                        
                                        "dn: $DN" | Out-File $LDIFFile -Append
                                        "changetype: modify" | Out-File $LDIFFile -Append
                                        "replace: msExchRemoteRecipientType" | Out-File $LDIFFile -Append
                                        "msExchRemoteRecipientType: 100" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "replace: msExchRecipientTypeDetails" | Out-File $LDIFFile -Append
                                        "msExchRecipientTypeDetails: 34359738368" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "replace: msExchRecipientDisplayType" | Out-File $LDIFFile -Append
                                        "msExchRecipientDisplayType: -2147483642" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "" | Out-File $LDIFFile -Append
                                    
                                    break
                                }
                                "RoomMailbox"  
                                {       Write-host "onprem AD $ADVariableName value needs to set to 33"
                                        
                                        "dn: $DN" | Out-File $LDIFFile -Append
                                        "changetype: modify" | Out-File $LDIFFile -Append
                                        "replace: msExchRemoteRecipientType" | Out-File $LDIFFile -Append
                                        "msExchRemoteRecipientType: 33" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "replace: msExchRecipientTypeDetails" | Out-File $LDIFFile -Append
                                        "msExchRecipientTypeDetails: 8589934592" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "replace: msExchRecipientDisplayType" | Out-File $LDIFFile -Append
                                        "msExchRecipientDisplayType: -2147481850" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "" | Out-File $LDIFFile -Append
                                    
                                    break
                                }
                                "MailUser"   {"RecipientTypeDetails: Mailuser"; break}
                            }
                          
                        }
                        else
                        {
                            If($AADUser.$AADVariableName -like $ADUser.$ADVariableName)
                            {
                            #Write-Host $AADUser.$AADVariableName "Matches" $ADUser.$ADVariableName
                            }
                            else
                            {
                                #$AADVariableName
                                #$ADVariableName
                                Write-Host $AADUser.$AADVariableName "DoesNotMatches" $ADUser.$ADVariableName
                                        
                                        $Value = $AADUser.$AADVariableName

                                        #"$ADVariableName" | Out-File $LDIFFile -Append
                                        #"$AADVariableName" | Out-File $LDIFFile -Append
                                        "dn: $DN" | Out-File $LDIFFile -Append
                                        "changetype: modify" | Out-File $LDIFFile -Append
                                        "replace: $ADVariableName" | Out-File $LDIFFile -Append
                                        "$ADVariableName : $Value" | Out-File $LDIFFile -Append
                                        "-" | Out-File $LDIFFile -Append
                                        "" | Out-File $LDIFFile -Append

                            }
                        }

                    }
                    else
                    {
                        If($AADVariableName -like "ImmutableID" -and ($AADUser.$AADVariableName).length -eq 0)
                        {
                            $ObjectID = $AADUser.ObjectID
                            $ImmutableID = $ADUser.ObjectGUID
                            "Set-AzureADUser -ObjectId $ObjectID -ImmutableId $ImmutableID" | Out-File $Set_AzureADUser_ImmutableIDScript -Append
                        }
                    #Write-Host "O365 attribute $AADVariableName is NULL/Empty. Hence, no need to compare it with value of on-prem AD"
                    }

                    $ADAttrCounter++
                }


        }
        else
        {
                #Write-Host "$AADFName $AADLName no match found"
                #Write-Host "============================================="
                "$AADFName|$AADLName|$AADObjectID|$MatchFound|$ADDN" | Out-File $MatchStats -Append

        }

    }
}



