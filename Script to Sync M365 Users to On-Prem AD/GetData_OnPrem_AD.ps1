#######################################################################
# Name: GetData_OnPrem_AD.ps1
# This script to get data from on-premise AD
# Written By: Samir Makwana (makwanasamir@hotmail.com)
# Last Modified: 12-April-2020
#######################################################################

#To prompt explorer to choose folder where Attributes.csv file is located
$browse = New-Object System.Windows.Forms.FolderBrowserDialog
$browse.ShowDialog()
$FolderPath = $Browse.SelectedPath
$ADUserDataFilePath = "$FolderPath\AADUserDataFilePath.txt"
$ADAttributeFilePath = "$FolderPath\Attributes.csv"

#Importing on-prem AD powershell module
Import-Module ActiveDirectory

#Fetching all AD Users
#$ADUsers = Get-ADUser -Filter * -Properties * -SearchBase "OU=UserMailboxes,OU=O365,DC=O365HybridDemo,DC=com"
$ADUsers = Get-ADUser -Filter * -Properties *

#Importing AD attribute csv file into arrary
$ADAttributes = Import-Csv $ADAttributeFilePath
$ADAttributeCount = $ADAttributes.Count

#defining varible by setting thier values to 0 or $Null
$ADAttrCounter = 0
[String]$Data = $null
[string]$email = $null
$value = $null
$ADUserData = New-Object System.Collections.ArrayList
$TitleString = $null
$UserCounter = 0

#Converting AADAttributes array to string by joining array values with | sign (Value1|value2|Value2...)
$TitleString = $ADAttributes.ADAttribute -join "|"
"$TitleString" | Out-File $ADUserDataFilePath -Append

#Section to start processing AD users one by one
Foreach($ADUser in $ADUsers)
{
    #Resetting variable values to 0 or $Null
    $VariableName=$null
    [string]$email = $null
    $value = $null
    $ADAttrCounter = 0
    $ADUserData = New-Object System.Collections.ArrayList

    $UserCounter++   
    Write-Host "[$UserCounter]" $ADUser.UserPrincipalName

    # Fetching value of each attributes listed in ADAttributes variable
    While($ADAttrCounter -lt $ADAttributeCount)
    {
        $VariableName=$null

        # We are storing ADAttribute name in to variable
        $VariableName = $ADAttributes[$ADAttrCounter].ADAttribute

        # if attribute name is ProxyAddresses and ProxyAddresses1 then fetch its value(Arrary) and join it by ; to convert it to string
        If($VariableName -like "ProxyAddresses" -or $VariableName -like "ProxyAddresses1")
        {
            $Value = $ADUser.ProxyAddresses -join ";"      
        }
        # If attribute name is TargetAddress then fetch its value(Arrary) and join it by ; to convert it to string
        elseIf($VariableName -like "targetaddress")
        {
            $Value = $ADUser.targetaddress -join ";"      
        }
        # If attribute name is ObjectGUID then convert it to Base64 (ImmutableID) format
        elseIf($VariableName -like "ObjectGUID")
        {
            $Value = [system.convert]::ToBase64String($ADUser.objectGUid.ToByteArray())
        }
        #If name of the attribute does not match any of the above condition then
        else
        {
            $Value = $ADUser.$VariableName
        }
        
        #Add value of attribute stored in custom arrary.
        # > $Null will help hide return value at the end of execution
        $ADUserData.Add($Value) > $null

        #While loop counter increment
        $ADAttrCounter++
    }

#Joining customer arrary using | to convert arrary into string value to be exported in txt file
$Data = $ADUserData -join "|"

#Exported ADuser data in single output file. Which will be converting into CSV and used as input file to generate LDF file.
$Data | out-File -FilePath $ADUserDataFilePath -Append

#Write-Host "==============================="

}