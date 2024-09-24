#######################################################################
# Name: GetData_AAD_MSOL.ps1
# This script to get data from Azure AD, MSOL and Exchange Online
# Written By: Samir Makwana (makwanasamir@hotmail.com)
# Last Modified: 12-April-2020
#######################################################################

#To prompt explorer to choose folder where Attributes.csv file is located
$browse = New-Object System.Windows.Forms.FolderBrowserDialog
$browse.ShowDialog()
$FolderPath = $Browse.SelectedPath
$AADUserDataFilePath = "$FolderPath\AADUserDataFilePath.txt"
$AADAttributeFilePath = "$FolderPath\Attributes.csv"


#Prompt for credential
Write-Host "You will be prompted for credential for O365 (Atleast Read Only Admin)" -ForegroundColor Green
$Cred = Get-Credential


#To connect MSOlService
Connect-MsolService -Credential $Cred


#To Connect AzureAD
Connect-AzureAD -Credential $Cred


#Create session with Exchange Online and Import session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
Import-PSSession $Session


#Get all Azure AD users. 
# We can use filter with this command to narraow the list to target specific users
$AADUsers = Get-AzureADUser


#Import list of attributes to fetch for each Azure AD user
$AADAttributes = Import-Csv $AADAttributeFilePath
$AADAttributeCount = $AADAttributes.Count


# Setting variable values to 0 or $Null
$AADAttrCounter = 0
[String]$Data = $null
[string]$email = $null
$value = $null
$AADUserData = New-Object System.Collections.ArrayList
$TitleString = $null
$UserCounter = 0


#Converting AADAttributes array to string by joining array values with | sign (Value1|value2|Value2...)
$TitleString = $AADAttributes.O365Attribute -join "|"
"$TitleString" | Out-File $AADUserDataFilePath -Append


#Section to start processing Azure AD users one by one
Foreach($AADUser in $AADUsers)
{
    $UserCounter++   
    Write-Host "[$UserCounter]" $AADUser.UserPrincipalName

    #Resetting variable values to 0 or $Null
    $VariableName=$null
    [string]$email = $null
    $value = $null
    $AADAttrCounter = 0
    $AADUserData = New-Object System.Collections.ArrayList

    
    # Fetching value of each attributes listed in AADAttributes variable
    While($AADAttrCounter -lt $AADAttributeCount)
    {
        $VariableName=$null

        # We are storing AADAttribute (O365) name in to variable
        $VariableName = $AADAttributes[$AADAttrCounter].O365Attribute

        #If attribute name is proxyaddresses then join arrary with ; to convert it to string value
        If($VariableName -like "ProxyAddresses")
        {
            $Value = $AADUser.ProxyAddresses -join ";"      
        }
        #If attribute value is ShowInAddressList is TRUE and export value as FALSE
        #This is to match ShowInAddressList value with on-prem AD msExchHideFromAddressList value
        # ShowInAddressList (TRUE) = msExchHideFromAddressList (FALSE)
        # ShowInAddressList (FALSE) = msExchHideFromAddressList (TRUE)
        elseIf($VariableName -like "ShowInAddressList")
        {
            #Write-Host "SHowINAddressList:" $AADUser.$VariableName
            if($AADUser.$VariableName -like "FALSE")
            {
               $Value = "TRUE"
               #$value
            }
            else
            {
               $Value = "FALSE"
            }

        }
        # If attribute name is RecipientTypeDetails,RecipientTypeDetails1,RecipientTypeDetails2 then
        # use Exchange online session to run get-recipient command and value of RecipientTypeDetails
        elseIf($VariableName -like "RecipientTypeDetails" -or $VariableName -like "RecipientTypeDetails1" -or $VariableName -like "RecipientTypeDetails2")
        {
            If($GetRecipient = Get-Recipient $AADUser.UserPrincipalName -ErrorAction SilentlyContinue)
            {
            $Value = $GetRecipient.RecipientTypeDetails
            }
            else
            {
            $Value = "NA"
            }
        }
        # value for onPremisesDistinguishedName shows value of on-prem AD OU where that user is located and synced to O365
        elseIf($VariableName -like "onPremisesDistinguishedName")
        {
            $value = $AADUser.ExtensionProperty.onPremisesDistinguishedName
        }
        #If name of the attribute does not match any of the above condition then
        else
        {
            $Value = $AADUser.$VariableName
        }

        #Add value of attribute stored in custom arrary.
        # > $Null will help hide return value at the end of execution
        $AADUserData.Add($Value) > $null

        #Increment counter for While loop
        $AADAttrCounter++
    }

#Joining customer arrary using | to convert arrary into string value to be exported in txt file
$Data = $AADUserData -join "|"

#Exported ADuser data in single output file. Which will be converting into CSV and used as input file to generate LDF file.
$Data | out-File -FilePath $AADUserDataFilePath -Append

#Write-Host "==============================="

}