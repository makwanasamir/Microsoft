
#######################################################################
# Name: RunMultipleLDFFiles.ps1
# This script is to Generate LDF files to revert changes for on-prem AD Users
# Written By: Samir Makwana (makwanasamir@hotmail.com)
# Last Modified: 12-April-2020
#######################################################################

#To prompt explorer to choose folder where LDF file/s are/is located
$browse = New-Object System.Windows.Forms.FolderBrowserDialog
$browse.ShowDialog()
$FolderPath = $Browse.SelectedPath
$ADVSCFilePath = "$FolderPath\ADUsersData.csv"
$AttrListFilePath = "$FolderPath\Attributes.csv"
$LDIFFile = $Null

#Import AD and Attribute csv files
$ADCSV = Import-Csv $ADVSCFilePath
$AttrList = Import-Csv $AttrListFilePath

#Defining variables to Null or 0
$DN = $Null
$ADAttrValue = $Null
$ADAttrName = $Null
$LDIFFile = $Null
$ADObj = $Null

#There are some attribute names in attribute file which might not required in LDF File comparison hence excuding them using array
$ADVariableToExclude = @("DistinguishedName","ProxyAddresses1","ObjectGUID")


#Foreach object in AD csv file
Foreach($Obj in $ADCSV)
{

    $DN = $Null
    $DN = $Obj.DistinguishedName
    $DN.Length

    #If Object DistingushedName is not null
    If($DN.length -ne 0)
    {

    #Defining User LDF File path
    $LDIFFile = $Null
    $LDIFFile = $FolderPath + "\" + $DN + ".ldf"

    #Getting AD Object properties to compare with CSV file data so that we can identify what has changed and based on it we can generate LDF file for each user
    $ADObj = $Null
    $ADObj = Get-ADObject "$DN" -Properties *

        #Foreach attributes in attribute list
        Foreach($Attr in $AttrList)
        {
            $ADAttrValue = $Null
            $ADAttrName = $Null

            $ADAttrName = $Attr.ADAttribute
            $ADAttrValue = $Obj.$ADAttrName
            
            #Making sure Attribute name is not in the exclusion arrary (List)
            If($ADVariableToExclude -notcontains $ADAttrName)
            {
                #Making sure attribute name and value is not set to null
                if($ADAttrValue.length -ne 0 -and $ADAttrName.length -ne 0)
                {
                    $ADAttrValue

                    #if attribute is proxyaddresses then generate LDF
                    If($ADAttrName -like "ProxyAddresses")
                    {
                        #Convert ProxyAddresses string to array
                        $ADProxyAddresses = $Obj.$ADAttrName -split ";"
                        #Write-Host "AAD Proxy::" $AADProxyAddresses
                        #Write-Host $ADProxyAddresses -ForegroundColor Green
                                        
                        #Generating LDF File with proxyaddresses. As Proxyaddresses is multivalue attribute.
                        #We need to remove the all values and populate the values again
                        "dn: $DN" | Out-File $LDIFFile -Append
                        "changetype: modify" | Out-File $LDIFFile -Append
                        "delete: ProxyAddresses" | Out-File $LDIFFile -Append
                        "-" | Out-File $LDIFFile -Append
                        "" | Out-File $LDIFFile -Append
                        "dn: $DN" | Out-File $LDIFFile -Append
                        "changetype: modify" | Out-File $LDIFFile -Append
                        "add: ProxyAddresses" | Out-File $LDIFFile -Append
                            
                        Foreach($Proxy in $ADProxyAddresses)
                        {
                            If($Proxy -notlike "$Tenant")
                            {
                                "ProxyAddresses: $Proxy" | Out-File $LDIFFile -Append
                            }
                        }
                        "-" | Out-File $LDIFFile -Append
                        "" | Out-File $LDIFFile -Append
                    }
                    elseIf($ADAttrName -like "ProxyAddresses1")
                    {
                    }
                    else
                    {
                        # If CSV attribute value is not equal to AD Object Attribute value and AD Object Attribute value is not null then update LDF File
                        # this is to avoid scenarios where attribute value is NULL in CSV and in AD. If we do not put condition below then it will generate LDF files.
                        If($Obj.$ADAttrName -notlike $ADObj.$ADAttrName -and $ADObj.$ADAttrName -ne $Null)
                        {
                        $Value = $Obj.$ADAttrName

                        "dn: $DN" | Out-File $LDIFFile -Append
                        "changetype: modify" | Out-File $LDIFFile -Append
                        "replace: $ADAttrName" | Out-File $LDIFFile -Append
                        "$ADAttrName : $Value" | Out-File $LDIFFile -Append
                        "-" | Out-File $LDIFFile -Append
                        "" | Out-File $LDIFFile -Append
                        }
                    }

                }
                #If CSV file attribute name is not null
                elseif($ADAttrName.length -ne 0)
                {       
                        # If CSV attribute value is not equal to AD Object Attribute value and AD Object Attribute value is not null then update LDF File
                        # this is to avoid scenarios where attribute value is NULL in CSV and in AD. If we do not put condition below then it will generate LDF files.
                        If($Obj.$ADAttrName -notlike $ADObj.$ADAttrName -and $ADObj.$ADAttrName -ne $Null)
                        {                      
                        "dn: $DN" | Out-File $LDIFFile -Append
                        "changetype: modify" | Out-File $LDIFFile -Append
                        "delete: $ADAttrName" | Out-File $LDIFFile -Append
                        "-" | Out-File $LDIFFile -Append
                        "" | Out-File $LDIFFile -Append
                        }
                }
            }
        }
    }

}

