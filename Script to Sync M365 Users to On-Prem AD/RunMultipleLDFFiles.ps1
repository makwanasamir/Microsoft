
#######################################################################
# Name: RunMultipleLDFFiles.ps1
# This script to run multiple LDF files from same folder
# Written By: Samir Makwana (makwanasamir@hotmail.com)
# Last Modified: 12-April-2020
#######################################################################

#To prompt explorer to choose folder where LDF file/s are/is located
$browse = New-Object System.Windows.Forms.FolderBrowserDialog
$browse.ShowDialog()
$FolderPath = $Browse.SelectedPath

# Get LDF file paths
$LDFFiles = (Get-ChildItem -Path $FolderPath -Filter *.ldf).VersionInfo.Filename
$LDFFiles.COunt

#Prompt to enter name of the domain controller
$DC = Read-Host "Enter FQDN of Domain Controller"

#If connection to domain controller works then continue else go to else section
If(Test-Connection -ComputerName $DC -Quiet)
{ 

    #Foreach LDF Files run import ldf command
    Foreach($LDFFile in $LDFFiles)
    {
        If(Test-Path $LDFFile)
        {
        Write-host "Processing $LDFFile" -ForegroundColor Green

        #LDF file import command
        ldifde -i -f $LDFFile -s $DC

        }
        else
        {
        Write-host "Unable to find LDF file : $LDFFile" -ForegroundColor Red
        }
    }
}
else
{
Write-host "Failed to connect domain controller : $DC" -ForegroundColor Red
}