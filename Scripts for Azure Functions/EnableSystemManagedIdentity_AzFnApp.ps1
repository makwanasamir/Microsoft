#Install-Module -Name Az.Functions -Scope CurrentUser

Connect-AzAccount -Tenant yourtenant.onmicrosoft.com -Credential

# Import the Az.Functions module if not already imported
Import-Module Az.Functions

# Variables
$resourceGroupName = "psglobalsummit2025"  # Replace with your resource group name
$functionAppName = "psglobalsummit2025"      # Replace with your function app name

# Enable system-assigned managed identity
#Set-AzFunctionApp -ResourceGroupName $resourceGroupName -Name $functionAppName -IdentityType SystemAssigned

# Verify the managed identity is enabled
$functionApp = Get-AzFunctionApp -ResourceGroupName $resourceGroupName -Name $functionAppName
$functionApp.Identity
