Install-Module Microsoft.Graph -Scope CurrentUser

#Tenant ID
$tenantId = "ba8f8bff-c7fb-4629-bfb6-d6674106f7a5"

#Function App Managed Identity get it from Azure Portal
$managedIdentityIds = "ffd12c24-c54b-478f-bd28-3a811b7e4b93"#, "ANOTHER-MANAGED-ID-IF-NEEDED"

Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All, Application.Read.All, RoleManagement.ReadWrite.Directory -TenantId $tenantId

# Office 365 Exchange Online (the GUID is the same in all tenants)
$resourceId = (Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'").Id
# Exchange.ManageAsApp (the GUID is the same in all tenants)
$appRoleId = "dc50a0fb-09a3-484d-be87-e023b12c6440"
$roleDefinitionId = (Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq 'Exchange Administrator'").Id

foreach ($managedIdentityId in $managedIdentityIds) {
    # Grant the Exchange.ManageAsApp API permission for the function app
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentityId -PrincipalId $managedIdentityId -AppRoleId $appRoleId -ResourceId $resourceId
    # Grant the Exchange Administrator role for the function app
    New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $managedIdentityId -RoleDefinitionId $roleDefinitionId -DirectoryScopeId "/"
}


# Reference: https://laurakokkarinen.com/how-to-use-exchange-online-powershell-on-azure-functions-with-managed-identity/