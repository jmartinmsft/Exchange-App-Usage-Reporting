# Graph-FindImpersonation

Report of users leveraging the ApplicationImpersonation RBAC role with third party EWS applications.

## Description
This script must be run a minimum of three times to get the report. Each run requires a different value for the Operation parameter. Here's a high-level overview of the the steps and what they do:

1. NewAuditQuery - Uses the Graph API to create a new audit log query in the tenant for Exchange mailbox events.
2. CheckAuditQuery - Checks the status of the audit log query to determine when it has succeeeded. The audit query ID from the NewAuditQuery should be provided.
3. GetQueryResults - Uses the Graph API to retrieve the audit records from the query and filter for the ApplicationImpersonation events. The results are outputed into a CSV file to the path specified in the command.

## Requirements
The script requires an application registration in Entra ID that has the Graph API AuditLogsQuery.Read.All permission. The permission may be Application or Delegated type.

To use delegated permissions a Redirect URI must be configured for Mobile and desktop applications with the value http://localhost:8004.

## Query for a single user
A query can be performed for a single user with the impersonation role using the **AdminSid** parameter. This can be used to verfiy an account is no longer using impersonation. It can also be used to create queries with a smaller set of records to be retrieved.

## Application permission usage
### Using a certificate
Step 1: Create the new audit log query:
```powershell
.\Graph-FindImpersonation.ps1 -Name AllResults -PermissionType Application -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -Operation NewAuditQuery -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date)
```
Step 2: Check the status of the audit log query until it shows succeeded:
```powershell
.\Graph-FindImpersonation.ps1 -PermissionType Application -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -Operation CheckAuditQuery -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be
```
Step 3: Retrieve the list of records from the audit log query:
```powershell
.\Graph-FindImpersonation.ps1 -PermissionType Application -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -Operation GetQueryResults -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be
```

### Using a secret
Step 1: Create the new audit log query:
```powershell
$secret = ConvertTo-SecureString "XXXXXXXXXXXXXXXXXXX" -AsPlainText -Force
.\Graph-FindImpersonation.ps1 -Name AllResults -PermissionType Application -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthClientSecret $secret -Operation NewAuditQuery -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date)
```
Step 2: Check the status of the audit log query until it shows succeeded:
```powershell
.\Graph-FindImpersonation.ps1 -Name AllResults -PermissionType Application -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthClientSecret $secret -Operation CheckAuditQuery -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be
```
Step 3: Retrieve the list of records from the audit log query:
```powershell
.\Graph-FindImpersonation.ps1 -Name AllResults -PermissionType Application -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthClientSecret $secret -Operation GetQueryResults -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be
```

## Delegate permission usage
Step 1: Create the new audit log query using delegated permission:
```powershell
.\Graph-FindImpersonation.ps1 -Name AllResults -PermissionType Delegated -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -Operation NewAuditQuery
```
Step 2: Check the status of the audit log query until it shows succeeded:
```powershell
.\Graph-FindImpersonation.ps1 -Name AllResults -PermissionType Delegated -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -Operation CheckAuditQuery -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be
```
Step 3: Retrieve the list of records from the audit log query:
```powershell
.\Graph-FindImpersonation.ps1 -Name AllResults -PermissionType Delegated -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -Operation GetQueryResults -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be
```

## Query for single user
```powershell
.\Graph-FindImpersonation.ps1 -Name JMartin -PermissionType Application -OAuthClientId 5c4abea3-43e5-4220-a35a-bb344d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthClientSecret $secret -Operation NewAuditQuery -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date) -AdminSid S-1-5-21-3145204594-529760289-3943512046-19014454
```

## Parameters

**Name** - The Name parameter specifies the name for the query and value to be appended to the output file.

**OutputPath** - The OutputPath parameter specifies the path for the output files.

**AzureEnvironment** - The AzureEnvironment parameter specifies the environment for the tenant (default is Global).

**PermissionType** - The PermissionType parameter specifies whether the app registrations uses delegated or application permissions

**OAuthClientId** - The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.

**OAuthTenantId** - The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as audit logs being accessed).

**OAuthRedirectUri** - The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.

**OAuthClientSecret** - The OAuthSecretKey parameter is the the secret for the registered application.

**OAuthCertificate** - The OAuthCertificate parameter is the certificate for the registered application.

**CertificateStore** - The CertificateStore parameter specifies the certificate store where the certificate is loaded.

**Scope** - The Scope parameter specifies the scope for the OAuth token request.

**AuditQueryId** - The AuditQueryId parameter specifies the id for the audit query.

**Operation** - The Operation parameter specifies what action you want the script to perform.
    NewAuditQuery - creates a new audit query for Exchange events
    CheckAuditQuery - checks the status of the audit query using the audit id specified
    GetQueryResults - retrieves the audit records from the audit query and filters for EWS impersonation events

**StartDate** - The StartDate parameter specifies the start date for the audit query.

**EndDate** - The EndDate parameter specifies the end date for the audit query.

**AdminSid** The AdminSid parameter specifies security description (SID) of the user with impersonation right.