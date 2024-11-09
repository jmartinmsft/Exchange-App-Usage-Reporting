**Graph-FindImpersonation**

Report of users leveraging the ApplicationImpersonation RBAC role with third party EWS applications.

# Description:
This script must be run a minimum of three times to get the report. Each run requires a different value for the Operation parameter. Here's a high-level overview of the the steps and what they do:

1. NewAuditQuery - Uses the Graph API to create a new audit log query in the tenant for Exchange mailbox events.
2. CheckAuditQuery - Checks the status of the audit log query to determine when it has succeeeded. The audit query ID from the NewAuditQuery should be provided.
3. GetQueryResults - Uses the Graph API to retrieve the audit records from the query and filter for the ApplicationImpersonation events. The results are outputed into a CSV file to the path specified in the command.

# Requirements:
The script requires an application registration in Entra ID that has the Graph API AuditLogsQuery.Read.All permission. The permission may be Application or Delegated type.

# Usage:
Step 1: Create the new audit log query:
```powershell
.\Graph-FindImpersonation.ps1 -PermissionType Application -OAuthClientId f733c1fb-e6d7-4d65-b542-33b5e4a604ca -OutputPath C:\Temp\Output\ -Scope AuditLog.Read.All -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -Operation NewAuditQuery -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date)
```
Step 2: Check the status of the audit log query until it shows succeeded:
```powershell
.\Graph-FindImpersonation.ps1 -PermissionType Application -OAuthClientId f733c1fb-e6d7-4d65-b542-33b5e4a604ca -OutputPath C:\Temp\Output\ -Scope AuditLog.Read.All -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be -Operation CheckAuditQuery
```
Step 3: Retrieve the list of records from the audit log query:
```powershell
.\Graph-FindImpersonation.ps1 -PermissionType Application -OAuthClientId f733c1fb-e6d7-4d65-b542-33b5e4a604ca -OutputPath C:\Temp\Output\ -Scope AuditLog.Read.All -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be -Operation GetQueryResults
```

# Parameters:

OutputPath - The OutputPath parameter specifies the path for the output files.

AzureEnvironment - The AzureEnvironment parameter specifies the environment for the tenant (default is Global).

PermissionType- The PermissionType parameter specifies whether the app registrations uses delegated or application permissions

OAuthClientId - The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.

OAuthTenantId - The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as audit logs being accessed).

OAuthRedirectUri - The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.

OAuthClientSecret - The OAuthSecretKey parameter is the the secret for the registered application.

OAuthCertificate - The OAuthCertificate parameter is the certificate for the registered application.

CertificateStore - The CertificateStore parameter specifies the certificate store where the certificate is loaded.

Scope - The Scope parameter specifies the scope for the OAuth token request.

AuditQueryId - The AuditQueryId parameter specifies the id for the audit query.

Operation - The Operation parameter specifies what action you want the script to perform.
    NewAuditQuery - creates a new audit query for Exchange events
    CheckAuditQuery - checks the status of the audit query using the audit id specified
    GetQueryResults - retrieves the audit records from the audit query and filters for EWS impersonation events

StartDate - The StartDate parameter specifies the start date for the audit query.

EndDate - The EndDate parameter specifies the end date for the audit query.
