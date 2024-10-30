Graph-FindImpersonation

The Find Impersonation script helps to find user accounts that are using the Exchange ApplicationImpersation role. It queries the Unified Audit Logs for Exchange events and filters for relevant results. The results can be used to help locate an applications that are leveraging the ApplicationImpersonation role prior to its retirement in Exchange Online.

Requirements
The script requires an application registration in Entra ID that has the Graph API AuditLog.Read.All permission. The permission may be Application or Delegated type.

Syntax

This cmdlet will run the Find Impersonation script to create a new audit query.
.\Graph-FindImpersonation.ps1 -PermissionType Application -OAuthClientId f733c1fb-e6d7-4d65-b542-33b5e4a604ca -OutputPath C:\Temp\Output\ -Scope AuditLog.Read.All -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -Operation NewAuditQuery -NumberOfDays 14

This cmdlet will run the Find Impersonation script to check the audit query status. It may take several hours for a query to complete.
.\Graph-FindImpersonation.ps1 -PermissionType Application -OAuthClientId f733c1fb-e6d7-4d65-b542-33b5e4a604ca -OutputPath C:\Temp\Output\ -Scope AuditLog.Read.All -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be -Operation CheckAuditQuery

This cmdlet will run the Find Impersonation script to retrieve the audit records and filter for EWS impersonation events.
.\Graph-FindImpersonation.ps1 -PermissionType Application -OAuthClientId f733c1fb-e6d7-4d65-b542-33b5e4a604ca -OutputPath C:\Temp\Output\ -Scope AuditLog.Read.All -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 24DCA626D48EE1383623FF26E6C8D852442D1DDC -CertificateStore CurrentUser -AuditQueryId ddc85df1-d5d1-4989-8d25-d7ba3c0bd2be -Operation GetQueryResults
Parameters

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

Operation - The Operation specifies what action you want the script to perform.
    NewAuditQuery - creates a new audit query for Exchange events
    CheckAuditQuery - checks the status of the audit query using the audit id specified
    GetQueryResults - retrieves the audit records from the audit query and filters for EWS impersonation events

NumberOfDays - The NumberOfDays parameter specifies how many days in the past to query the audit logs.
