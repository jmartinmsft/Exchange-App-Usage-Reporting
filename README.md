Find-ImpersonationUsers

The Find Impersonation Users script helps to find user accounts that are using the Exchange ApplicationImpersation role. It queries the Unified Audit Logs for Exchange events and filters for relevant results. The results can be used to help locate an applications that are leveraging the ApplicationImpersonation role prior to its retirement in Exchange Online.

Requirements
The script requires an application registration in Entra ID that has the Office 365 Management APIs ActiveFeed.Read permission. The permission may be Application or Delegated type.

Syntax

This cmdlet will run the Find Impersonation User script using an application secret.
.\Find-ImpersonationUsers.ps1 -PermissionType Application  -OAuthClientSecret $secret -OAuthClientId f733c1fb-e6d7-5e76-b542-33b5e4a604ca -OAuthTenantId 9101fc97-6cf6-4438-a1d7-83e051e52057 -OutputPath C:\Scripts\Results\

This cmdlet will run the Find Impersonation User script using delegated permissions.

.\Find-ImpersonationUsers.ps1 -OAuthClientSecret $secret -PermissionType Delegated -OAuthClientId f733c1fb-e6d7-5e76-b542-33b5e4a604ca -OutputPath C:\Scripts\Results\ -Scope ActivityFeed.Read

This cmdlet will run the Find Impersonation User script using a certificate.

.\Find-ImpersonationUsers.ps1 -PermissionType Application -OAuthClientId f733c1fb-e6d7-5e76-b542-33b5e4a604ca -OAuthCertificate 6389EA02A19D671CAF9BFA03CA428FC7BB9AC16D -CertificateStore LocalMachine -OutputPath C:\Scripts\Results\

Parameters

OutputPath - The OutputPath parameter specifies the path for the output files.

AzureEnvironment - The AzureEnvironment parameter specifies the environment for the tenant (default is Global).

PermissionType- The PermissionType parameter specifies whether the app registrations uses delegated or application permissions

OAuthClientId - The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.

OAuthTenantId - The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as audit logs being accessed).

OAuthRedirectUri - The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.

OAuthClientSecret - The OAuthSecretKey parameter is the the secret for the registered application.

OAuthCertificate - The OAuthCertificate parameter is the certificate for the registered application. Certificate auth requires MSAL libraries to be available.

CertificateStore - The CertificateStore parameter specifies the certificate store where the certificate is loaded.

Scope - The Scope parameter specifies the scope for the OAuth token request.
