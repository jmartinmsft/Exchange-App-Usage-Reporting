Get-AppRegistrationsByApiPermission


The Get AppRegistrations by API permisison script helps to find applications that are using EWS API permissions. It uses the Graph API to query Entra ID for all app registrations and service principals within the tenant. It then correlates the results to provide a list of applications with both Application and Delegated EWS API permissions.

Requirements

The script requires an application registration in Entra ID that has the Microsoft.Graph Application.Read.All permission. The permission may be Application or Delegated type.

Syntax

This cmdlet will run the Get AppRegistrations by API permission script using an application secret.
.\Get-AppRegistrationsByApiPermission.ps1 -Api EWS -OAuthClientSecret $secret -PermissionType Application -OAuthClientId f733c1fb-e6d7-5e76-b542-33b5e4a604ca -OAuthTenantId 9101fc97-6cf6-4438-a1d7-83e051e52057 -OutputPath C:\Scripts\Results\

This cmdlet will run the Find Impersonation User script using delegated permissions.

.\Get-AppRegistrationsByApiPermission.ps1 -Api EWS -PermissionType Delegated -OAuthClientId f733c1fb-e6d7-5e76-b542-33b5e4a604ca -OutputPath C:\Scripts\Results\

This cmdlet will run the Find Impersonation User script using a certificate.

.\Get-AppRegistrationsByApiPermission.ps1 -Api EWS -PermissionType Application -OAuthClientId f733c1fb-e6d7-4d65-b542-33b5e4a604ca -OAuthTenantId 9101fc97-6cf6-4438-a1d7-83e051e52057  -OAuthCertificate 6389EA02A19D671CAF8AFA03CA428FC7BB9AC16D -CertificateStore LocalMachine -OutputPath C:\Scripts\Results\


