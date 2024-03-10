Get-EwsImpersonation

This PowerShell script can be used to locate app registrations that are using the specified API and generate a sign-in report for those app registrations.

Requirements
An application registration must be created in Entra for the tenant and this application must have Application.Read.All and AuditLog.Read.All Graph API permission (either Application or Delegated). The script also requires the MSAL.PS PowerShell module.

How To Run
This syntax will get sign-in logs for app registrations with EWS permissions using delegated permissions (prompt for credentials).

.\Get-EwsImpersonation.ps1 -Api EWS -PermissionType Delegate -OAuthClientId abcdefg-1234-hijklm -OAuthClientSecret Pl3a$eD0n'tSh@rE -OAuthTenantId 91000000-11111-1234-3000000 -OAuthRedirectUri https://login.microsoftonline.com/common/oauth2/nativeclient -OutputPath c:\temp

This syntax will get sign-in logs for app registrations with EWS permissions using application permissions (using a certificate).

.\Get-EwsImpersonation.ps1 -Api EWS -PermissionType Application -OAuthClientId abcdefg-1234-hijklm -OAuthCertificate 654321ABCDEFG023 -OAuthTenantId 91000000-11111-1234-3000000 -OAuthRedirectUri https://login.microsoftonline.com/common/oauth2/nativeclient -OutputPath c:\temp


Parameters

Api
The Api parameter specifies which API Permisions to export for esach Application registration.

PermissionType
The PermissionType parameter specifies whether the app registrations uses delegated or application permissions (Default value is Application).

OAuthClientId
The OAuthClientId parameter specifies the the app ID for the OAuth token request.

OAuthClientSecret
The OAuthClientSecret parameter specifies the the app secret for the OAuth token request.

OAuthTenantId
The OAuthTenantId parameter specifies the the tenant ID for the OAuth token request.

OAuthRedirectUri
The OAuthRedirectUri specifies the redirect Uri of the Azure registered application.

OAuthCertificate
The OAuthCertificate parameter is the certificate for the registerd application.

CertificateStore
The CertificateStore parameter specifies the certificate store where the certificate is loaded.$null,

OutputPath
The OutputPath parameter specifies the path for the EWS usage report.

NumberOfDays
The NumberOfDays parameter specifies how many days of sign-in logs to query (default is one).
