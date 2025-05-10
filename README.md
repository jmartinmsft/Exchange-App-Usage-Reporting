# Find-EwsUsage

Report of applications using EWS to connect to mailboxes.

## Requirements
The script requires an application registration in Entra ID that has the Graph API AuditLogsQuery.Read.All and Application.Read.All permission. The permission may be Application or Delegated type.

To use delegated permissions a Redirect URI must be configured for Mobile and desktop applications with the value http://localhost:8004.

## EWS usage report
This script provides two ways to collect EWS usage information:
1. Sign-in log query
2. Audit log query

Not all mailboxes have auditing enabled to capture EWS activity, but these logs can provide more insight into the level of activity for an application. The sign-in logs will not provide good activity usage, but they will show all applications registered within the tenant showing EWS activity.


## Sign-in logs
This script requires two steps to get the usage activity for an application.

### EWS sign-in activity report
This script generates a report of all applications registered in the tenant with EWS API permissions. It then checks the sign-in activity for those applications and provides the date timestamp of the last sign-in event.
```powershell
.\Find-EwsUsage.ps1 -PermissionType Application -OAuthClientId 1d5cdea4-32e6-1234-a35a-cc443d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 67DCA626D48EE1626623FF26E6C8D856262D1DDC -CertificateStore CurrentUser -Operation GetEwsActivity
```
Then once this list of applications is discovered, a query for sign-in activity can be run:
```powershell
.\Find-EwsUsage.ps1 -PermissionType Application -OAuthClientId 1d5cdea4-32e6-1234-a35a-cc443d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 67DCA626D48EE1626623FF26E6C8D856262D1DDC -CertificateStore CurrentUser -Operation GetAppUsage -QueryType signInLogs -Name ExchangeServerApp -AppId 61b01baf-f6f5-40ec-946b-e3491855fca8 -StartDate (Get-Date).AddDays(-2) -EndDate (Get-Date) -Interval 4
```

## Audit logs
When using the audit logs, this script must be run a minimum of three times to get the report. Each run requires a different value for the AuditQueryStep parameter. Here's a high-level overview of the the steps and what they do:

1. NewAuditQuery - Uses the Graph API to create a new audit log query in the tenant for Exchange mailbox events.
2. CheckAuditQuery - Checks the status of the audit log query to determine when it has succeeeded. The audit query ID from the NewAuditQuery should be provided.
3. GetQueryResults - Uses the Graph API to retrieve the audit records from the query The results are outputed into a CSV file to the path specified in the command and a summary report displayed.


## Application permission usage
### Using a certificate
Step 1: Create the new audit log query:
```powershell
.\Find-EwsUsage.ps1 -PermissionType Application -OAuthClientId 1d5cdea4-32e6-1234-a35a-cc443d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 67DCA626D48EE1626623FF26E6C8D856262D1DDC -CertificateStore CurrentUser -Operation GetAppUsage -QueryType AuditLogs -Name DemoForReadMe -AuditQueryStep NewAuditQuery -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date)
```
Step 2: Check the status of the audit log query until it shows succeeded:
```powershell 
.\Find-EwsUsage.ps1 -PermissionType Application -OAuthClientId 1d5cdea4-32e6-1234-a35a-cc443d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 67DCA626D48EE1626623FF26E6C8D856262D1DDC -CertificateStore CurrentUser -Operation GetAppUsage -QueryType AuditLogs -Name DemoForReadMe -AuditQueryStep CheckAuditQuery -AuditQueryId 8536b790-f0b2-4e00-8a74-1118307e5d65
```
Step 3: Retrieve the list of records from the audit log query:
```powershell
.\Find-EwsUsage.ps1 -PermissionType Application -OAuthClientId 1d5cdea4-32e6-1234-a35a-cc443d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -OAuthCertificate 67DCA626D48EE1626623FF26E6C8D856262D1DDC -CertificateStore CurrentUser -Operation GetAppUsage -QueryType AuditLogs -Name DemoForReadMe -AuditQueryStep GetQueryResults -AuditQueryId 8536b790-f0b2-4e00-8a74-1118307e5d65
```

## Delegate permission usage
Step 1: Create the new audit log query using delegated permission:
```powershell
.\Find-EwsUsage.ps1 -PermissionType Delegated -OAuthClientId 1d5cdea4-32e6-1234-a35a-cc443d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -Operation GetAppUsage -QueryType AuditLogs -Name DemoDelegate -AuditQueryStep NewAuditQuery -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date)
```
Step 2: Check the status of the audit log query until it shows succeeded:
```powershell
.\Find-EwsUsage.ps1 -PermissionType Delegated -OAuthClientId 1d5cdea4-32e6-1234-a35a-cc443d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -Operation GetAppUsage -QueryType AuditLogs -Name DemoDelegate -AuditQueryStep CheckAuditQuery -AuditQueryId 6d7ec4a6-83ee-4ca0-81f3-90fbb2391ea2
```
Step 3: Retrieve the list of records from the audit log query:
```powershell
.\Find-EwsUsage.ps1 -PermissionType Delegated -OAuthClientId 1d5cdea4-32e6-1234-a35a-cc443d697cab -OutputPath C:\Temp\Output\ -OAuthTenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -Operation GetAppUsage -QueryType AuditLogs -Name DemoDelegate -AuditQueryStep GetQueryResults -AuditQueryId 6d7ec4a6-83ee-4ca0-81f3-90fbb2391ea2
```

## Parameters

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

**AppId** - The AppId parameter specifies the application ID used for the usage report.

**Operation** - The Operation parameter specifies the operation the script should perform.

**AuditQueryStep** - The AuditQueryStep parameter specifies the step in the AuditQuery the script should perform.
    NewAuditQuery - creates a new audit query for Exchange events
    CheckAuditQuery - checks the status of the audit query using the audit id specified
    GetQueryResults - retrieves the audit records from the audit query and filters for EWS impersonation events

**QueryType** -The QueryType parameter specifies the type of query for EWS usage.

**Name** - The Name parameter specifies the name for the query and value to be appended to the output file.

**StartDate** - The StartDate parameter specifies the start date for the audit query.

**EndDate** - The EndDate parameter specifies the end date for the audit query.

**Interval** - The Interval parameter specifies the number of hours for sign-in log query.