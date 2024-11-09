# Demo-EWS-Traffic

Script that simulates EWS traffic in Exchange Online either as a user or with impersonation rights.

## Description
The high-level overview of what this script does is the following:

1. Connects to the specified folder in the mailbox provided in the command.
2. For all operations other than Send, the operation is performed against the first item in the folder. For example, MailItemAccessed operation would result in the first item in the folder being accessed by the script.
3. The Send operation will send a message to the mailbox from itself with the message being saved in the Sent Items folder.

## Requirements
1. The script requires an application registration in Entra ID that has the Office 365 Exchange Online EWS.AccessAsUser.All delegated permission.
2. The MSAL.PS module to obtain a delegated auth token.
3. The Microsoft.Exchange.WebServices.dll file on the system running the script. Recommended having in the same folder as the script. This can be copied from an Exchange server.

## Usage
Move an item from the Inbox to the Deleted Items folder using impersonation:
```powershell
.\Demo-EWS-Traffic.ps1 -MailboxName jim@contoso.com -OAuthClientId <YourAppId> -OAuthTenantId <YourTenantId> -UserAgent DemoEwsApp -FolderName Inbox -Operation MoveToDeletedItems -UseImpersonation
```
Send a message using impersonation
```powershell
.\Demo-EWS-Traffic.ps1 -MailboxName jim@contoso.com -OAuthClientId <YourAppId> -OAuthTenantId <YourTenantId> -UserAgent DemoEwsApp -Operation Send -UseImpersonation
```
Permanently delete a message from the Inbox as the mailbox owner
```powershell
.\Demo-EWS-Traffic.ps1 -MailboxName jim@contoso.com -OAuthClientId <YourAppId> -OAuthTenantId <YourTenantId> -UserAgent DemoEwsApp -FolderName Inbox -Operation HardDelete
```

## Parameters
**OutputPath** - The OutputPath parameter specifies the path for the output files.
**MailboxName** - The MailboxName parameter specifies the mailbox to be accessed.
**FolderName** - The FolderName parameter specfies the folder to be accessed.
**Operation** - The Operation parameter specifies the action to be taken against the item. Valid values for this parameter include: MailItemsAccessed, MoveToDeletedItems, SoftDelete, HardDelete, Update, Send, Move
**OAuthClientId** - The OAuthClientId parameter specifies the app ID for the OAuth token request.
**OAuthTenantId** - The OAuthTenantId parameter specifies the the tenant ID for the OAuth token request.
**OAuthRedirectUri** - The OAuthRedirectUri specifies the redirect Uri of the Azure registered application.
**UserAgent** - The UserAgent parameter specifies the user agent passed in the request.
**UseImpersonation** The UseImpersonation switch specifies whether the request should use impersonation.