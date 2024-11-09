<#
//***********************************************************************
//
// Copyright (c) 2018 Microsoft Corporation. All rights reserved.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//**********************************************************************​
//
#>
# Version 2024.11.08.1530
param(
    [Parameter(Mandatory=$false, HelpMessage="The MailboxName parameter specifies the mailbox to be accessed.")]
    [string] $MailboxName='thanos@thejimmartin.com',

    [Parameter(Mandatory=$false, HelpMessage="The FolderName parameter specfies the folder to be accessed.")]
    [string] $FolderName='Inbox',

    [ValidateSet("MailItemsAccessed", "MoveToDeletedItems", "SoftDelete", "HardDelete","Update","Send")]
    [Parameter(Mandatory = $false, HelpMessage="The Operation parameter specifies the action to be taken against the item.")]
    [string]$Operation = "MailItemsAccessed",

    [Parameter(Mandatory=$false, HelpMessage="The OAuthClientId parameter specifies the app ID for the OAuth token request.")]
    [string]$OAuthClientId='2f79178b-54c3-4e81-83a0-a7d16010a424',

    [Parameter(Mandatory=$false, HelpMessage="The OAuthTenantId parameter specifies the the tenant ID for the OAuth token request.")]
    [string]$OAuthTenantId,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri specifies the redirect Uri of the Azure registered application.")]
    [string]$OAuthRedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",
    
    [Parameter(Mandatory=$false, HelpMessage="The UserAgent parameter specifies the user agent passed in the request.")]
    [string]$UserAgent='DemoAppWithNoScope',

    [Parameter(Mandatory=$false, HelpMessage="The UseImpersonation switch specifies whether the request should use impersonation.")]
    [switch]$UseImpersonation
)

#region Disclaimer
Write-Host -ForegroundColor Yellow '//***********************************************************************'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '// Copyright (c) 2018 Microsoft Corporation. All rights reserved.'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR'
Write-Host -ForegroundColor Yellow '// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,'
Write-Host -ForegroundColor Yellow '// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE'
Write-Host -ForegroundColor Yellow '// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER'
Write-Host -ForegroundColor Yellow '// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,'
Write-Host -ForegroundColor Yellow '// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN'
Write-Host -ForegroundColor Yellow '// THE SOFTWARE.'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '//***********************************************************************'
#endregion

#region LoadEwsManagedAPI
#Check for EWS Managed API, exit if missing
$ewsDLL = (($(Get-ItemProperty -ErrorAction Ignore -Path Registry::$(Get-ChildItem -ErrorAction Ignore -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services' |Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory'))
if($ewsDLL -notlike $null) {
    $ewsDLL = $ewsDLL+"Microsoft.Exchange.WebServices.dll"
}else {
    $ScriptPath = Get-Location
    $ewsDLL = "$ScriptPath\Microsoft.Exchange.WebServices.dll"
}
if (Test-Path $ewsDLL) {
    Import-Module $ewsDLL
}else {
    Write-Warning "This script requires the EWS Managed API 1.2 or later."
    exit
}
#endregion

#region GetOAuthToken
$Token = (Get-MsalToken -Interactive -TenantId $OAuthTenantId -Scopes https://outlook.office.com/.default -RedirectUri $OAuthRedirectUri -ClientId $OAuthClientId).AccessToken
$OAuthToken = "Bearer {0}" -f $Token
#endregion

#region EwsService
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.HttpHeaders.Clear()
$service.HttpHeaders.Add("Authorization", " $($OAuthToken)")
$service.Url = "https://outlook.office365.com/ews/exchange.asmx"
$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
$service.UserAgent = $UserAgent
if($UseImpersonation) {
    Write-Host "Using impersonation with the signed-in user." -ForegroundColor Cyan
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
}

#endregion

$WellKnownFolderNames = @("ArchiveDeletedItems",
            "ArchiveMsgFolderRoot",
            "ArchiveRecoverableItemsDeletions",
            "ArchiveRecoverableItemsPurges",
            "ArchiveRecoverableItemsRoot",
            "ArchiveRecoverableItemsVersions",
            "ArchiveRoot",
            "Calendar",
            "Conflicts",
            "Contacts",
            "ConversationHistory",
            "DeletedItems",
            "Drafts",
            "Inbox",
            "Journal",
            "JunkEmail",
            "LocalFailures",
            "MsgFolderRoot",
            "Notes",
            "Outbox",
            "PublicFoldersRoot",
            "QuickContacts",
            "RecipientCache",
            "RecoverableItemsDeletions",
            "RecoverableItemsPurges",
            "RecoverableItemsRoot",
            "RecoverableItemsVersions",
            "Root",
            "SearchFolders",
            "SentItems",
            "ServerFailures",
            "SyncIssues",
            "Tasks",
            "ToDoSearch",
            "VoiceMail"
    )

$script:RequiredPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
    [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
    [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId,
    [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ReceivedBy,
[Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)


if($FolderName.Replace(" ","") -notin $WellKnownFolderNames) {
    Write-Host "Searching for $FolderName in the mailbox..." -ForegroundColor Cyan
    $folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
    $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
    $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $SfSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName)
    $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView)
    if ($findFolderResults.TotalCount -gt 0){ 
        foreach($folder in $findFolderResults.Folders){ 
            $folderid = $folder.Id
        } 
    } 
    else{ 
        Write-Warning "$FolderName was not found in the mailbox for $MailboxName"  
        exit  
    }
}
else {
    $FolderName = $FolderName.Replace(" ","")
    $folderid= New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$FolderName,$MailboxName)
}

Write-Host "Connecting to the $FolderName for $MailboxName..." -ForegroundColor Cyan -NoNewline
try { 
    $MailboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid) 
    Write-Host "COMPLETE"
}
catch { 
    Write-Host "FAILED" -ForegroundColor Red 
    exit
}
#endregion

#region GetItems
$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  
$Item = $MailboxFolder.FindItems($ivItemView)
#foreach($Item in $fiResult.Items){  
    #$Item.Subject
    #$ItemBind = [Microsoft.Exchange.WebServices.data.Item]::Bind($service, $item.Id,$script:RequiredPropSet)
#}
if($Operation -ne "Send") {
    Write-Host "Action is being taken against the message with the subject '$($Item.Subject)'."
}
switch($Operation) {
    "MailItemsAccessed" {[Microsoft.Exchange.WebServices.data.Item]::Bind($service, $item.Id,$script:RequiredPropSet) | Out-Null}
    "MoveToDeletedItems" {$Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)}
    "SoftDelete" {$Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)}
    "HardDelete" {$Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)}
    "Update" {
        $Item.IsRead = $true
        $Item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
    }
    "Send" {
        $SentItemsId= New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)
        try {
            $SentItemsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$SentItemsId)
        }
        catch { 
            Write-Warning "Unable to connect to the Sent Items folder for $($MailboxName)"; 
            exit
        }
        $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
        $message.Subject = "Test message sent using EWS"
        $message.Body = "This message was sent using EWS on $(Get-Date)."
        $message.ToRecipients.Add($MailboxName)
        $message.SendAndSaveCopy($SentItemsFolder.Id) | Out-Null
    }
    "Move" {
        $DeletedItemsId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems,$MailboxName)
        try {
            $DeletedItemsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$DeletedItemsId)
        }
        catch { 
            Write-Warning "Unable to connect to the Deleted Items folder for $($MailboxName)"; 
            exit
        }
        $Item.Move($DeletedItemsFolder.Id) | Out-Null
    }
}
#endregion