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
//
// This script attempts to send messages from a random EXO mailbox to a random group of EXO recipients using EWS.
// This script uses application permissions and is restricted by an RBAC management role assignment
#>

param(
    [Parameter(Mandatory=$false, HelpMessage="Number of message the script should send.")] [string] $MailboxName='thanos@thejimmartin.com',
    [Parameter(Mandatory=$false, HelpMessage="Number of message the script should send.")] [string] $FolderName='Inbox',
    [Parameter(Mandatory=$false, HelpMessage="Number of message the script should send.")] [string] $OAuthClientId='2f79178b-54c3-4e81-83a0-a7d16010a424',
    [Parameter(Mandatory=$false, HelpMessage="The OAuthTenantId parameter specifies the the tenant ID for the OAuth token request.")][string] $OAuthTenantId,
    [Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri specifies the redirect Uri of the Azure registered application.")][string] $OAuthRedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",
    [Parameter(Mandatory=$false, HelpMessage="Number of message the script should send.")] [string] $UserAgent='DemoAppWithNoScope'
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
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
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
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(15)  
        $fiResult = $MailboxFolder.FindItems($ivItemView)
        foreach($Item in $fiResult.Items){  
            $Item.Subject
        }
        #$Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
        #endregion