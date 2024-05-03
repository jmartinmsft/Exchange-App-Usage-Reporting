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
    [Parameter(Mandatory=$false, HelpMessage="Number of message the script should send.")] [string] $UserAgent='DemoAppWithNoScope'
)

function Enable-TraceHandler(){
$sourceCode = @"
    public class ewsTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
    {
        public System.String LogFile {get;set;}
        public void Trace(System.String traceType, System.String traceMessage)
        {
            System.IO.File.AppendAllText(this.LogFile, traceMessage);
        }
    }
"@    

    Add-Type -TypeDefinition $sourceCode -Language CSharp -ReferencedAssemblies $ewsDLL
    $TraceListener = New-Object ewsTraceListener
   return $TraceListener
}

function Get-OAuthToken{
    #Change the AppId, AppSecret, and TenantId to match your registered application

    $OAuthClientSecret = $OAuthClientSecret | ConvertTo-SecureString -Force -AsPlainText
    $OAuthClientCertificate = Get-Item Cert:\CurrentUser\My\6389EA02A19D671CAF8AFA03CA428FC7BB9AC16D
    $OAuthRedirectUri =  "https://login.microsoftonline.com/common/oauth2/nativeclient"
    $Uri = "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token"
    $Scope = "https://outlook.office365.com/.default"

    <#
    #Build the URI for the token request
        $Body = @{
        client_id     = $OAuthClientId
        scope         = $Scope
        client_secret = $OAuthClientSecret
        grant_type    = "client_credentials"
    }
    $TokenRequest = Invoke-WebRequest -Method Post -Uri $Uri -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
    #Unpack the access token
    $Token = ($TokenRequest.Content | ConvertFrom-Json).Access_Token
    #>

    if((Get-Date).Hour -ge 12) {
        Write-Host "Using client secret to request OAUth token" -ForegroundColor Yellow
        $MsalParams = @{
            ClientId = $OAuthClientId
            TenantId = $OAuthTenantId
            ClientSecret = $OAuthClientSecret
            RedirectUri = $OAuthRedirectUri
            Scopes =  $Scope
        }    
    }
    else {
        Write-Host "Using client certificate to request OAUth token" -ForegroundColor Yellow
        $MsalParams = @{
            ClientCertificate = $OAuthClientCertificate
            ClientId = $OAuthClientId
            TenantId = $OAuthTenantId
            RedirectUri = $OAuthRedirectUri
            Scope = $Scope
        }    
    }
    $Token = (Get-MsalToken @MsalParams).AccessToken
    #Get-MsalToken -ClientCertificate (Get-Item Cert:\CurrentUser\My\6389EA02A19D671CAF8AFA03CA428FC7BB9AC16D) -ClientId a60993cf-6629-4c11-8f41-6a767072d97d -TenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -RedirectUri "https://login.microsoftonline.com/common/oauth2/nativeclient"
    return $Token
}

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
Write-Host -ForegroundColor Yellow '//**********************************************************************​'
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
$Token = (Get-MsalToken -Interactive -TenantId 9101fc97-5be5-4438-a1d7-83e051e52057 -Scopes https://outlook.office.com/.default -RedirectUri https://login.microsoftonline.com/common/oauth2/nativeclient -ClientId $OAuthClientId).AccessToken
$OAuthToken = "Bearer {0}" -f $Token
#endregion

#region EwsService
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.HttpHeaders.Clear()
$service.UserAgent = "EwsPowerShellScript"
$service.HttpHeaders.Add("Authorization", " $($OAuthToken)")
$service.Url = "https://outlook.office365.com/ews/exchange.asmx"
$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
$service.UserAgent = $UserAgent
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
#endregion
#$service.HttpHeaders.Remove("Authorization")

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
        $FolderCheck = $FolderName.Replace(" ","")
    
        if($WellKnownFolderNames -notcontains $FolderCheck) {
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
                #$tfTargetFolder = $null  
                exit  
            }
        }
        #region ConnectToFolder
        else {
            $folderid= New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$FolderCheck,$MailboxName)
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
        #    do{
        $fiResult = $MailboxFolder.FindItems($ivItemView)
        foreach($Item in $fiResult.Items){  
            $Item.Subject
            #$Item.Id
        }

        #if ($HardDelete) {
            $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete
        #} else {
        #    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete
        #}
        $Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)

        
        $ivItemView.offset += $fiResult.Items.Count  
        #    }
        #    while($fiResult.MoreAvailable -eq $true)
        #endregion
