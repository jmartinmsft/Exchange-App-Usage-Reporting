<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 24.03.11.0955

param(
    [Parameter(Mandatory=$true, Position=1, HelpMessage="The Api parameter specifies which API Permisions to export for esach Application registration")]
    [ValidateSet('OutlookRESTv2','EWS')]
    [string]$Api,

    [Parameter(Mandatory=$false, HelpMessage="The PermissionType parameter specifies whether the app registrations uses delegated or application permissions")]
    [ValidateSet('Application','Delegated')]
    [string]$PermissionType,

    [Parameter(Mandatory=$false, HelpMessage="The OAuthClientId parameter specifies the the app ID for the OAuth token request.")]
    [string] $OAuthClientId="1950a258-227b-4e31-a9cf-717495945fc2",

    [Parameter(Mandatory=$false, HelpMessage="The OAuthClientSecret parameter specifies the the app secret for the OAuth token request.")]
    [securestring] $OAuthClientSecret,

    [Parameter(Mandatory=$false, HelpMessage="The OAuthTenantId parameter specifies the the tenant ID for the OAuth token request.")]
    [string] $OAuthTenantId,

    [Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri specifies the redirect Uri of the Azure registered application.")]
    [string] $OAuthRedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",

    [Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate parameter is the certificate for the registerd application.")]
    [string] $OAuthCertificate = $null,

    [Parameter(Mandatory=$False,HelpMessage="The CertificateStore parameter specifies the certificate store where the certificate is loaded.")]
    [ValidateSet("CurrentUser", "LocalMachine")]
    [string] $CertificateStore = $null,

    [ValidateScript({ Test-Path $_ })]
    [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")]
    [string] $OutputPath,

    [Parameter(Mandatory=$false, HelpMessage="The NumberOfDays parameter specifies how many days of sign-in logs to query (default is three).")]
    [int] $NumberOfDays=1,

    [ValidateSet("Global", "USGovernmentL4", "USGovernmentL5", "ChinaCloud")]
    [Parameter(Mandatory = $false)]
    [string]$AzureEnvironment = "Global",

    [Parameter(Mandatory=$False,HelpMessage="The ImpersonationCheck parameter is a switch that enables checking accounts with EWS impersonation rights.")][switch]$ImpersonationCheck
)

function Write-VerboseLog ($Message) {
    $Script:Logger = $Script:Logger | Write-LoggerInstance $Message
}

function Write-HostLog ($Message) {
    $Script:Logger = $Script:Logger | Write-LoggerInstance $Message
}

function Write-Host {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Proper handling of write host with colors')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [object]$Object,
        [switch]$NoNewLine,
        [string]$ForegroundColor
    )
    process {
        $consoleHost = $host.Name -eq "ConsoleHost"

        if ($null -ne $Script:WriteHostManipulateObjectAction) {
            $Object = & $Script:WriteHostManipulateObjectAction $Object
        }

        $params = @{
            Object    = $Object
            NoNewLine = $NoNewLine
        }

        if ([string]::IsNullOrEmpty($ForegroundColor)) {
            if ($null -ne $host.UI.RawUI.ForegroundColor -and
                $consoleHost) {
                $params.Add("ForegroundColor", $host.UI.RawUI.ForegroundColor)
            }
        } elseif ($ForegroundColor -eq "Yellow" -and
            $consoleHost -and
            $null -ne $host.PrivateData.WarningForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.WarningForegroundColor)
        } elseif ($ForegroundColor -eq "Red" -and
            $consoleHost -and
            $null -ne $host.PrivateData.ErrorForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.ErrorForegroundColor)
        } else {
            $params.Add("ForegroundColor", $ForegroundColor)
        }

        Microsoft.PowerShell.Utility\Write-Host @params

        if ($null -ne $Script:WriteHostDebugAction -and
            $null -ne $Object) {
            &$Script:WriteHostDebugAction $Object
        }
    }
}

function SetProperForegroundColor {
    $Script:OriginalConsoleForegroundColor = $host.UI.RawUI.ForegroundColor

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.WarningForegroundColor) {
        Write-Verbose "Foreground Color matches warning's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.ErrorForegroundColor) {
        Write-Verbose "Foreground Color matches error's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }
}

function RevertProperForegroundColor {
    $Host.UI.RawUI.ForegroundColor = $Script:OriginalConsoleForegroundColor
}

function SetWriteHostAction ($DebugAction) {
    $Script:WriteHostDebugAction = $DebugAction
}

function SetWriteHostManipulateObjectAction ($ManipulateObject) {
    $Script:WriteHostManipulateObjectAction = $ManipulateObject
}

function Write-Verbose {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Verbose from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )

    process {

        if ($null -ne $Script:WriteVerboseManipulateMessageAction) {
            $Message = & $Script:WriteVerboseManipulateMessageAction $Message
        }

        Microsoft.PowerShell.Utility\Write-Verbose $Message

        if ($null -ne $Script:WriteVerboseDebugAction) {
            & $Script:WriteVerboseDebugAction $Message
        }

        # $PSSenderInfo is set when in a remote context
        if ($PSSenderInfo -and
            $null -ne $Script:WriteRemoteVerboseDebugAction) {
            & $Script:WriteRemoteVerboseDebugAction $Message
        }
    }
}

function SetWriteVerboseAction ($DebugAction) {
    $Script:WriteVerboseDebugAction = $DebugAction
}

function SetWriteRemoteVerboseAction ($DebugAction) {
    $Script:WriteRemoteVerboseDebugAction = $DebugAction
}

function SetWriteVerboseManipulateMessageAction ($DebugAction) {
    $Script:WriteVerboseManipulateMessageAction = $DebugAction
}

function Write-Warning {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Warning from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )
    process {

        if ($null -ne $Script:WriteWarningManipulateMessageAction) {
            $Message = & $Script:WriteWarningManipulateMessageAction $Message
        }

        Microsoft.PowerShell.Utility\Write-Warning $Message

        # Add WARNING to beginning of the message by default.
        $Message = "WARNING: $Message"

        if ($null -ne $Script:WriteWarningDebugAction) {
            & $Script:WriteWarningDebugAction $Message
        }

        # $PSSenderInfo is set when in a remote context
        if ($PSSenderInfo -and
            $null -ne $Script:WriteRemoteWarningDebugAction) {
            & $Script:WriteRemoteWarningDebugAction $Message
        }
    }
}

function SetWriteWarningAction ($DebugAction) {
    $Script:WriteWarningDebugAction = $DebugAction
}

function SetWriteRemoteWarningAction ($DebugAction) {
    $Script:WriteRemoteWarningDebugAction = $DebugAction
}

function SetWriteWarningManipulateMessageAction ($DebugAction) {
    $Script:WriteWarningManipulateMessageAction = $DebugAction
}

function Get-NewLoggerInstance {
    [CmdletBinding()]
    param(
        [string]$LogDirectory = (Get-Location).Path,

        [ValidateNotNullOrEmpty()]
        [string]$LogName = "Script_Logging",

        [bool]$AppendDateTime = $true,

        [bool]$AppendDateTimeToFileName = $true,

        [int]$MaxFileSizeMB = 10,

        [int]$CheckSizeIntervalMinutes = 10,

        [int]$NumberOfLogsToKeep = 10
    )

    $fileName = if ($AppendDateTimeToFileName) { "{0}_{1}.txt" -f $LogName, ((Get-Date).ToString('yyyyMMddHHmmss')) } else { "$LogName.txt" }
    $fullFilePath = [System.IO.Path]::Combine($LogDirectory, $fileName)

    if (-not (Test-Path $LogDirectory)) {
        try {
            New-Item -ItemType Directory -Path $LogDirectory -ErrorAction Stop | Out-Null
        } catch {
            throw "Failed to create Log Directory: $LogDirectory. Inner Exception: $_"
        }
    }

    return [PSCustomObject]@{
        FullPath                 = $fullFilePath
        AppendDateTime           = $AppendDateTime
        MaxFileSizeMB            = $MaxFileSizeMB
        CheckSizeIntervalMinutes = $CheckSizeIntervalMinutes
        NumberOfLogsToKeep       = $NumberOfLogsToKeep
        BaseInstanceFileName     = $fileName.Replace(".txt", "")
        Instance                 = 1
        NextFileCheckTime        = ((Get-Date).AddMinutes($CheckSizeIntervalMinutes))
        PreventLogCleanup        = $false
        LoggerDisabled           = $false
    } | Write-LoggerInstance -Object "Starting Logger Instance $(Get-Date)"
}

function Write-LoggerInstance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]$Object
    )
    process {
        if ($LoggerInstance.LoggerDisabled) { return }

        if ($LoggerInstance.AppendDateTime -and
            $Object.GetType().Name -eq "string") {
            $Object = "[$([System.DateTime]::Now)] : $Object"
        }

        # Doing WhatIf:$false to support -WhatIf in main scripts but still log the information
        $Object | Out-File $LoggerInstance.FullPath -Append -WhatIf:$false

        #Upkeep of the logger information
        if ($LoggerInstance.NextFileCheckTime -gt [System.DateTime]::Now) {
            return
        }

        #Set next update time to avoid issues so we can log things
        $LoggerInstance.NextFileCheckTime = ([System.DateTime]::Now).AddMinutes($LoggerInstance.CheckSizeIntervalMinutes)
        $item = Get-ChildItem $LoggerInstance.FullPath

        if (($item.Length / 1MB) -gt $LoggerInstance.MaxFileSizeMB) {
            $LoggerInstance | Write-LoggerInstance -Object "Max file size reached rolling over" | Out-Null
            $directory = [System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)
            $fileName = "$($LoggerInstance.BaseInstanceFileName)-$($LoggerInstance.Instance).txt"
            $LoggerInstance.Instance++
            $LoggerInstance.FullPath = [System.IO.Path]::Combine($directory, $fileName)

            $items = Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*"

            if ($items.Count -gt $LoggerInstance.NumberOfLogsToKeep) {
                $item = $items | Sort-Object LastWriteTime | Select-Object -First 1
                $LoggerInstance | Write-LoggerInstance "Removing Log File $($item.FullName)" | Out-Null
                $item | Remove-Item -Force
            }
        }
    }
    end {
        return $LoggerInstance
    }
}

function Invoke-LoggerInstanceCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance
    )
    process {
        if ($LoggerInstance.LoggerDisabled -or
            $LoggerInstance.PreventLogCleanup) {
            return
        }

        Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*" |
            Remove-Item -Force
    }
}

function Invoke-CatchActionError {
    [CmdletBinding()]
    param(
        [ScriptBlock]$CatchActionFunction
    )

    if ($null -ne $CatchActionFunction) {
        & $CatchActionFunction
    }
}

function Get-NewOAuthToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantID,

        [Parameter(Mandatory = $true)]
        [string]$ClientID,

        [Parameter(Mandatory = $true)]
        [string]$Secret,

        [Parameter(Mandatory = $true)]
        [string]$Endpoint,

        [Parameter(Mandatory = $false)]
        [string]$TokenService = "oauth2/v2.0/token",

        [Parameter(Mandatory = $false)]
        [switch]$CertificateBasedAuthentication,

        [Parameter(Mandatory = $true)]
        [string]$Scope
    )

    <#
        Shared function to create an OAuth token by using a JWT or secret.
        If you want to use a certificate, set the CertificateBasedAuthentication switch and pass a JWT token as the Secret parameter.
        You can use the Get-NewJsonWebToken function to create a JWT token.
        If you want to use a secret, pass the secret as the Secret parameter.
        This function returns a PSCustomObject with the OAuth token, status and the time the token was created.
        If the request fails, the PSCustomObject will contain the exception message.
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
        $oAuthTokenCallSuccess = $false
        $exceptionMessage = $null

        Write-Verbose "TenantID: $TenantID - ClientID: $ClientID - Endpoint: $Endpoint - TokenService: $TokenService - Scope: $Scope"
        $body = @{
            scope      = $Scope
            client_id  = $ClientID
            grant_type = "client_credentials"
        }

        if ($CertificateBasedAuthentication) {
            Write-Verbose "Function was called with CertificateBasedAuthentication switch"
            $body.Add("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer")
            $body.Add("client_assertion", $Secret)
        } else {
            Write-Verbose "Authentication is based on a secret"
            $body.Add("client_secret", $Secret)
        }

        $invokeRestMethodParams = @{
            ContentType = "application/x-www-form-urlencoded"
            Method      = "POST"
            Body        = $body # Create string by joining bodyList with '&'
            Uri         = "$Endpoint/$TenantID/$TokenService"
        }
    }
    process {
        try {
            Write-Verbose "Now calling the Invoke-RestMethod cmdlet to create an OAuth token"
            $oAuthToken = Invoke-RestMethod @invokeRestMethodParams
            Write-Verbose "Invoke-RestMethod call was successful"
            $oAuthTokenCallSuccess = $true
        } catch {
            Write-Host "We fail to create an OAuth token - Exception: $($_.Exception.Message)" -ForegroundColor Red
            $exceptionMessage = $_.Exception.Message
        }
    }
    end {
        return [PSCustomObject]@{
            OAuthToken           = $oAuthToken
            Successful           = $oAuthTokenCallSuccess
            ExceptionMessage     = $exceptionMessage
            LastTokenRefreshTime = (Get-Date)
        }
    }
}

function Get-NewJsonWebToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint,

        [ValidateSet("CurrentUser", "LocalMachine")]
        [Parameter(Mandatory = $false)]
        [string]$CertificateStore = "CurrentUser",

        [Parameter(Mandatory = $false)]
        [string]$Issuer,

        [Parameter(Mandatory = $false)]
        [string]$Audience,

        [Parameter(Mandatory = $false)]
        [string]$Subject,

        [Parameter(Mandatory = $false)]
        [int]$TokenLifetimeInSeconds = 3600,

        [ValidateSet("RS256", "RS384", "RS512")]
        [Parameter(Mandatory = $false)]
        [string]$SigningAlgorithm = "RS256"
    )

    <#
        Shared function to create a signed Json Web Token (JWT) by using a certificate.
        It is also possible to use a secret key to sign the token, but that is not supported in this function.
        The function returns the token as a string if successful, otherwise it returns $null.
        https://www.rfc-editor.org/rfc/rfc7519
        https://learn.microsoft.com/azure/active-directory/develop/active-directory-certificate-credentials
        https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
    }
    process {
        try {
            $certificate = Get-ChildItem Cert:\$CertificateStore\My\$CertificateThumbprint
            if ($certificate.HasPrivateKey) {
                $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
                # Base64url-encoded SHA-1 thumbprint of the X.509 certificate's DER encoding
                $x5t = [System.Convert]::ToBase64String($certificate.GetCertHash())
                $x5t = ((($x5t).Replace("\+", "-")).Replace("/", "_")).Replace("=", "")
                Write-Verbose "x5t is: $x5t"
            } else {
                Write-Verbose "We don't have a private key for certificate: $CertificateThumbprint and so cannot sign the token"
                return
            }
        } catch {
            Write-Verbose "Unable to import the certificate - Exception: $($Error[0].Exception.Message)"
            return
        }

        $header = [ordered]@{
            alg = $SigningAlgorithm
            typ = "JWT"
            x5t = $x5t
        }

        # "iat" (issued at) and "exp" (expiration time) must be UTC and in UNIX time format
        $payload = @{
            iat = [Math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date -Date "01/01/1970")).TotalSeconds)
            exp = [Math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date -Date "01/01/1970")).TotalSeconds) + $TokenLifetimeInSeconds
        }

        # Issuer, Audience and Subject are optional as per RFC 7519
        if (-not([System.String]::IsNullOrEmpty($Issuer))) {
            Write-Verbose "Issuer: $Issuer will be added to payload"
            $payload.Add("iss", $Issuer)
        }

        if (-not([System.String]::IsNullOrEmpty($Audience))) {
            Write-Verbose "Audience: $Audience will be added to payload"
            $payload.Add("aud", $Audience)
        }

        if (-not([System.String]::IsNullOrEmpty($Subject))) {
            Write-Verbose "Subject: $Subject will be added to payload"
            $payload.Add("sub", $Subject)
        }

        $headerJson = $header | ConvertTo-Json -Compress
        $payloadJson = $payload | ConvertTo-Json -Compress

        $headerBase64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($headerJson)).Split("=")[0].Replace("+", "-").Replace("/", "_")
        $payloadBase64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($payloadJson)).Split("=")[0].Replace("+", "-").Replace("/", "_")

        $signatureInput = [System.Text.Encoding]::ASCII.GetBytes("$headerBase64.$payloadBase64")

        Write-Verbose "Header (Base64) is: $headerBase64"
        Write-Verbose "Payload (Base64) is: $payloadBase64"
        Write-Verbose "Signature input is: $signatureInput"

        $signingAlgorithmToUse = switch ($SigningAlgorithm) {
            ("RS384") { [Security.Cryptography.HashAlgorithmName]::SHA384 }
            ("RS512") { [Security.Cryptography.HashAlgorithmName]::SHA512 }
            default { [Security.Cryptography.HashAlgorithmName]::SHA256 }
        }
        Write-Verbose "Signing the Json Web Token using: $SigningAlgorithm"

        $signature = $privateKey.SignData($signatureInput, $signingAlgorithmToUse, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $signature = [Convert]::ToBase64String($signature).Split("=")[0].Replace("+", "-").Replace("/", "_")
    }
    end {
        if ((-not([System.String]::IsNullOrEmpty($headerBase64))) -and
            (-not([System.String]::IsNullOrEmpty($payloadBase64))) -and
            (-not([System.String]::IsNullOrEmpty($signature)))) {
            Write-Verbose "Returning Json Web Token"
            return ("$headerBase64.$payloadBase64.$signature")
        } else {
            Write-Verbose "Unable to create Json Web Token"
            return
        }
    }
}

function CheckTokenExpiry {
    param(
            $ApplicationInfo,
            [ref]$EWSService,
            [ref]$Token,
            [string]$Environment,
            $EWSOnlineURL,
            $AuthScope,
            $AzureADEndpoint
        )

    # if token is going to expire in next 5 min then refresh it
    if($PermissionType -eq "Application") {
    if ($null -eq $script:tokenLastRefreshTime -or $script:tokenLastRefreshTime.AddMinutes(55) -lt (Get-Date)) {
        Write-Verbose "Requesting new OAuth token as the current token expires at $($script:tokenLastRefreshTime)."
        $createOAuthTokenParams = @{
            TenantID                       = $ApplicationInfo.TenantID
            ClientID                       = $ApplicationInfo.ClientID
            Endpoint                       = $AzureADEndpoint
            CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($ApplicationInfo.CertificateThumbprint)))
            Scope                          = $AuthScope
        }

        # Check if we use an app secret or certificate by using regex to match Json Web Token (JWT)
        if ($ApplicationInfo.AppSecret -match "^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)") {
            $jwtParams = @{
                CertificateThumbprint = $ApplicationInfo.CertificateThumbprint
                CertificateStore      = $CertificateStore
                Issuer                = $ApplicationInfo.ClientID
                Audience              = "$AzureADEndpoint/$($ApplicationInfo.TenantID)/oauth2/v2.0/token"
                Subject               = $ApplicationInfo.ClientID
            }
            $jwt = Get-NewJsonWebToken @jwtParams

            if ($null -eq $jwt) {
                Write-Host "Unable to sign a new Json Web Token by using certificate: $($ApplicationInfo.CertificateThumbprint)" -ForegroundColor Red
                exit
            }

            $createOAuthTokenParams.Add("Secret", $jwt)
        } else {
            $createOAuthTokenParams.Add("Secret", $ApplicationInfo.AppSecret)
        }

        $oAuthReturnObject = Get-NewOAuthToken @createOAuthTokenParams
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Host "Unable to refresh EWS OAuth token. Please review the error message below and re-run the script:" -ForegroundColor Red
            Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
            exit
        }
        Write-Host "Obtained a new token" -ForegroundColor Green
        $Script:GraphToken = $oAuthReturnObject.OAuthToken
        $script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
        return $oAuthReturnObject.OAuthToken.access_token
    }
    else {
        return $Script:Token
    }
    }
    else {
        if(($GraphToken.ExpiresOn.LocalDateTime).AddMinutes(-5) -le (Get-Date)) { 
            $Script:GraphToken = Get-MsalToken -ClientId $OAuthClientId -TenantId $OAuthTenantId -Scopes $Script:Scope -ForceRefresh
            $Script:Token = $Script:GraphToken.AccessToken
        }
    }
}

function TestInstalledModules {
    # Function to check if running as Administrator
    function IsAdmin {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    
    if($PermissionType -eq "Delegated") {
        Write-Verbose "Checking for the MSAL.PS PowerShell module."
        if(-not (Get-InstalledModule -Name MSAL.PS -MinimumVersion '4.37.0.0' -ErrorAction SilentlyContinue)) {
            if(-not (IsAdmin)) {
                Write-Host "Administrator privileges required to install 'MSAL.PS' module. Re-run PowerShell or the script as Admin." -ForegroundColor Red
                exit
            }
            else{
                Write-Host "Prerequisite not found: Attempting to install 'MSAL.PS' module..." -ForegroundColor Yellow
                try{
                    Install-Module -Name MSAL.PS -MinimumVersion 4.37.0.0 -Repository PSGallery -Force
                }
                catch{
                    Write-Host "Failed to install 'MSAL.PS' module. Please install it manually." -ForegroundColor Red
                    exit
                }
                # Check again for MSAL.PS module installation
                if(-not (Get-InstalledModule -Name MSAL.PS -MinimumVersion 4.37.0.0)) {
                    Write-Host "Failed to install 'MSAL.PS' module. Please install it manually." -ForegroundColor Red
                    exit
                }
            }
        }
    }
 
    if($ImpersonationCheck) {
        # Test for ExchangeOnlineManagement module
        if(-not (Get-InstalledModule -Name ExchangeOnlineManagement -MinimumVersion 3.4.0)) {
            if(-not (IsAdmin)) {
                Write-Host "Administrator privileges required to install 'ExchangeOnlineManagement' module." -ForegroundColor Red
                exit
            }
            Write-Host "Attempting to install 'ExchangeOnlineManagement' module..." -ForegroundColor Yellow
            Install-Module -Name ExchangeOnlineManagement -MinimumVersion 3.4.0 -Force
        }
 
        # Check again for ExchangeOnlineManagement module installation
        if(-not (Get-InstalledModule -Name ExchangeOnlineManagement -MinimumVersion 3.4.0)) {
            Write-Host "Failed to install 'ExchangeOnlineManagement' module. Please install it manually." -ForegroundColor Red
            exit
        }
    }
}

function Invoke-GraphApiRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query,

        [ValidateSet("v1.0", "beta")]
        [Parameter(Mandatory = $false)]
        [string]$Endpoint = "v1.0",

        [Parameter(Mandatory = $false)]
        [string]$Method = "GET",

        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json",

        [Parameter(Mandatory = $false)]
        [string]$Body,

        [Parameter(Mandatory = $true)]
        [ValidatePattern("^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)")]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [int]$ExpectedStatusCode = 200,

        [Parameter(Mandatory = $true)]
        [string]$GraphApiUrl
    )

    <#
        This shared function is used to make requests to the Microsoft Graph API.
        It returns a PSCustomObject with the following properties:
            Content: The content of the response (converted from JSON to a PSCustomObject)
            Response: The full response object
            StatusCode: The status code of the response
            Successful: A boolean indicating whether the request was successful
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
        $successful = $false
        $content = $null
    }
    process {
        $graphApiRequestParams = @{
            Uri             = "$GraphApiUrl/$Endpoint/$($Query.TrimStart("/"))"
            Header          = @{ Authorization = "Bearer $AccessToken" }
            Method          = $Method
            ContentType     = $ContentType
            UseBasicParsing = $true
            ErrorAction     = "Stop"
        }

        if (-not([System.String]::IsNullOrEmpty($Body))) {
            Write-Verbose "Body: $Body"
            $graphApiRequestParams.Add("Body", $Body)
        }

        Write-Verbose "Graph API uri called: $($graphApiRequestParams.Uri)"
        Write-Verbose "Method: $($graphApiRequestParams.Method) ContentType: $($graphApiRequestParams.ContentType)"
        $graphApiResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $graphApiRequestParams

        if (($null -eq $graphApiResponse) -or
            ([System.String]::IsNullOrEmpty($graphApiResponse.StatusCode))) {
            Write-Verbose "Graph API request failed - no response"
        } elseif ($graphApiResponse.StatusCode -ne $ExpectedStatusCode) {
            Write-Verbose "Graph API status code: $($graphApiResponse.StatusCode) does not match expected status code: $ExpectedStatusCode"
        } else {
            Write-Verbose "Graph API request successful"
            $successful = $true
            $content = $graphApiResponse.Content | ConvertFrom-Json
        }
    }
    end {
        return [PSCustomObject]@{
            Content    = $content
            Response   = $graphApiResponse
            StatusCode = $graphApiResponse.StatusCode
            Successful = $successful
        }
    }
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")]
        [string]
        $Uri,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]
        $UseBasicParsing,

        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")]
        [hashtable]
        $ParametersObject,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [string]
        $OutFile
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    try {
        Invoke-WebRequest @params
    } catch {
        Write-VerboseErrorInformation
    }
}

function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TargetUri
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Verbose "Proxy server configuration detected"
            Write-Verbose $proxyObject.OriginalString
            return $true
        } else {
            Write-Verbose "No proxy server configuration detected"
            return $false
        }
    } catch {
        Write-Verbose "Unable to check for proxy server configuration"
        return $false
    }
}

function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.InvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Get-CloudServiceEndpoint {
    [CmdletBinding()]
    param(
        [string]$EndpointName
    )

    <#
        This shared function is used to get the endpoints for the Azure and Microsoft 365 services.
        It returns a PSCustomObject with the following properties:
            GraphApiEndpoint: The endpoint for the Microsoft Graph API
            ExchangeOnlineEndpoint: The endpoint for Exchange Online
            AutoDiscoverSecureName: The endpoint for Autodiscover
            AzureADEndpoint: The endpoint for Azure Active Directory
            EnvironmentName: The name of the Azure environment
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
    }
    process {
        # https://learn.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints
        switch ($EndpointName) {
            "Global" {
                $environmentName = "AzureCloud"
                $graphApiEndpoint = "https://graph.microsoft.com"
                $exchangeOnlineEndpoint = "https://outlook.office.com"
                $autodiscoverSecureName = "https://autodiscover-s.outlook.com"
                $azureADEndpoint = "https://login.microsoftonline.com"
                break
            }
            "USGovernmentL4" {
                $environmentName = "AzureUSGovernment"
                $graphApiEndpoint = "https://graph.microsoft.us"
                $exchangeOnlineEndpoint = "https://outlook.office365.us"
                $autodiscoverSecureName = "https://autodiscover-s.office365.us"
                $azureADEndpoint = "https://login.microsoftonline.us"
                break
            }
            "USGovernmentL5" {
                $environmentName = "AzureUSGovernment"
                $graphApiEndpoint = "https://dod-graph.microsoft.us"
                $exchangeOnlineEndpoint = "https://outlook-dod.office365.us"
                $autodiscoverSecureName = "https://autodiscover-s-dod.office365.us"
                $azureADEndpoint = "https://login.microsoftonline.us"
                break
            }
            "ChinaCloud" {
                $environmentName = "AzureChinaCloud"
                $graphApiEndpoint = "https://microsoftgraph.chinacloudapi.cn"
                $exchangeOnlineEndpoint = "https://partner.outlook.cn"
                $autodiscoverSecureName = "https://autodiscover-s.partner.outlook.cn"
                $azureADEndpoint = "https://login.partner.microsoftonline.cn"
                break
            }
        }
    }
    end {
        return [PSCustomObject]@{
            EnvironmentName        = $environmentName
            GraphApiEndpoint       = $graphApiEndpoint
            ExchangeOnlineEndpoint = $exchangeOnlineEndpoint
            AutoDiscoverSecureName = $autodiscoverSecureName
            AzureADEndpoint        = $azureADEndpoint
        }
    }
}

function GetAzureAdApplications{
    Write-Host "Getting a list of Entra App registrations..." -ForegroundColor Green
    $Script:AadApplications = New-Object System.Collections.ArrayList
    $AadApplicationResults = Invoke-GraphApiRequest -Query "applications?`$select=appId,createdDateTime,displayName,description,notes,requiredResourceAccess" -AccessToken $Script:Token -GraphApiUrl $cloudService.graphApiEndpoint
    foreach($application in $AadApplicationResults.Content.Value){
        $Script:AadApplications.Add($application) | Out-Null
    }
    # Check if response includes more results link
    while($null -ne $AadApplicationResults.Content.'@odata.nextLink') {
        $Query = $AadApplicationResults.Content.'@odata.nextLink'.Substring($AadApplicationResults.Content.'@odata.nextLink'.IndexOf("applications"))
        $AadApplicationResults = Invoke-GraphApiRequest -Query $Query -AccessToken $Script:Token -GraphApiUrl $cloudService.graphApiEndpoint
        foreach($application in $AadApplicationResults.Content.Value) {
            $Script:AadApplications.Add($application) | Out-Null
        }
    }
    $Script:AadApplications | Export-Csv "$OutputPath\EntraAppRegistrations-$((Get-Date).ToString("yyyyMMddhhmmss")).csv" -NoTypeInformation
}

function GetAzureAdServicePrincipals{
    Write-Host "Getting a list of all Entra service applications..." -ForegroundColor Green
    $Script:ServicePrincipals = New-Object System.Collections.ArrayList
    $ServicePrincipalsResults = Invoke-GraphApiRequest -Query "servicePrincipals?`$select=id,appDisplayName,appDescription,appId,createdDateTime,displayName,servicePrincipalType,appRoles,oauth2PermissionScopes" -AccessToken $Script:Token -GraphApiUrl $cloudService.graphApiEndpoint
    foreach($ServicePrincipal in $ServicePrincipalsResults.Content.Value) {
        $Script:ServicePrincipals.Add($ServicePrincipal) | Out-Null
    }
    # Check if response includes more results link
    while($null -ne $ServicePrincipalsResults.Content.'@odata.nextLink') {
        $Query = $ServicePrincipalsResults.Content.'@odata.nextLink'.Substring($ServicePrincipalsResults.Content.'@odata.nextLink'.IndexOf("servicePrincipals"))
        $ServicePrincipalsResults = Invoke-GraphApiRequest -Query $Query -AccessToken $Script:Token -GraphApiUrl $cloudService.graphApiEndpoint
        foreach($ServicePrincipal in $ServicePrincipalsResults.Content.Value){
            $Script:ServicePrincipals.Add($ServicePrincipal) | Out-Null
        }
    }
    $Script:ServicePrincipals | Export-Csv "$OutputPath\EntraServicePrincipals-$((Get-Date).ToString("yyyyMMddhhmmss")).csv" -NoTypeInformation
}

function GetAppsByApi {
    $Script:ApiPermissions = New-Object System.Collections.ArrayList
    Write-Host "Filtering app registrations that use the $($Api) API..." -ForegroundColor Green
    foreach($application in $Script:AadApplications) {
        Write-Verbose "Checking resource(s) accessed by the applications."
        foreach($RequiredResourceAccess in $application.requiredResourceAccess) {
            Write-Verbose "Finding service princlipals that match the resource."
            $sp = $Script:ServicePrincipals | Where-Object {$_.AppId -eq $RequiredResourceAccess.ResourceAppId}
            $ExchangeOnlineAccess = $false
            $GraphAccess = $false
            if($RequiredResourceAccess.ResourceAppId -eq "00000002-0000-0ff1-ce00-000000000000") {
                $ExchangeOnlineAccess = $true
            }
            elseif ($RequiredResourceAccess.ResourceAppId -eq "00000003-0000-0000-c000-000000000000") {
                $GraphAccess =$true
            }
            if($ExchangeOnlineAccess -or $GraphAccess) {
                foreach($ResourceAccess in $RequiredResourceAccess.ResourceAccess) {
                    if($($ResourceAccess.Type) -eq "Scope") {
                        $AppScope = $sp.Oauth2PermissionScopes | Where-Object {$_.id -eq "$($ResourceAccess.id)"}
                        $Script:AppPermission = [PSCustomObject]@{
                            'ApplicationDisplayName'  = $application.displayName
                            'ApplicationID'           = $application.appId
                            'PermissionType'          = "Delegate"
                            'PermissionValue'         = $AppScope.Value
                            "ResourceDisplayName"     = $RequiredResourceAccess.ResourceAppId
                        }
                        switch ($Api) {
                            'EWS' {
                                if($appScope.value -in $Script:EWSPermissions) {
                                     $Script:ApiPermissions.Add($Script:AppPermission) | Out-Null
                                }
                            }
                            'OutlookRESTv2' {
                                if($appScope.value -in $Script:Delegated_OutlookRESTPermissions -and $ExchangeOnlineAccess) {
                                    $Script:ApiPermissions.Add($Script:AppPermission) | Out-Null
                                }
                                elseif($appScope.value -in $Script:EWSPermissions -and $GraphAccess) {
                                    $Script:ApiPermissions.Add($Script:AppPermission) | Out-Null
                                }
                            }
                        }
                    }
                    elseif ($($ResourceAccess.Type) -eq "Role") {
                        Write-Verbose "Finding application permissions for the application $($application.displayName)."
                        $AppRole = $sp.appRoles | Where-Object {$_.id -eq "$($ResourceAccess.id)"}
                        $Script:AppPermission = [PSCustomObject]@{
                            'ApplicationDisplayName'  = $application.displayName
                            'ApplicationID'           = $application.appId
                            'PermissionType'          = "Application"
                            'PermissionValue'         = $AppRole.Value
                            'ResourceDisplayName'     = $RequiredResourceAccess.ResourceAppId
                        }
                        switch ($Api) {
                            'EWS' {
                                if($appRole.value -in $Script:EWSPermissions -and $ExchangeOnlineAccess) {
                                    $Script:ApiPermissions.Add($Script:AppPermission) | Out-Null
                                }
                            }
                            'OutlookRESTv2' {
                                if($appRole.value -in $Script:Application_OutlookRESTPermissions -and $ExchangeOnlineAccess) {
                                    $Script:ApiPermissions.Add($Script:AppPermission) | Out-Null
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    $Script:ApiPermissions | Export-Csv "$OutputPath\$Api-EntraAppRegistrations-$((Get-Date).ToString("yyyyMMddhhmmss")).csv" -NoTypeInformation
}

function GetEwsSignIns{
    $Script:ApiSignInsFile = "$OutputPath\$Api-SignInEvents-$((Get-Date).ToString("yyyyMMddhhmmss")).csv"
    $ApplicationPermissions = $Script:ApiPermissions | Sort-Object -Property ApplicationId -Unique
    $NumberOfApps = $ApplicationPermissions.Count
    $AppsCompleted = 0
    $StartDate = (Get-Date).AddDays(-$NumberOfDays)
    $TempDate = [datetime]$StartDate
    $TempDate = $TempDate.ToUniversalTime()
    $SearchStartDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempDate
    
    Write-Host "Searching for sign-in events for the $($Api) API..." -ForegroundColor Green
    foreach($App in $ApplicationPermissions) {
        if($ImpersonationCheck) {
            if($App.PermissionValue -ne "EWS.AccessAsUser.All") {
                continue
            }
        }
        CheckTokenExpiry -Token ([ref]$Script:GraphToken) -ApplicationInfo $Script:applicationInfo -AzureADEndpoint $azureADEndpoint
        Write-Progress -Activity "Searching for EWS sign-in attempts" -Status "Checking $($App.ApplicationDisplayName)" -PercentComplete ((($AppsCompleted)/$NumberOfApps)*100)
        $Query = "auditLogs/signIns?`$filter=appid eq '$($App.ApplicationId)' and signInEventTypes/any(t: t eq 'interactiveUser' or t eq 'nonInteractiveUser' or t eq 'servicePrincipal' or t eq 'managedIdentity') and CreatedDateTime ge $SearchStartDate"
        $SignIns = Invoke-GraphApiRequest -Query $Query -Endpoint beta -AccessToken $Script:Token -GraphApiUrl $cloudService.graphApiEndpoint
        $SignIns.Content.Value | Select-Object id, createdDateTime, appId, appDisplayName, correlationId, clientCredentialType, resourceDisplayName, resourceId, servicePrincipalId , userDisplayName, userPrincipalName, @{Name='SignInEventTypes';Expression={$_.signInEventTypes -join '; ' } } | Export-Csv -Path $Script:ApiSignInsFile -NoTypeInformation -NoClobber -Append
        $AppsCompleted++
    }
}

# Start the main script
$loggerParams = @{
    LogDirectory             = $OutputPath
    LogName                  = "ExchangeAppUsage-$((Get-Date).ToString("yyyyMMddhhmmss"))-Debug"
    AppendDateTimeToFileName = $false
    ErrorAction              = "SilentlyContinue"
}

$Script:Logger = Get-NewLoggerInstance @loggerParams

SetWriteHostAction ${Function:Write-HostLog}
SetWriteVerboseAction ${Function:Write-VerboseLog}
SetWriteWarningAction ${Function:Write-HostLog}

#Define variables
$Script:EWSPermissions = @("EWS.AccessAsUser.All", "full_access_as_app")
$Script:Delegated_OutlookRESTPermissions = @("PeopleSettings.Read.All", "PeopleSettings.ReadWrite.All", "ReportingWebService.Read", "Organization.ReadWrite.All",
    "Organization.Read.All", "Mail.ReadBasic", "Notes.Read", "Notes.ReadWrite", "User.Read.All", "User.ReadBasic.All", "MailboxSettings.Read", "Calendars.Read.Shared",
    "Calendars.ReadWrite.Shared", "Mail.Send.Shared", "Mail.ReadWrite.Shared", "Mail.Read.Shared", "Contacts.ReadWrite.Shared", "Contacts.Read.Shared", "Tasks.Read.Shared",
    "Tasks.ReadWrite.Shared", "Mail.Read", "Mail.ReadWrite", "Mail.Send", "Calendars.Read", "Calendars.ReadWrite", "Contacts.Read", "Contacts.ReadWrite", "Group.Read.All",
    "Group.ReadWrite.All", "User.Read", "User.ReadWrite", "User.ReadBasic.All", "People.Read", "People.ReadWrite", "Tasks.Read", "Tasks.ReadWrite", "MailboxSettings.ReadWrite",
    "Contacts.ReadWrite.All", "Contacts.Read.All", "Calendars.ReadWrite.All", "Calendars.Read.All", "Mail.Send.All", "Mail.ReadWrite.All", "Mail.Read.All", "Place.Read.All",
    "OPX.MyDay", "OPX.MyDay.Shared", "OPX.MyDay.All")
$Script:Application_OutlookRESTPermissions = @("PeopleSettings.ReadWrite.All", "PeopleSettings.Read.All", "Organization.ReadWrite.All", "Organization.Read.All",
    "Mailbox.Migration", "User.Read.All", "User.ReadBasic.All", "MailboxSettings.Read", "Mail.Send", "Calendars.Read", "Contacts.Read", "Mail.Read", "Mail.ReadWrite",
    "Contacts.ReadWrite", "MailboxSettings.ReadWrite", "Tasks.Read", "Tasks.ReadWrite", "Calendars.ReadWrite.All", "Calendars.Read.All", "Place.Read.All")

# Call function to confirm required PowerShell module(s) are installed
if($PermissionType -eq "Delegated") {
    TestInstalledModules
}

# Call function to obtain OAuth token
$cloudService = Get-CloudServiceEndpoint $AzureEnvironment
$Script:Scope = "$($cloudService.graphApiEndpoint)/.default"
$azureADEndpoint = $cloudService.AzureADEndpoint

if($PermissionType -eq "Application") {
Write-Host "Requesting an OAuth token to collect the data." -ForegroundColor Green
$applicationInfo = @{
    "TenantID" = $OAuthTenantId
    "ClientID" = $OAuthClientId
}

if ([System.String]::IsNullOrEmpty($OAuthCertificate)) {
    
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($OAuthClientSecret)
    $secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $applicationInfo.Add("AppSecret", $secret)
}
else {
     $jwtParams = @{
        CertificateThumbprint = $OAuthCertificate
        CertificateStore      = $CertificateStore
        Issuer                = $OAuthClientId
        Audience              = "$azureADEndpoint/$OAuthTenantId/oauth2/v2.0/token"
        Subject               = $OAuthClientId
    }
    $jwt = Get-NewJsonWebToken @jwtParams

    if ($null -eq $jwt) {
        Write-Host "Unable to generate Json Web Token by using certificate: $OAuthCertificate" -ForegroundColor Red
        exit
    }

    $applicationInfo.Add("AppSecret", $jwt)
    $applicationInfo.Add("CertificateThumbprint", $OAuthCertificate)
}

$createOAuthTokenParams = @{
    TenantID                       = $OAuthTenantId
    ClientID                       = $OAuthClientId
    Secret                         = $applicationInfo.AppSecret
    Scope                          = $Script:Scope
    Endpoint                       = $azureADEndpoint
    CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($OAuthCertificate)))
}

#Create OAUTH token
$oAuthReturnObject = Get-NewOAuthToken @createOAuthTokenParams
if ($oAuthReturnObject.Successful -eq $false) {
    Write-Host ""
    Write-Host "Unable to fetch an OAuth token for accessing EWS. Please review the error message below and re-run the script:" -ForegroundColor Red
    Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
    exit
}
$Script:GraphToken = $oAuthReturnObject.OAuthToken
$script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
$Script:Token = $Script:GraphToken.access_token
}
else {
    $Script:GraphToken = Get-MsalToken -ClientId $OAuthClientId -TenantId $OAuthTenantId -Interactive -Scopes $Script:Scope
    $Script:Token = $Script:GraphToken.AccessToken
}

# Call function to obtain list of app registrations from Entra
GetAzureADApplications

# Call function to obtain list of service principals from Entra
GetAzureAdServicePrincipals

# Call function to Filter app registrations using the selected API
GetAppsByApi
$Script:ApiPermissions | Format-Table -AutoSize

# Call function to obtain sign-in logs for app registrations using the selected API
if($Script:ApiPermissions.Count -ge 1) {
    GetEwsSignIns
}
else {
    Write-Host "No app registrations were found using the $($Api) API." -ForegroundColor Yellow
}


# Find users with EWS impersonation rights
Write-Host "Checking for possible application impersonation sign-in events..." -ForegroundColor Green
if($ImpersonationCheck){
    if(-not((Get-ConnectionInformation).Name -like "ExchangeOnline*")) {
        Connect-ExchangeOnline
    }
    $ImpersonationAccounts = New-Object System.Collections.ArrayList
    Write-Host "Checking for users with the ApplicationImpersonation role" -ForegroundColor Green
    $ImpersonationRoleAssignments = Get-ManagementRoleAssignment -Role ApplicationImpersonation -GetEffectiveUsers | Where-Object {$_.Name -ne "ApplicationImpersonation-Organization Management-Delegating" -and $_.Name -notlike "ApplicationImpersonation-RIM-MailboxAdmins*"}
    $ImpersonationRoleAssignments | Export-Csv "$OutputPath\ManagementRoleAssignments-$((Get-Date).ToString("yyyyMMddhhmmss")).csv" -NoTypeInformation
    $ApiSignIns = Import-Csv -Path $Script:ApiSignInsFile
    foreach($RoleAssignment in $ImpersonationRoleAssignments) {
        foreach($SignIn in $ApiSignIns) {
            if($SignIn.userDisplayName -eq $RoleAssignment.EffectiveUserName) {
                $ImpersonationAccounts.Add($RoleAssignment.EffectiveUserName) | Out-Null
            }
        }
    }
}
if($ImpersonationAccounts.Count -ge 1) {
$ImpersonationAccounts = $ImpersonationAccounts | Sort-Object -Unique
foreach($i in $ImpersonationAccounts) {
    Write-Host "$($i) has impersonation rights and signed in using Delegated permissions." -ForegroundColor Cyan
}
}
else {
    Write-Host "No users have signed in with impersonation rights in the past $($NumberOfDays) day(s)." -ForegroundColor Yellow
}