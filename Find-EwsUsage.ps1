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

# Version 20250519.1632
[CmdletBinding(DefaultParameterSetName = '__AllParameterSets')]
param (
    [ValidateScript({ Test-Path $_ })]
    [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")]
    [string] $OutputPath,

    [ValidateSet("Global", "USGovernmentL4", "USGovernmentL5", "ChinaCloud")]
    [Parameter(Mandatory = $false)]
    [string]$AzureEnvironment = "Global",

    [Parameter(Mandatory=$false, HelpMessage="The PermissionType parameter specifies whether the app registrations uses delegated or application permissions")] [ValidateSet('Application','Delegated')]
    [string]$PermissionType="Application",
    
    [Parameter(Mandatory=$true,HelpMessage="The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.")] 
    [string]$OAuthClientId,
    
    [Parameter(Mandatory=$true,HelpMessage="The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).")] 
    [string]$OAuthTenantId,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.")] 
    [string]$OAuthRedirectUri = "http://localhost:8004",
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthSecretKey parameter is the the secret for the registered application.")] 
    [SecureString]$OAuthClientSecret,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate parameter is the certificate for the registered application. Certificate auth requires MSAL libraries to be available.")] 
    [string]$OAuthCertificate = $null,
  
    [Parameter(Mandatory=$False,HelpMessage="The CertificateStore parameter specifies the certificate store where the certificate is loaded.")] [ValidateSet("CurrentUser", "LocalMachine")]
    [string]$CertificateStore = $null,

    [Parameter(Mandatory=$false)] [object]$Scope= @("AuditLogsQuery.Read.All"),
    
    [Parameter(ParameterSetName="AppUsage",Mandatory=$False,HelpMessage="The AppId parameter specifies the application ID.")]
    [string]$AppId,

    [Parameter(ParameterSetName="AppUsage",Mandatory=$False,HelpMessage="The AuditQueryId parameter specifies the query ID.")]
    [string]$AuditQueryId,

    [Parameter(Mandatory=$true,HelpMessage="The Operation parameter specifies the operation the script should perform.")] [ValidateSet("GetEwsActivity","GetAppUsage")]
    [string]$Operation = $null,

    [Parameter(Mandatory=$false,HelpMessage="The Operation parameter specifies the operation the script should perform.")] [ValidateSet("NewAuditQuery", "CheckAuditQuery","GetQueryResults")]
    [string]$AuditQueryStep = $null,

    [Parameter(ParameterSetName="AppUsage",Mandatory=$true,HelpMessage="The QueryType parameter specifies the type of query for EWS usage.")] [ValidateSet("SignInLogs","AuditLogs")]
    [string]$QueryType,

    [Parameter(ParameterSetName="AppUsage",Mandatory=$false,HelpMessage="The Name parameter specifies the name for the query and value to be appended to the output file.")]
    [string]$Name,

    [Parameter(ParameterSetName="AppUsage",Mandatory=$False,HelpMessage="The StartDate parameter specifies start date for the audit log query.")]
    [datetime]$StartDate=(Get-Date).AddDays(-7),

    [Parameter(ParameterSetName="AppUsage",Mandatory=$False,HelpMessage="The EndDate parameter specifies the end date for the audit log query.")]
    [datetime]$EndDate=(Get-Date),

    [Parameter(ParameterSetName="AppUsage",Mandatory=$False,HelpMessage="The Interval parameter specifies the number of hours for sign-in log query.")]
    [ValidateRange(1, 24)]
    [int]$Interval=1
)

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
                $managementApiEndpoint = "https://manage.office.com"
                break
            }
            "USGovernmentL4" {
                $environmentName = "AzureUSGovernment"
                $graphApiEndpoint = "https://graph.microsoft.us"
                $exchangeOnlineEndpoint = "https://outlook.office365.us"
                $autodiscoverSecureName = "https://autodiscover-s.office365.us"
                $azureADEndpoint = "https://login.microsoftonline.us"
                $managementApiEndpoint = "https://manage.office365.us"
                break
            }
            "USGovernmentL5" {
                $environmentName = "AzureUSGovernment"
                $graphApiEndpoint = "https://dod-graph.microsoft.us"
                $exchangeOnlineEndpoint = "https://outlook-dod.office365.us"
                $autodiscoverSecureName = "https://autodiscover-s-dod.office365.us"
                $azureADEndpoint = "https://login.microsoftonline.us"
                $managementApiEndpoint = "https://manage.protection.apps.mil"
                break
            }
            "ChinaCloud" {
                $environmentName = "AzureChinaCloud"
                $graphApiEndpoint = "https://microsoftgraph.chinacloudapi.cn"
                $exchangeOnlineEndpoint = "https://partner.outlook.cn"
                $autodiscoverSecureName = "https://autodiscover-s.partner.outlook.cn"
                $azureADEndpoint = "https://login.partner.microsoftonline.cn"
                $managementApiEndpoint = "https://manage.office.cn"
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
            ManagementApiEndpoint  = $managementApiEndpoint
        }
    }
}

function Get-NewJsonWebToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$CertificateThumbprint,
        [ValidateSet("CurrentUser", "LocalMachine")][Parameter(Mandatory = $false)][string]$CertificateStore = "CurrentUser",
        [Parameter(Mandatory = $false)][string]$Issuer,
        [Parameter(Mandatory = $false)][string]$Audience,
        [Parameter(Mandatory = $false)][string]$Subject,
        [Parameter(Mandatory = $false)][int]$TokenLifetimeInSeconds = 3600,
        [ValidateSet("RS256", "RS384", "RS512")][Parameter(Mandatory = $false)][string]$SigningAlgorithm = "RS256"
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

function Get-ApplicationAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$TenantID,
        [Parameter(Mandatory = $true)][string]$ClientID,
        [Parameter(Mandatory = $true)][string]$Secret,
        [Parameter(Mandatory = $true)][string]$Endpoint,
        [Parameter(Mandatory = $false)][string]$TokenService = "oauth2/v2.0/token",
        [Parameter(Mandatory = $false)][switch]$CertificateBasedAuthentication,
        [Parameter(Mandatory = $true)][string]$Scope
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
    if ($null -eq $script:tokenLastRefreshTime -or $script:tokenLastRefreshTime.AddMinutes(55) -lt (Get-Date)) {
        Write-Verbose "Requesting new OAuth token as the current token expires at $($script:tokenLastRefreshTime)."
        if($PermissionType -eq "Application") {
        $createOAuthTokenParams = @{
            TenantID                       = $ApplicationInfo.TenantID
            ClientID                       = $ApplicationInfo.ClientID
            Endpoint                       = $AzureADEndpoint
            CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($ApplicationInfo.CertificateThumbprint)))
            #Scope                          = $AuthScope
            Scope                           = $Script:GraphScope
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

        $oAuthReturnObject = Get-ApplicationAccessToken @createOAuthTokenParams
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Host "Unable to refresh EWS OAuth token. Please review the error message below and re-run the script:" -ForegroundColor Red
            Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
            exit
        }
        Write-Host "Obtained a new token" -ForegroundColor Green
        $Script:Token = $oAuthReturnObject.OAuthToken.access_token
        $script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
        #return $oAuthReturnObject.OAuthToken.access_token
        }
        else {
            #$connectionSuccessful = $false
    
            # Request an authorization code from the Microsoft Azure Active Directory endpoint
            $redeemAuthCodeParams = @{
                Uri             = "$AzureADEndpoint/organizations/oauth2/v2.0/token"
                Method          = "POST"
                ContentType     = "application/x-www-form-urlencoded"
                Body            = @{
                    client_id     = $ApplicationInfo.ClientID
                    scope         = $AuthScope
                    grant_type    = "refresh_token"
                    refresh_token =  $Script:RefreshToken
                }
                UseBasicParsing = $true
            }
            $redeemAuthCodeResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $redeemAuthCodeParams

            if ($redeemAuthCodeResponse.StatusCode -eq 200) {
                $tokens = $redeemAuthCodeResponse.Content | ConvertFrom-Json
                $script:tokenLastRefreshTime = (Get-Date)
                $Script:RefreshToken = $tokens.refresh_token
                $Script:Token = $tokens.access_token
            } 
            else {
                Write-Host "Unable to redeem the authorization code for an access token." -ForegroundColor Red
                exit
            }
        }
    }
    #return $Script:Token
}

function Get-DelegatedAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$AzureADEndpoint = "https://login.microsoftonline.com",
        [Parameter(Mandatory = $false)][string]$GraphApiUrl = "https://graph.microsoft.com",
        [Parameter(Mandatory = $false)][string]$Scope = "$($GraphApiUrl)//Mail.Read email openid profile offline_access",
        [Parameter(Mandatory = $false)][string]$ClientID,
        [Parameter(Mandatory = $false)][string]$RedirectUri
    )

    <#
        This function is used to get an access token for the Azure Graph API by using the OAuth 2.0 authorization code flow
        with PKCE (Proof Key for Code Exchange). The OAuth 2.0 authorization code grant type, or auth code flow,
        enables a client application to obtain authorized access to protected resources like web APIs.
        The auth code flow requires a user-agent that supports redirection from the authorization server
        (the Microsoft identity platform) back to your application.

        More information about the auth code flow with PKCE can be found here:
        https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-auth-code-flow#protocol-details
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
       
        $responseType = "code" # Provides the code as a query string parameter on our redirect URI
        $prompt = "select_account" # We want to show the select account dialog
        $codeChallengeMethod = "S256" # The code challenge method is S256 (SHA256)
        $codeChallengeVerifier = Get-NewS256CodeChallengeVerifier
        $state = ([guid]::NewGuid()).Guid
        $connectionSuccessful = $false
    }
    process {
        $codeChallenge = $codeChallengeVerifier.CodeChallenge
        $codeVerifier = $codeChallengeVerifier.Verifier

        # Request an authorization code from the Microsoft Azure Active Directory endpoint
        $authCodeRequestUrl = "$AzureADEndpoint/organizations/oauth2/v2.0/authorize?client_id=$clientId" +
        "&response_type=$responseType&redirect_uri=$redirectUri&scope=$scope&state=$state&prompt=$prompt" +
        "&code_challenge_method=$codeChallengeMethod&code_challenge=$codeChallenge"

        Start-Process -FilePath $authCodeRequestUrl
        $authCodeResponse = Start-LocalListener

        if ($null -ne $authCodeResponse) {
            # Redeem the returned code for an access token
            $redeemAuthCodeParams = @{
                Uri             = "$AzureADEndpoint/organizations/oauth2/v2.0/token"
                Method          = "POST"
                ContentType     = "application/x-www-form-urlencoded"
                Body            = @{
                    client_id     = $ClientID
                    scope         = $Scope
                    code          = ($($authCodeResponse.Split("=")[1]).Split("&")[0])
                    redirect_uri  = $RedirectUri
                    grant_type    = "authorization_code"
                    code_verifier = $codeVerifier
                }
                UseBasicParsing = $true
            }
            $redeemAuthCodeResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $redeemAuthCodeParams

            if ($redeemAuthCodeResponse.StatusCode -eq 200) {
                $tokens = $redeemAuthCodeResponse.Content | ConvertFrom-Json
                $connectionSuccessful = $true
            } else {
                Write-Host "Unable to redeem the authorization code for an access token." -ForegroundColor Red
            }
        } else {
            Write-Host "Unable to acquire an authorization code from the Microsoft Azure Active Directory endpoint." -ForegroundColor Red
        }
    }
    end {
        if ($connectionSuccessful) {
            return [PSCustomObject]@{
                AccessToken = $tokens.access_token
                RefreshToken = $tokens.refresh_token
                #TenantId    = (Convert-JsonWebTokenToObject $tokens.id_token).Payload.tid
                LastTokenRefreshTime = (Get-Date)
                Successful           = $true
            }
        }
        exit
    }
}

function Convert-JsonWebTokenToObject {
    param(
        [Parameter(Mandatory = $true)][ValidatePattern("^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)")][string]$Token
    )

    <#
        This function can be used to split a JSON web token (JWT) into its header, payload, and signature.
        The JWT is expected to be in the format of <header>.<payload>.<signature>.
        The function returns a PSCustomObject with the following properties:
            Header    - The header of the JWT
            Payload   - The payload of the JWT
            Signature - The signature of the JWT

            It returns $null if the JWT is not in the expected format or conversion fails.
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
        function ConvertJwtFromBase64StringWithoutPadding {
            param(
                [Parameter(Mandatory = $true)]
                [string]$Jwt
            )
            $Jwt = ($Jwt.Replace("-", "+")).Replace("_", "/")
            switch ($Jwt.Length % 4) {
                0 { return [System.Convert]::FromBase64String($Jwt) }
                2 { return [System.Convert]::FromBase64String($Jwt + "==") }
                3 { return [System.Convert]::FromBase64String($Jwt + "=") }
                default { throw "The JWT is not a valid Base64 string." }
            }
        }
    }
    process {
        $tokenParts = $Token.Split(".")
        $tokenHeader = $tokenParts[0]
        $tokenPayload = $tokenParts[1]
        $tokenSignature = $tokenParts[2]

        Write-Verbose "Now processing token header..."
        $tokenHeaderDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenHeader))

        Write-Verbose "Now processing token payload..."
        $tokenPayloadDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenPayload))

        Write-Verbose "Now processing token signature..."
        $tokenSignatureDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenSignature))
    }
    end {
        if (($null -ne $tokenHeaderDecoded) -and
            ($null -ne $tokenPayloadDecoded) -and
            ($null -ne $tokenSignatureDecoded)) {
            Write-Verbose "Conversion of the token was successful"
            return [PSCustomObject]@{
                Header    = ($tokenHeaderDecoded | ConvertFrom-Json)
                Payload   = ($tokenPayloadDecoded | ConvertFrom-Json)
                Signature = $tokenSignatureDecoded
            }
        }

        Write-Verbose "Conversion of the token failed"
        return $null
    }
}

function Get-NewS256CodeChallengeVerifier {
    param()

    <#
        This function can be used to generate a new SHA256 code challenge and verifier following the PKCE specification.
        The Proof Key for Code Exchange (PKCE) extension describes a technique for public clients to mitigate the threat
        of having the authorization code intercepted. The technique involves the client first creating a secret,
        and then using that secret again when exchanging the authorization code for an access token.

        The function returns a PSCustomObject with the following properties:
        Verifier: The verifier that was generated
        CodeChallenge: The code challenge that was generated

        It returns $null if the code challenge and verifier generation fails.

        More information about the auth code flow with PKCE can be found here:
        https://www.rfc-editor.org/rfc/rfc7636
    #>

    Write-Verbose "Calling $($MyInvocation.MyCommand)"

    $bytes = [System.Byte[]]::new(64)
    ([System.Security.Cryptography.RandomNumberGenerator]::Create()).GetBytes($bytes)
    $b64String = [Convert]::ToBase64String($bytes)
    $verifier = (($b64String.TrimEnd("=")).Replace("+", "-")).Replace("/", "_")

    $newMemoryStream = [System.IO.MemoryStream]::new()
    $newStreamWriter = [System.IO.StreamWriter]::new($newMemoryStream)
    $newStreamWriter.write($verifier)
    $newStreamWriter.Flush()
    $newMemoryStream.Position = 0
    $hash = Get-FileHash -InputStream $newMemoryStream | Select-Object Hash
    $hex = $hash.Hash

    $bytesArray = [byte[]]::new($hex.Length / 2)

    for ($i = 0; $i -lt $hex.Length; $i+=2) {
        $bytesArray[$i/2] = [Convert]::ToByte($hex.Substring($i, 2), 16)
    }

    $base64Encoded = [Convert]::ToBase64String($bytesArray)
    $base64UrlEncoded = (($base64Encoded.TrimEnd("=")).Replace("+", "-")).Replace("/", "_")

    if ((-not([System.String]::IsNullOrEmpty($verifier))) -and
        (-not([System.String]::IsNullOrEmpty(($base64UrlEncoded))))) {
        Write-Verbose "Verifier and CodeChallenge generated successfully"
        return [PSCustomObject]@{
            Verifier      = $verifier
            CodeChallenge = $base64UrlEncoded
        }
    }

    Write-Verbose "Verifier and CodeChallenge generation failed"
    return $null
}

function Start-LocalListener {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Only non-destructive operations are performed in this function.')]
    param(
        [Parameter(Mandatory = $false)][int]$Port = 8004,
        [Parameter(Mandatory = $false)][int]$TimeoutSeconds = 60,
        [Parameter(Mandatory = $false)][string]$UrlContains = "code=",
        [Parameter(Mandatory = $false)][string]$ExpectedHttpMethod = "GET",
        [Parameter(Mandatory = $false)][string]$ResponseOutput = "Authentication complete. You can return to the application. Feel free to close this browser tab."
    )

    <#
        This function is used to start a local listener on the specified port (default: 8004).
        It will wait for the specified amount of seconds (default: 60) for a request to be made.
        The function will return the URL of the request that was made.
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
        $url = $null
        $signalled = $false
        $stopwatch = New-Object System.Diagnostics.Stopwatch
        $listener = New-Object Net.HttpListener
    }
    process {
        $listener.Prefixes.add("http://localhost:$($Port)/")
        try {
            Write-Verbose "Starting listener..."
            Write-Verbose "Listening on port: $($Port)"
            Write-Verbose "Waiting $($TimeoutSeconds) seconds for request to be made to url that contains: $($UrlContains)"
            $stopwatch.Start()
            $listener.Start()

            while ($listener.IsListening) {
                $task = $listener.GetContextAsync()

                while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
                    if ($task.AsyncWaitHandle.WaitOne(100)) {
                        $signalled = $true
                        break
                    }
                    Start-Sleep -Milliseconds 100
                }

                if ($signalled) {
                    $context = $task.GetAwaiter().GetResult()
                    $request = $context.Request
                    $response = $context.Response
                    $url = $request.RawUrl
                    $content = [byte[]]@()

                    if (($url.Contains($UrlContains)) -and
                        ($request.HttpMethod -eq $ExpectedHttpMethod)) {
                        Write-Verbose "Request made to listener and url that was called is as expected. HTTP Method: $($request.HttpMethod)"
                        $content = [System.Text.Encoding]::UTF8.GetBytes($ResponseOutput)
                        $response.StatusCode = 200 # OK
                        $response.OutputStream.Write($content, 0, $content.Length)
                        $response.Close()
                        break
                    } else {
                        Write-Verbose "Request made to listener but the url that was called is not as expected. URL: $($url)"
                        $response.StatusCode = 404 # Not Found
                        $response.OutputStream.Write($content, 0, $content.Length)
                        $response.Close()
                        break
                    }
                } else {
                    Write-Verbose "Timeout of $($TimeoutSeconds) seconds reached..."
                    break
                }
            }
        } finally {
            Write-Verbose "Stopping listener..."
            Start-Sleep -Seconds 2
            $stopwatch.Stop()
            $listener.Stop()
        }
    }
    end {
        return $url
    }
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")][string]$Uri,
        [Parameter(Mandatory = $false, ParameterSetName = "Default")][switch]$UseBasicParsing,
        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")][hashtable]$ParametersObject,
        [Parameter(Mandatory = $false, ParameterSetName = "Default")][string]$OutFile
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
        [Parameter(Mandatory = $true)][string]$TargetUri
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
        [ValidateSet("Write-Host", "Write-Verbose")][string]$Cmdlet
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

function Get-OAuthToken {
    param(
        [array]$AppScope,
        [string]$ApiEndpoint
    )
    if($PermissionType -eq "Application") {
        $Script:GraphScope = "$($ApiEndpoint)/.default"
        if ([System.String]::IsNullOrEmpty($OAuthCertificate)) {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($OAuthClientSecret)
            $Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $Script:applicationInfo.Add("AppSecret", $Secret)
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
                Write-Host "Unable to generate Json Web Token by using certificate: $CertificateThumbprint" -ForegroundColor Red
                exit
            }
    
            $Script:applicationInfo.Add("AppSecret", $jwt)
            $Script:applicationInfo.Add("CertificateThumbprint", $OAuthCertificate)
        }
    
        $createOAuthTokenParams = @{
            TenantID                       = $OAuthTenantId
            ClientID                       = $OAuthClientId
            Secret                         = $Script:applicationInfo.AppSecret
            Scope                          = $Script:GraphScope
            Endpoint                       = $azureADEndpoint
            CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($OAuthCertificate)))
        }
    
        #Create OAUTH token
        $oAuthReturnObject = Get-ApplicationAccessToken @createOAuthTokenParams
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Host "Unable to fetch an OAuth token for accessing EWS. Please review the error message below and re-run the script:" -ForegroundColor Red
            Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
            exit
        }
        $Script:Token = $oAuthReturnObject.OAuthToken.access_token
        $Script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
    }
    elseif ($PermissionType -eq "Delegated") {
        if(-not(($AppScope.Contains("email")))) {
            $AppScope += "email"
        }
        if(-not(($AppScope.Contains("openid")))) {
            $AppScope += "openid"
        }
        if(-not(($AppScope.Contains("offline_access")))) {
            $AppScope += "offline_access"
        }
        $Script:GraphScope = "$($ApiEndpoint)//$($Scope)"
        $oAuthReturnObject = Get-DelegatedAccessToken -AzureADEndpoint $cloudService.AzureADEndpoint -GraphApiUrl $cloudService.GraphApiEndpoint -Scope $Script:GraphScope -ClientID $OAuthClientId -RedirectUri $OAuthRedirectUri
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Host "Unable to fetch an OAuth token for accessing EWS. Please review the error message below and re-run the script:" -ForegroundColor Red
            Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
            exit
        }    
        $Script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
        $Script:Token = $oAuthReturnObject.AccessToken
        $Script:RefreshToken = $oAuthReturnObject.RefreshToken
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
        $Body,

        [Parameter(Mandatory = $true)]
        [ValidatePattern("^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)")]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [int]$ExpectedStatusCode = 200,

        [Parameter(Mandatory = $true)]
        [string]$GraphApiUrl,

        [Parameter(Mandatory = $false)]
        [boolean]$Search = $false
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

        if($Search) {
            $graphApiRequestParams.Header.add("ConsistencyLevel", "eventual")
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

function getFileName{
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    #path should end with \
    if (-not ($OutputPath.EndsWith("\"))) {
        $OutputPath = "$($OutputPath)\"
    }

    # path should not be on root drive
    if ($OutputPath.EndsWith(":\")) {
        $OutputPath += "results\"
    }

    # verify folder exists, if not try to create it
    if (!(Test-Path($OutputPath))) {
        Write-Host -ForegroundColor Yellow ">> Warning: '$OutputPath' does not exist. Creating one now..."
        Write-host -ForegroundColor Gray "Creating '$OutputPath': " -NoNewline
        try {
            New-Item -ItemType "directory" -Path $OutputPath -ErrorAction Stop | Out-Null
            Write-Host -ForegroundColor Green "Path '$OutputPath' has been created successfully"
        }
        catch {
            write-host -ForegroundColor Red "FAILED to create '$OutputPath'"
            Write-Host -ForegroundColor Red ">> ERROR: The directory '$OutputPath' could not be created."
            Write-Host -ForegroundColor Red $error[0]
        }
    }
    else{
        Write-Verbose "Path '$OutputPath' already exists"
    }
    if([string]::IsNullOrEmpty($Name)){
        $CSVfilename = "$($OutputPath)Ews-Usage-Results.csv"
    }
    else{
        $CSVfilename = "$($OutputPath)Ews-Usage-Results-$($Name).csv"
    }
    if((Test-Path($CSVfilename)) -and $Operation -like "*Query*") {
        Remove-Item $CSVfilename -Confirm:$false -Force
    }
    return $CSVfilename
}

function GetAzureAdApplications{
    Write-Host "Getting a list of Entra App registrations..." -ForegroundColor Green
    $Script:AadApplications = New-Object System.Collections.ArrayList
    $AadApplicationResults = Invoke-GraphApiRequest -Query "applications?`$select=appId,createdDateTime,displayName,description,notes,requiredResourceAccess" -AccessToken $Script:Token -GraphApiUrl $APIResource
    foreach($application in $AadApplicationResults.Content.Value){
        $Script:AadApplications.Add($application) | Out-Null
    }
    # Check if response includes more results link
    while($null -ne $AadApplicationResults.Content.'@odata.nextLink') {
        $Query = $AadApplicationResults.Content.'@odata.nextLink'.Substring($AadApplicationResults.Content.'@odata.nextLink'.IndexOf("applications"))
        $AadApplicationResults = Invoke-GraphApiRequest -Query $Query -AccessToken $Script:Token -GraphApiUrl $APIResource
        foreach($application in $AadApplicationResults.Content.Value) {
            $Script:AadApplications.Add($application) | Out-Null
        }
    }
    $Script:AadApplications | Export-Csv "$OutputPath\EntraAppRegistrations-$((Get-Date).ToString("yyyyMMddhhmmss")).csv" -NoTypeInformation
}

function GetAzureAdServicePrincipals{
    Write-Host "Getting a list of all Entra service applications..." -ForegroundColor Green
    $Script:ServicePrincipals = New-Object System.Collections.ArrayList
    $ServicePrincipalsResults = Invoke-GraphApiRequest -Query "servicePrincipals?`$select=id,appDisplayName,appDescription,appId,createdDateTime,displayName,servicePrincipalType,appRoles,oauth2PermissionScopes" -AccessToken $Script:Token -GraphApiUrl $APIResource
    foreach($ServicePrincipal in $ServicePrincipalsResults.Content.Value) {
        $Script:ServicePrincipals.Add($ServicePrincipal) | Out-Null
    }
    # Check if response includes more results link
    while($null -ne $ServicePrincipalsResults.Content.'@odata.nextLink') {
        $Query = $ServicePrincipalsResults.Content.'@odata.nextLink'.Substring($ServicePrincipalsResults.Content.'@odata.nextLink'.IndexOf("servicePrincipals"))
        $ServicePrincipalsResults = Invoke-GraphApiRequest -Query $Query -AccessToken $Script:Token -GraphApiUrl $APIResource
        foreach($ServicePrincipal in $ServicePrincipalsResults.Content.Value){
            $Script:ServicePrincipals.Add($ServicePrincipal) | Out-Null
        }
    }
    $Script:ServicePrincipals | Export-Csv "$OutputPath\EntraServicePrincipals-$((Get-Date).ToString("yyyyMMddhhmmss")).csv" -NoTypeInformation
}

function GetAzureAdOauth2PermissionGrants{
    write-host "Getting a list of all Entra OAuth2 permission grants..." -ForegroundColor Green
    $Script:Oauth2PermissionGrants = New-Object System.Collections.ArrayList
    $Oauth2PermissionGrantsResults = Invoke-GraphApiRequest -Query "oauth2PermissionGrants?`$select=clientId,scope,resourceId" -AccessToken $Script:Token -GraphApiUrl $APIResource
    foreach($Oauth2PermissionGrant in $Oauth2PermissionGrantsResults.Content.Value) {
        $Script:Oauth2PermissionGrants.Add($Oauth2PermissionGrant) | Out-Null
    }
    # Check if response includes more results link
    while($null -ne $Oauth2PermissionGrantsResults.Content.'@odata.nextLink') {
        $Query = $Oauth2PermissionGrantsResults.Content.'@odata.nextLink'.Substring($Oauth2PermissionGrantsResults.Content.'@odata.nextLink'.IndexOf("oauth2PermissionGrants"))
        $Oauth2PermissionGrantsResults = Invoke-GraphApiRequest -Query $Query -AccessToken $Script:Token -GraphApiUrl $APIResource
        foreach($Oauth2PermissionGrant in $Oauth2PermissionGrantsResults.Content.Value){
            $Script:Oauth2PermissionGrants.Add($Oauth2PermissionGrant) | Out-Null
        }
    }
    $Script:Oauth2PermissionGrants | Export-Csv "$OutputPath\Oauth2PermissionGrants-$((Get-Date).ToString("yyyyMMddhhmmss")).csv" -NoTypeInformation
}

function GetAppsWithApplicationPermissions {
    $Script:ApiPermissions = New-Object System.Collections.ArrayList
    $ExoSpn = $Script:ServicePrincipals | Where-Object {$_.appId -eq "00000002-0000-0ff1-ce00-000000000000"}
    $EwsAccessAsApp = ($ExoSpn.AppRoles | Where-Object {$_.Value -eq "full_access_as_app"}).Id
    
    foreach($application in $Script:AadApplications){
        Write-Verbose "Checking application $($application.displayName) for EWS full_access_as_app permission..."
        foreach($r in $application.requiredResourceAccess) {
            if($r.ResourceAccess.Id -eq $EwsAccessAsApp){
                $Script:AppPermission = [PSCustomObject]@{
                            'ApplicationDisplayName'  = $application.displayName
                            'ApplicationID'           = $application.appId
                            'PermissionType'          = "Application"
                            'PermissionValue'         = "full_access_as_app"
                            'ResourceDisplayName'     = $application.RequiredResourceAccess.ResourceAppId
                        }
                        $Script:ApiPermissions.Add($Script:AppPermission) | Out-Null
            }
        }
    }
}

function GetAppsByOAuthPermissionGrant{
    Write-Host "Getting application IDs for applications that have the EWS.AccessAsUser.All permission..." -ForegroundColor Green
    foreach($Oauth2PermissionGrant in $Script:Oauth2PermissionGrants) {
        if($Oauth2PermissionGrant.scope -like "*EWS.AccessAsUser.All*") {
            $AadApplicationResults = $Script:ServicePrincipals | Where-Object {$_.Id -eq $Oauth2PermissionGrant.clientId}
            $Script:AppPermission = [PSCustomObject]@{
                    'ApplicationDisplayName'  = $AadApplicationResults.displayName
                    'ApplicationID'           = $AadApplicationResults.appId
                    'PermissionType'          = "Delegate"
                    'PermissionValue'         = "EWS.AccessAsUser.All"
                    'ResourceDisplayName'     = $Oauth2PermissionGrant.resourceId
            }
            $Script:ApiPermissions.Add($Script:AppPermission) | Out-Null
        }
    }
}

function CreateAuditQuery{
    param(
        [Parameter(Mandatory=$false)]
        [string]$Admin
    )
    try{
        Write-Host "Attempting to create a new audit query..." -ForegroundColor Cyan -NoNewline
        $Body = $Body | ConvertTo-JSON -Depth 6
        $q = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query "security/auditLog/queries" -AccessToken $Script:Token -Method POST -Body $Body -Endpoint beta # | Out-Null
        [string]$auditQuery = ($q.Response.Content | ConvertFrom-Json).id
        Write-Host "OK" -ForegroundColor Green
        $CSVfilename = getFileName -OutputPath $outputPath
        @{auditQueryId=$auditQuery} | Export-Csv $CSVfilename -NoTypeInformation
    }
    catch {
        Write-Host "FAILED" -ForegroundColor Red
        Write-Host "Failed to create the audit query." -ForegroundColor Red
    }
    Write-Host "New audit log query created with the id: "#$($auditQuery)"
    return $auditQuery
}

#Define variables
$Script:EWSPermissions = @("EWS.AccessAsUser.All", "full_access_as_app")
$cloudService = Get-CloudServiceEndpoint $AzureEnvironment
$azureADEndpoint = $cloudService.AzureADEndpoint
$Script:applicationInfo = @{
    "TenantID" = $OAuthTenantId
    "ClientID" = $OAuthClientId
}
$APIResource = $cloudService.GraphApiEndpoint

switch($Operation) {
    "GetAppUsage"{
        switch($QueryType){
            "SignInLogs"{
                $Scope= @("AuditLog.Read.All")
                Get-OAuthToken -AppScope $Scope -ApiEndpoint $APIResource
                Write-Host "Searching sign-in logs for the application $($AppId)..." -ForegroundColor Green
                $OutputFile = "$OutputPath\$($AppId)-SignInEvents-$((Get-Date).ToString("yyyyMMddhhmmss")).csv"
                
                # Breaking down the sign-in logs requests to hour intervals
                $StartTime = $StartDate
                while ($StartTime -lt $EndDate) {
                    $EndTime = $StartTime.AddHours($Interval)
                    if ($EndTime -gt $EndDate) {
                        $EndTime = $EndDate
                    }
                    $EndTime = $EndTime.ToUniversalTime()
                    $StartTime = $StartTime.ToUniversalTime()
                    $EndSearch = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $EndTime
                    $StartSearch = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $StartTime
                    $CurrentProgress = $ProgressPreference
                    $ProgressPreference = "Continue"
                    $CurrentView = $PSStyle.Progress.View
                    try { $PSStyle.Progress.View = "Classic"}
                    catch { $CurrentView = $null }
                    # Set the filter for the sign-in log query
                    Write-Progress -Activity "Getting sign-in activity for $($Appid)" -CurrentOperation "Searching between $($StartSearch) and $($EndSearch)"
                    $Query = "auditLogs/signIns?`$filter=createdDateTime ge $StartSearch and createdDateTime le $EndSearch and appId eq '$($AppId)' and (signInEventTypes/any(t:t eq 'interactiveUser') or signInEventTypes/any(t:t eq 'nonInteractiveUser') or signInEventTypes/any(t:t eq 'servicePrincipal'))"
                    $Script:SignInEvents = New-Object System.Collections.ArrayList
                    $results = Invoke-GraphApiRequest -Query $Query -AccessToken $Script:Token -GraphApiUrl $APIResource -Endpoint beta
                    foreach($r in $results.Content.Value){
                        New-Object -TypeName PSCustomObject -Property @{
                            CreatedDateTime = $r.createdDateTime
                            UserPrincipalName = $r.userPrincipalName
                            AppId = $r.appId
                            AppDisplayName = $r.appDisplayName
                            IsInteractive = $r.isInteractive
                            ClientCredentialType = $r.clientCredentialType
                            SignInEventTypes = $r.signInEventTypes -join ","
                        } | Export-Csv -Path $OutputFile -NoTypeInformation -Append
                    }
                    
                    # Check if response includes more results link
                    while($null -ne $results.Content.'@odata.nextLink') {
                        $Query = $results.Content.'@odata.nextLink'.Substring($results.Content.'@odata.nextLink'.IndexOf("auditLog"))
                        $results = Invoke-GraphApiRequest -Query $Query -AccessToken $Script:Token -GraphApiUrl $APIResource -Endpoint beta
                        foreach($r in $results.Content.Value) {
                            New-Object -TypeName PSCustomObject -Property @{
                                CreatedDateTime = $r.createdDateTime
                                UserPrincipalName = $r.userPrincipalName
                                AppId = $r.appId
                                AppDisplayName = $r.appDisplayName
                                IsInteractive = $r.isInteractive
                                ClientCredentialType = $r.clientCredentialType
                                SignInEventTypes = $r.signInEventTypes -join ","
                            } | Export-Csv -Path $OutputFile -NoTypeInformation -Append
                        }
                    }

                    # Move to the next hour
                    $StartTime = $EndTime
                }

                # Display the results in a grid view
                if(Test-Path $OutputFile) {
                    Import-Csv $OutputFile | Group-Object -Property userPrincipalName,signInEventTypes,appDisplayName | Select-Object Count,@{n='UPN';e={$_.Group[0].userPrincipalName}},@{n='AppId';e={$_.Group[0].AppId}},@{n='signInEventType';e={$_.Group[0].signInEventTypes}},@{n='appDisplayName';e={$_.Group[0].appDisplayName}} | Sort-Object Count -Descending | Out-GridView -Title "EWS Application Results"
                }

                if(-not([string]::IsNullOrEmpty($CurrentView))){
                    $PSStyle.Progress.View = $CurrentView
                }
                $ProgressPreference = $CurrentProgress
                
            }
            "AuditLogs"{
                $Scope= @("AuditLogsQuery.Read.All")
                Get-OAuthToken -AppScope $Scope -ApiEndpoint $APIResource
                switch($AuditQueryStep){
                    "NewAuditQuery" {
                        $Hour = (New-TimeSpan -Minutes 60).Ticks
                        $EndTime = (Get-Date -Date $EndDate).ToUniversalTime()
                        $Ticks = ([Math]::Round($EndTime.Ticks / $Hour, 0) * $Hour) -as [long]
                        $EndTime = [datetime]$Ticks
                        $StartTime = (Get-Date -Date $StartDate).ToUniversalTime()
                        $Ticks = ([Math]::Round($StartTime.Ticks / $Hour, 0) * $Hour) -as [long]
                        $StartTime = [datetime]$Ticks
                        $EndSearch = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $EndTime
                        $StartSearch = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $StartTime
                        $Body = @{
                            "@odata.type"       = "#microsoft.graph.security.auditLogQuery"
                            filterStartDateTime = $StartSearch
                            filterEndDateTime   = $EndSearch
                            displayName         = "Audit-query-$($Name)"
                            recordTypeFilters   = @("exchangeItem", "exchangeAggregatedOperation", "exchangeItemAggregated", "exchangeItemGroup")
                            keywordFilter       = "Client=WebServices"
                            serviceFilter       = "Exchange"
                        }

                        #Create the audit query
                        $AuditQueryId = CreateAuditQuery
                        return $AuditQueryId
                        exit
                    }
                    "CheckAuditQuery" {
                        Write-Host "Checking the audit query status for $($AuditQueryId)..." -ForegroundColor Cyan -NoNewline
                        $q = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query "security/auditLog/queries/$($AuditQueryId)" -AccessToken $Script:Token -Method GET -Endpoint beta
                        Write-Host $q.content.status
                        exit
                    }
                    "GetQueryResults" {
                        Write-Host "Checking the audit query status for $($AuditQueryId)..." -ForegroundColor Cyan -NoNewline
                        $q = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query "security/auditLog/queries/$($AuditQueryId)" -AccessToken $Script:Token -Method GET -Endpoint beta
                        if ($q.Content.status -eq "succeeded") {
                            Write-Host $q.content.status -ForegroundColor Green
                            $CSVfilename = getFileName $outputPath
                            Write-Host "Attempting to get the audit log records for EWS usage." -ForegroundColor Cyan -NoNewline
                            # Retrieve 1000 records per request instead of the default 150
                            $r = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query "security/auditLog/queries/$($AuditQueryId)/records?`$top=1000" -AccessToken $Script:Token -Method GET -Endpoint beta
                            $r.Content.value | Select-Object -ExpandProperty AuditData | Select-Object -ExcludeProperty "@odata.type", MailboxOwnerSid, AppAccessContext | Export-Csv $CSVfilename -NoTypeInformation
                            # Check if response includes more results link
                            while ($null -ne $r.Content.'@odata.nextLink') {
                                $Query = $r.Content.'@odata.nextLink'.Substring($r.Content.'@odata.nextLink'.IndexOf("security"))
                                $r = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -AccessToken $Script:Token -Query $Query -Endpoint beta
                                $r.Content.value | Select-Object -ExpandProperty AuditData | Select-Object -ExcludeProperty "@odata.type", MailboxOwnerSid, AppAccessContext | Export-Csv $CSVfilename -NoTypeInformation
                                Write-Host "." -NoNewline
                            }
                            #Export results to CSV
                            Import-Csv $CSVfilename | Group-Object -Property AppId, ClientInfoString | Select-Object Count, @{n = 'AppId'; e = { $_.Group[0].AppId } }, @{n = 'ClientInfoString'; e = { $_.Group[0].ClientInfoString } } | Out-GridView -Title "EWS Application Results"
                        }
                        else {
                            Write-Host "Audit query did not complete successfully. Check query status for more information." -ForegroundColor Yellow
                        }
                    }
                }
                
            }
        }
    }
    "GetEwsActivity"{
        $Scope= @("Application.Read.All","Directory.Read.All")
        Get-OAuthToken -AppScope $Scope -ApiEndpoint $APIResource

        # Call function to obtain list of app registrations from Entra
        GetAzureADApplications

        # Call function to obtain list of service principals from Entra
        GetAzureAdServicePrincipals

        # Call function to obtain list of OAuth2 permission grants from Entra
        GetAzureAdOauth2PermissionGrants

        # Call function to get list of applications with EWS application permissions
        GetAppsWithApplicationPermissions

        # Call function to get list of applications with EWS delegated permissions
        GetAppsByOAuthPermissionGrant

        $Script:ApiPermissions | Export-Csv "$OutputPath\EWS-ApiPermissions-$((Get-Date).ToString("yyyyMMddhhmmss")).csv" -NoTypeInformation
        
        $Apps = $Script:ApiPermissions | Sort-Object -Property ApplicationId -Unique | Select-Object ApplicationId,ApplicationDisplayName
        $AppActivity = New-Object System.Collections.ArrayList
        Write-Host "Checking the sign-in activity for the applications..." -ForegroundColor Green
        $TotalApps = $Apps.Count
        $x = 1
        $CurrentProgress = $ProgressPreference
        $ProgressPreference = "Continue"
        $CurrentView = $PSStyle.Progress.View
        try { $PSStyle.Progress.View = "Classic" }
        catch { $CurrentView = $null }
        $OutputFile = "$OutputPath\App-SignInActivity-$((Get-Date).ToString("yyyyMMddhhmmss")).csv"
        
        foreach($App in $Apps){
            Write-Progress -Activity "Getting sign-in activity for service principals" -CurrentOperation "Processing $($App.ApplicationDisplayName)" -PercentComplete (($x / $TotalApps) * 100)
            $q = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query "reports/servicePrincipalSignInActivities?`$filter=appId eq '$($App.ApplicationId)'" -AccessToken $Script:Token -Endpoint beta
            if([string]::IsNullOrEmpty($q.Content.value.appId)){
                $SpActivity = [PSCustomObject]@{
                    AppId = $App.ApplicationId
                    Application = $App.ApplicationDisplayName
                    LastSignIn =  "None"
                }
            }
            else{
                $SpActivity = [PSCustomObject]@{
                    AppId = $q.Content.value.appId
                    Application = $App.ApplicationDisplayName
                    LastSignIn =  $q.Content.value.lastSignInActivity.LastSignInDateTime
                }
            }
            $AppActivity.Add($SpActivity) | Out-Null
            $x++
        }
        Write-Host "The following applications have sign-in activity: `n" -ForegroundColor Green
        
        $AppActivity | Sort-Object -Property LastSignIn | Format-Table -AutoSize
        $AppActivity | Export-Csv -Path $OutputFile -NoTypeInformation
        $ProgressPreference = $CurrentProgress
        try{
            $PSStyle.Progress.View = $CurrentView
        }
        catch{}
        exit
    }
}