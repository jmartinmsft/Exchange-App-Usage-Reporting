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

# Version 24.11.11.1421

param (
    [ValidateScript({ Test-Path $_ })]
    [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")]
    [string] $OutputPath,

    [ValidateSet("Global", "USGovernmentL4", "USGovernmentL5", "ChinaCloud")]
    [Parameter(Mandatory = $false)]
    [string]$AzureEnvironment = "Global",

    [Parameter(Mandatory=$false, HelpMessage="The PermissionType parameter specifies whether the app registrations uses delegated or application permissions")] [ValidateSet('Application','Delegated')]
    [string]$PermissionType,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.")] 
    [string]$OAuthClientId,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).")] 
    [string]$OAuthTenantId,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.")] 
    [string]$OAuthRedirectUri = "http://localhost:8004",
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthSecretKey parameter is the the secret for the registered application.")] 
    [SecureString]$OAuthClientSecret,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate parameter is the certificate for the registered application. Certificate auth requires MSAL libraries to be available.")] 
    [string]$OAuthCertificate = $null,
  
    [Parameter(Mandatory=$False,HelpMessage="The CertificateStore parameter specifies the certificate store where the certificate is loaded.")] [ValidateSet("CurrentUser", "LocalMachine")]
    [string]$CertificateStore = $null,

    [Parameter(Mandatory=$false)] [object]$Scope= @("AuditLog.Read.All"),
    
    [Parameter(Mandatory=$False)]
    [string]$AuditQueryId,

    [Parameter(Mandatory=$true)] [ValidateSet("NewAuditQuery", "CheckAuditQuery","GetQueryResults","CovertSidsToUpn","ListQueries","GetApplicationName")]
    [string]$Operation = $null,

    [Parameter(Mandatory=$False)]
    [datetime]$StartDate=(Get-Date).AddDays(-7),

    [Parameter(Mandatory=$False)]
    [datetime]$EndDate=(Get-Date)
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
        #$Script:GraphScope = "$($cloudService.graphApiEndpoint)/.default"
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
        #$Script:GraphScope = "$($cloudService.GraphApiEndpoint)//$($Scope)"
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
    
    $CSVfilename = ("$($OutputPath)Ews-Impersonation-Results.csv")
    if((Test-Path($CSVfilename)) -and $Operation -like "*Query*") {
        Remove-Item $CSVfilename -Confirm:$false -Force
    }
    return $CSVfilename
}


$cloudService = Get-CloudServiceEndpoint $AzureEnvironment
$azureADEndpoint = $cloudService.AzureADEndpoint
$Script:applicationInfo = @{
    "TenantID" = $OAuthTenantId
    "ClientID" = $OAuthClientId
}
$APIResource = $cloudService.GraphApiEndpoint

Get-OAuthToken -AppScope $Scope -ApiEndpoint $APIResource

switch($Operation) {
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
    
        $Body = (@{
            "@odata.type" = "#microsoft.graph.security.auditLogQuery"
            filterStartDateTime = $StartSearch
            filterEndDateTime = $EndSearch
            displayName = "Audit-query-$(($Endtime-$StartTime).Days)-Days-From-$(Get-Date -Date $StartDate -f 'yyyyMMdd')"
            recordTypeFilters = @("exchangeItem","exchangeAggregatedOperation","exchangeItemAggregated","exchangeItemGroup")
            operationsFilters = @("MailItemAccessed","UpdateCalendarDelegation","Copy","Create","SoftDelete","Move","MoveToDeletedItems","HardDelete","Send","SendAs","SendOnBehalf","Update")
        }) | ConvertTo-JSON -Depth 6

        try{
            Write-Host "Attempting to create a new audit query..." -ForegroundColor Cyan -NoNewline
            $q = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query "security/auditLog/queries" -AccessToken $Script:Token -Method POST -Body $Body -Endpoint beta # | Out-Null
            [string]$auditQuery = ($q.Response.Content | ConvertFrom-Json).id
            Write-Host "OK" -ForegroundColor Green
            $CSVfilename = getFileName $outputPath
            @{auditQueryId=$auditQuery} | Export-Csv $CSVfilename -NoTypeInformation
        }
        catch {
            Write-Host "FAILED" -ForegroundColor Red
            Write-Host "Failed to create the audit query." -ForegroundColor Red
        }
        Write-Host "New audit log query created with the id: $($auditQuery)"
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
        if($q.Content.status -eq "succeeded"){
            Write-Host $q.content.status -ForegroundColor Green
            $CSVfilename = getFileName $outputPath
            Write-Host "Attempting to get the audit log records for EWS impersonation..." -ForegroundColor Cyan
            $r = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query "security/auditLog/queries/$($AuditQueryId)/records" -AccessToken $Script:Token -Method GET -Endpoint beta
            $r.Content.value | Select-Object -ExpandProperty AuditData | Where-Object {($_.UserId -match '^S-1-[0-59]-\d{2}-\d{8,10}-\d{8,10}-\d{8,10}-[1-9]\d{3}' -and $_.LogonType -eq 1)} | Select-Object -ExcludeProperty "@odata.type", MailboxOwnerSid, AppAccessContext | Export-Csv $CSVfilename -NoTypeInformation
            while($null -ne $r.Content.'@odata.nextLink'){
                $Query = $r.Content.'@odata.nextLink'.Substring($r.Content.'@odata.nextLink'.IndexOf("security"))
                $r = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -AccessToken $Script:Token -Query $Query -Endpoint beta
                $r.Content.value | Select-Object -ExpandProperty AuditData | Where-Object {($_.UserId -match '^S-1-[0-59]-\d{2}-\d{8,10}-\d{8,10}-\d{8,10}-[1-9]\d{3}' -and $_.LogonType -eq 1)} | Select-Object -ExcludeProperty "@odata.type", MailboxOwnerSid, AppAccessContext | Export-Csv $CSVfilename -NoTypeInformation -Append -Force
            }       
        }
        else {
            Write-Host "Audit query did not complete successfully. Check query status for more information." -ForegroundColor Yellow
        }
    }
    "CovertSidsToUpn" {
        if(-not(Get-Command Get-CASMailbox -ErrorAction Ignore)){
            Write-Host "Connection to Exchange Online not found." -ForegroundColor Yellow
            if(-not(Get-InstalledModule -Name ExchangeOnlineManagement -ErrorAction Ignore)){
                Write-Host "Exchange Online Managment PowerShell module is missing." -ForegroundColor Yellow
                Write-Host "The following is a link on how to install the Exchange Online PowerShell module:" -ForegroundColor Yellow
                Write-Host "https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module `n"
                exit
            }
            Write-Host "Attempting to connect to Exchange Online." -ForegroundColor Cyan
            try {
                Connect-ExchangeOnline -ShowBanner:$false -SkipLoadingCmdletHelp
            }
            catch {
                Write-Host "Unable to connect to Exchange Online." -ForegroundColor Red
            }
        }
        $CSVfilename = getFileName $outputPath
        while(-not(Test-Path $CSVfilename)){
            $CSVfilename = Read-Host "Enter the full path to the Ews-Impersonation-Results.csv "
            if(-not($CSVfilename.EndsWith("Ews-Impersonation-Results.csv"))) {
                $CSVfilename = "$($CSVfilename)\Ews-Impersonation-Results.csv"
            }
        }
        $LogonUserSids = (Import-Csv $CSVfilename).LogonUserSid | Sort-Object -Unique
        Write-Host "Getting the UserPrincipalName value for each LogonUserSid." -ForegroundColor Cyan
        foreach($Sid in $LogonUserSids) {
            $Upn = (Get-User $Sid).UserPrincipalName
            $SidToUserName += @([PSCustomObject]@{
                Sid = $Sid
                Upn = $Upn
            })
        }
        Write-Host "Updating the CSV file by replacing the SID values with the UPN values." -ForegroundColor Cyan
        foreach($user in $SidToUserName){
            (Get-Content $CSVfilename).Replace($user.Sid, $user.Upn) | Set-Content $CSVfilename
        }
    }
    "ListQueries" {
        [string]$Query = "security/auditLog/queries"
        $q = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query $Query -AccessToken $Script:Token -Method GET -Endpoint beta
        $q.content.value
    }
    "GetApplicationName" {
        $CSVfilename = getFileName $outputPath
        while(-not(Test-Path $CSVfilename)){
            $CSVfilename = Read-Host "Enter the full path to the Ews-Impersonation-Results.csv "
            if(-not($CSVfilename.EndsWith("Ews-Impersonation-Results.csv"))) {
                $CSVfilename = "$($CSVfilename)\Ews-Impersonation-Results.csv"
            }
        }
        $Applications = (Import-Csv $CSVfilename).AppId | Sort-Object -Unique
        Write-Host "Getting the application name for each AppId." -ForegroundColor Cyan
        foreach($app in $Applications) {
            [string]$Query = "applications(appId='$($app)')"
            $q = Invoke-GraphApiRequest -GraphApiUrl $APIResource -Query $Query -AccessToken $Script:Token -Method GET -Endpoint v1.0
            Write-Host "$($app): $($q.content.displayName)"
        }
    }
}

# SIG # Begin signature block
# MIIoLwYJKoZIhvcNAQcCoIIoIDCCKBwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA2V0hDu7fcq/PQ
# RFyIEqyXrNsYeVSHWJG//rzE/NyAjKCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
# Bv9XKydyAAAAAAQEMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjQwOTEyMjAxMTE0WhcNMjUwOTExMjAxMTE0WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC0KDfaY50MDqsEGdlIzDHBd6CqIMRQWW9Af1LHDDTuFjfDsvna0nEuDSYJmNyz
# NB10jpbg0lhvkT1AzfX2TLITSXwS8D+mBzGCWMM/wTpciWBV/pbjSazbzoKvRrNo
# DV/u9omOM2Eawyo5JJJdNkM2d8qzkQ0bRuRd4HarmGunSouyb9NY7egWN5E5lUc3
# a2AROzAdHdYpObpCOdeAY2P5XqtJkk79aROpzw16wCjdSn8qMzCBzR7rvH2WVkvF
# HLIxZQET1yhPb6lRmpgBQNnzidHV2Ocxjc8wNiIDzgbDkmlx54QPfw7RwQi8p1fy
# 4byhBrTjv568x8NGv3gwb0RbAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU8huhNbETDU+ZWllL4DNMPCijEU4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMjkyMzAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjmD9IpQVvfB1QehvpC
# Ge7QeTQkKQ7j3bmDMjwSqFL4ri6ae9IFTdpywn5smmtSIyKYDn3/nHtaEn0X1NBj
# L5oP0BjAy1sqxD+uy35B+V8wv5GrxhMDJP8l2QjLtH/UglSTIhLqyt8bUAqVfyfp
# h4COMRvwwjTvChtCnUXXACuCXYHWalOoc0OU2oGN+mPJIJJxaNQc1sjBsMbGIWv3
# cmgSHkCEmrMv7yaidpePt6V+yPMik+eXw3IfZ5eNOiNgL1rZzgSJfTnvUqiaEQ0X
# dG1HbkDv9fv6CTq6m4Ty3IzLiwGSXYxRIXTxT4TYs5VxHy2uFjFXWVSL0J2ARTYL
# E4Oyl1wXDF1PX4bxg1yDMfKPHcE1Ijic5lx1KdK1SkaEJdto4hd++05J9Bf9TAmi
# u6EK6C9Oe5vRadroJCK26uCUI4zIjL/qG7mswW+qT0CW0gnR9JHkXCWNbo8ccMk1
# sJatmRoSAifbgzaYbUz8+lv+IXy5GFuAmLnNbGjacB3IMGpa+lbFgih57/fIhamq
# 5VhxgaEmn/UjWyr+cPiAFWuTVIpfsOjbEAww75wURNM1Imp9NJKye1O24EspEHmb
# DmqCUcq7NqkOKIG4PVm3hDDED/WQpzJDkvu4FrIbvyTGVU01vKsg4UfcdiZ0fQ+/
# V0hf8yrtq9CkB8iIuk5bBxuPMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGg8wghoLAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIFeXHdYMW8gCu4ML6GNzOuES
# xtst+cRDaeOj4ZNIMC2xMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQCK1AAk26Rc0tEhiu1KBn2tcN6Pea1DV1hJC0LNWWLWxFCTgWwwFhE3
# 9VOMxoPzvH0arXL89LulyGApljgkoRJbFPy0iHP39PKktZad8UPnUdsOlm83IZRI
# woMMQW/MNsq+nMGoP4fEAWNfrgL3db8Mn+siQiMUQGo8lAxzLGymTJf9/EdX4LdY
# grXh6guUB7fOvc5lWXEpwVw83hm0CI1zn8i0BOOoaZfP2Rvt1UsLXP+9qF+Hp6LK
# +nIgkwBswZaPN1Qx69YMFLvZKJLGxnqC/tSkX1yBDJ5H33TJli0/7cbs8O7dqOTQ
# imXnXDw49QmIMlo2sU6kSYLPCtassrmtoYIXlzCCF5MGCisGAQQBgjcDAwExgheD
# MIIXfwYJKoZIhvcNAQcCoIIXcDCCF2wCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIMke0wH/zQvEzp+yXQC6YYXqtZBZunpuOs2JWv4Yv9hYAgZnGo5p
# nr0YEzIwMjQxMTExMTkzMDA1LjQ4NFowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpGMDAy
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEe0wggcgMIIFCKADAgECAhMzAAAB8j4y12SscJGUAAEAAAHyMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTIzMTIwNjE4
# NDU1OFoXDTI1MDMwNTE4NDU1OFowgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpGMDAyLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBALzl88sXCmliDHBjGRIR5i9AG2dglO0oqPYUrHMf
# HR+BXpeAgiuYJaakqX0g7O858n+TqI/RGehGjkXz0B3b153MZ2VZsKPVDLHkdQc1
# jzK70SUk6Z2B6429MrhFbjC72IHn/PZJ4K5irJf+/zPo+m/b2HW201axJz8o8566
# HNIBeqQDbrkFIVPmTKTG/MHQvGjFLqhahdYrrDHXvY1ElFhwg19cOFRG9R8PvSOK
# gT3atb86CNw4rFmR9DEuXBoVKtKcazteEyun1OxSCbCzJxMQ4F0ZWZ/UcIPtY5rP
# kQRxDIhLYGlFhjCw8xsHre4eInXnyo2HVIle6gvnAYO79tlTM34HNwuP3qLELvAk
# ZAwGLFYf1375XxuXXRFh1cNmWWNEC9LqIXA3OtqG7gOthvtvwzu+/CEQvTEI69vt
# YUyyy2xxd+R0TmD41JpymGAV9yh+1Dmo8PY81WasbfwOYcOhiGCP26o8s/u+ehd/
# uPr4tbxWifXnwPRauaTsK6a5xBOIdHJ6kRpUOecDYaSImh6H+vd9KEvoIeA+hMHu
# hhT93ok6dxGKgNiqpF9XbCWkpU7xv5VgcvyGfXUlEXHqnr2YvwFG1Jnp0b8YURUT
# 59WaDFh8gJSumCHJCURMk8hMQFLXkixpS5bQa9eUtKh8Z/a3kMCgOS4oJsL7dV0+
# aVhVAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUlVuHACbq0DEEzlwfwGDT5jrihnkw
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAD1Lp47gex8HTRek6A9ptw3dBl7KKmCK
# VxBINnyDpUK/0VUfN1Kr1ekCyWNlIo1ZIKWEkTPk6jdSb+1o+ehsX7wKQB2RwtCE
# t2RKF+v3WTPL28M+s6aUIDYVD2NWEVpq3ZAzffPWn4YI/m26+KsVpRbNRZUMU6mj
# 87nMOnOg9i1OvRwWDe5dpEtPnhRDdji49heqfrC6dm1RBEyIkzPGlSW919YZS0K+
# dbd4MGKQOSLHVcT3xVxgjPb7l91y+sdV5RqsZfLgtG3DObCmwK1SHu1HrCEKtViR
# voW50F1YztNW+OLukaB+N6yCcBJoP8KEu7Hro8bBohoX7EvOTRs3GwCPS6F3pB1a
# vpNPf2b9I1nX9RdTuTMSh3S8BjeYifxfkDgj7397WcE2lREnpiIMpB3lhWDGy5kJ
# a/hDBvSZeEch70K5t9KpmO8NrB/Yjbb03cuy0MlRKvW8YUHyJDlbxkszk/BPy+2w
# oQHAcRibCy5aazGSKYgXkFBtLOD3DPU7qN1ZPEYbQ5S3VxdY4wlQnPIQfhZIpkc7
# HnepwC8P2HRTqMQXZ+4GO0n9AOtZtvi6u8B+u+o2f2UfuBU+mWo08Mi9DwORneW9
# tCxiqXPrXt7vqBrtJjTDvX5A/XrkI93NRjfp63ZKbim+ykQryGWWrchhzJfS/z3v
# 5f1h55wzU9vWMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgzMjI1WjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxOdcjKNVf2AX9sSuDivbk+
# F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQGOhgfWpSg0S3po5GawcU
# 88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq/XJprx2rrPY2vjUmZNqY
# O7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVWTe/dvI2k45GPsjksUZzp
# cGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7mka97aSueik3rMvrg0Xn
# Rm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De+JKRHh09/SDPc31BmkZ1
# zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM9ituqBJR6L8FA6PRc6ZN
# N3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEzOUyOArxCaC4Q6oRRRuLR
# vWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2ZItboKaDIV1fMHSRlJTY
# uVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqvUAV6bMURHXLvjflSxIUX
# k8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q4i6tAgMBAAGjggHdMIIB
# 2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQWBBQqp1L+ZMSavoKR
# PEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXAYDVR0g
# BFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRvcnkuaHRtMBMGA1UdJQQM
# MAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQE
# AwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQ
# W9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNv
# bS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBa
# BggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0MA0GCSqG
# SIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8qW/qXBS2Pk5HZHixBpOX
# PTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6U03dmLq2HnjYNi6c
# qYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2CDPVtI1TkeFN1JFe53Z/z
# jj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BAljis9/kpicO8F7BUhUKz
# /AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJeBTpkbKpW99Jo3QMvOyR
# gNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1MMU0sHrYUP4KWN1APMdU
# bZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz138eW0QBjloZkWsNn6Qo
# 3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1V1QJsWkBRH58oWFsc/4K
# u+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLBgqJ7Fx0ViY1w/ue10Cga
# iQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0llOZ0dFtq0Z4+7X6gMTN9
# vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFxBmoQtB1VM1izoXBm8qGC
# A1AwggI4AgEBMIH5oYHRpIHOMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046RjAwMi0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# AGuL3jdwUsfZN9AR8HTlIsgKDvgIoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDq3MaRMCIYDzIwMjQxMTExMTgw
# ODE3WhgPMjAyNDExMTIxODA4MTdaMHcwPQYKKwYBBAGEWQoEATEvMC0wCgIFAOrc
# xpECAQAwCgIBAAICBjQCAf8wBwIBAAICE+AwCgIFAOreGBECAQAwNgYKKwYBBAGE
# WQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDAN
# BgkqhkiG9w0BAQsFAAOCAQEARCqSbJ6jHMjhdsRisR7AyZIxZmHebAgsfvC2cXfu
# dzJwStXuqqtmblGHMYKB+crP8leN9QoSfgAe/Ngsdc/bM+eGXAlw7FEAvblAtpCm
# NgLMRaWVUSco6GmkpQl6l5ZpZczfXMi3tYnrFX1wRhCjVUxd82mCplI6O0NTltbu
# PF+3WPP72hZse8jpbYDr/l4OqUHipKFvTBiNc23ZoHyjAVm2c10yECvjQrzWycu6
# 0hlhv7AdJs+nN0Rsxqfb9CHB7tz9lalVh/cocInhzEYjOkPR8c7gYbJelInjw6v+
# CdKoAarVvUziNTnt+Cz6bcBl6cz0wzMZyV2Y5QlQHJYLPDGCBA0wggQJAgEBMIGT
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB8j4y12SscJGUAAEA
# AAHyMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQ
# AQQwLwYJKoZIhvcNAQkEMSIEIBv/HPnNyj0LHRFB00rONTdoipZeMrhmSSPkZ6so
# puMtMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQg+No+HS4xUlzTj5jhG7kF
# RRscTiy5nqdEdJS7RddKQ0QwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMAITMwAAAfI+MtdkrHCRlAABAAAB8jAiBCAEDMgbC5wtuzley2CR/ELY
# 7G9fLTKlU5cSjreZOr+aKDANBgkqhkiG9w0BAQsFAASCAgAiGKJr7noteJP5q0o1
# Z4tSyk1pTRI1vPaRluwwkK4wsXDotpCK/+HfQlNtGZ+ScTow6DYLqKEgC8QaSHPg
# S9aOb1oKq7B1Np96+XWesDZqtH9wIhH2Dk6WW0ZcyboFZhDw+z74DUoZkcDEapQ9
# E/y2Tp8Mnpzr5myse+gcM/hePM3aT36IUaK6w0Tw3asoENFQ3hcpo4n/1PyzZFNA
# BYX8hj67U3GYcX/T4zwWSJK/Px6IW3+0j0RSvzKS2cIWurlt1NumuMLOTKfEap7H
# 9jvaOqrNCBKDDIBddp1sSPYrPFVFOhI1ANAlLA92RFPPi807b2imLzdk7p29Ml43
# beqBUHtjIKwp5X/4/W5+CYOmsar8ibAUtciyfFwuKco/+CQ+r0D/M5Tj4i3kPPN0
# eo2BH151KI/UD115Usc4WLQuGYA0ItKSJ0pTD2q8GBfTKn2QAlsg6NI21r1aNBcC
# AlxY1/2eogqeir4mi9uwt2r/zXzDXyGqIpzdczdAxMZBKoMEKTUww6KTAs62RGfH
# VWXtpiVF/bHm2Soq14C6HG/ZB6CIdyEBsjQ2qtXGae1Wh7pud2/OUrggtgGt33fy
# Jj+PVVxsLXHyhdmM54ZS7tlhhE4w/h8mXmO+2nI9F8VSvmHq+8W2vJpdaiHkDSWU
# hkjqSCoLRUQYavTmpLKXXjYJkg==
# SIG # End signature block
