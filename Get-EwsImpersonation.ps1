param(
    [Parameter(Mandatory=$true, Position=1, HelpMessage="The Api parameter specifies which API Permisions to export for esach Application registration")] [ValidateSet('OutlookRESTv2','EWS')] [string]$Api,
    [Parameter(Mandatory=$false, HelpMessage="The AuthMethod parameter specifies whether to use delegate or application permissions")] [ValidateSet('Application','Delegate')] [string]$AuthMethod,
    [Parameter(Mandatory=$false, HelpMessage="The OAuthClientId parameter specifies the the app ID for the OAuth token request.")] [string] $OAuthClientId = "5bd878c3-a3f8-4996-96c9-09c5d7965e72",
    [Parameter(Mandatory=$false, HelpMessage="The OAuthClientSecret parameter specifies the the app secret for the OAuth token request.")] [securestring] $OAuthClientSecret, #="xgv8Q~7l.1mJfkMc1Qol79xumsSzy3yydgCvcclk",
    [Parameter(Mandatory=$false, HelpMessage="The OAuthTenantId parameter specifies the the tenant ID for the OAuth token request.")] [string] $OAuthTenantId = "9101fc97-5be5-4438-a1d7-83e051e52057",
    [Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri specifies the redirect Uri of the Azure registered application.")] [string] $OAuthRedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",
    [Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate parameter is the certificate for the registerd application. Certificate auth requires MSAL libraries to be available.")] $OAuthCertificate = $null,
    #[ValidateScript({ Test-Path -Path $_ -PathType Container })] [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")] [ValidatePattern("^^[a-zA-Z]:(\\w+)*")] [string] $OutputPath,
    [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")] [string] $OutputPath,
    [Parameter(Mandatory=$false, HelpMessage="The NumberOfDays parameter specifies how many days of sign-in logs to query (default is three).")] [int] $NumberOfDays=1,
    [Parameter(Mandatory=$False,HelpMessage="The LogFile parameter specifies the Log file path - activity is logged to this file if specified.")][string]$LogFile = "",
    [Parameter(Mandatory=$False,HelpMessage="The VerboseLogFile parameter is a switch that enables verbose log file.  Verbose logging is written to the log whether -Verbose is enabled or not.")]	[switch]$VerboseLogFile,
    [Parameter(Mandatory=$False,HelpMessage="The DebugLogFile parameter is a switch that enables debug log file.  Debug logging is written to the log whether -Debug is enabled or not.")][switch]$DebugLogFile,
    [Parameter(Mandatory=$False,HelpMessage="The FastFileLogging parameter is a switch that if selected, an optimised log file creator is used that should be signficantly faster (but may leave file lock applied if script is cancelled).")][switch]$FastFileLogging,
    [Parameter(Mandatory=$False,HelpMessage="The ImpersonationCheck parameter is a switch that enables checking accounts with EWS impersonation rights.")][switch]$ImpersonationCheck
)

#>** LOGGING FUNCTIONS START **#
Function LogToFile([string]$Details) {
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
}

Function UpdateDetailsWithCallingMethod([string]$Details) {
    # Update the log message with details of the function that logged it
    $timeInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())"
    $callingFunction = (Get-PSCallStack)[2].Command # The function we are interested in will always be frame 2 on the stack
    if (![String]::IsNullOrEmpty($callingFunction))
    {
        return "$timeInfo [$callingFunction] $Details"
    }
    return "$timeInfo $Details"
}

Function LogToFile([string]$logInfo) {
    if ( [String]::IsNullOrEmpty($LogFile) ) { return }

    if ($FastFileLogging)
    {
        # Writing the log file using a FileStream (that we keep open) is significantly faster than using out-file (which opens, writes, then closes the file each time it is called)
        $fastFileLogError = $Error[0]
        if (!$script:logFileStream)
        {
            # Open a filestream to write to our log
            Write-Verbose "Opening/creating log file: $LogFile"
            $script:logFileStream = New-Object IO.FileStream($LogFile, ([System.IO.FileMode]::Append), ([IO.FileAccess]::Write), ([IO.FileShare]::Read) )
            if ( $Error[0] -ne $fastFileLogError )
            {
                $FastFileLogging = $false
                Write-Host "Fast file logging disabled due to error: $Error[0]" -ForegroundColor Red
                $script:logFileStream = $null
            }
        }
        if ($script:logFileStream)
        {
            if (!$script:logFileStreamWriter)
            {
                $script:logFileStreamWriter = New-Object System.IO.StreamWriter($script:logFileStream)
            }
            $script:logFileStreamWriter.WriteLine($logInfo)
            $script:logFileStreamWriter.Flush()
            if ( $Error[0] -ne $fastFileLogError )
            {
                $FastFileLogging = $false
                Write-Host "Fast file logging disabled due to error: $Error[0]" -ForegroundColor Red
            }
            else
            {
                return
            }
        }
    }

	$logInfo | Out-File $LogFile -Append
}

Function Log([string]$Details, [ConsoleColor]$Colour) {
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Host $Details -ForegroundColor $Colour
    LogToFile $Details
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting" Green

Function LogVerbose([string]$Details) {
    Write-Verbose $Details
    #if ( !$VerboseLogFile -and !$DebugLogFile -and ($VerbosePreference -eq "SilentlyContinue") ) { return }
    LogToFile $Details
}

Function LogDebug([string]$Details) {
    Write-Debug $Details
    if (!$DebugLogFile -and ($DebugPreference -eq "SilentlyContinue") ) { return }
    LogToFile $Details
}

$script:LastError = $Error[0]
Function ErrorReported($Context) {
    # Check for any error, and return the result ($true means a new error has been detected)

    # We check for errors using $Error variable, as try...catch isn't reliable when remoting
    if ([String]::IsNullOrEmpty($Error[0])) { return $false }

    # We have an error, have we already reported it?
    if ($Error[0] -eq $script:LastError) { return $false }

    # New error, so log it and return $true
    $script:LastError = $Error[0]
    if ($Context)
    {
        Log "ERROR ($Context): $($Error[0])" Red
    }
    else
    {
        $log = UpdateDetailsWithCallingMethod("ERROR: $($Error[0])")
        Log $log Red
    }
    return $true
}

Function ReportError($Context) {
    # Reports error without returning the result
    ErrorReported $Context | Out-Null
}
#>** LOGGING FUNCTIONS END **#

function GetOAuthToken {
    Log "Obtaining access token to use for the Graph API calls." Cyan
    # Check if using a certificate for authentication
    if($null -notlike $OAuthCertificate) {
        try {
            $Cert = Get-Item Cert:\CurrentUser\My\$OAuthCertificate -ErrorAction Ignore
        }
        catch {
            $cert = Get-Item Cert:\LocalMachine\My\$OAuthCertificate
        }
        $Script:OAuthToken = Get-MsalToken -ClientId $OAuthClientId -RedirectUri $RedirectUri -TenantId $OAuthTenantId -Scopes $Script:Scope -AzureCloudInstance AzurePublic -ClientCertificate $Cert
    }
    else {
        # Determine if using delegate or application permissions
        if($AuthMethod -eq "Application") {
            $Script:OAuthToken = Get-MsalToken -ClientId $OAuthClientId -ClientSecret $OAuthClientSecret -TenantId $OAuthTenantId -Scopes $Script:Scope #-AzureCloudInstance AzurePublic
        }
        else {
            $Script:OAuthToken = Get-MsalToken -ClientId $OAuthClientId -TenantId $OAuthTenantId -Scopes $Script:Scope -Interactive -RedirectUri $OAuthRedirectUri
        }
    }
    return $Script:OAuthToken.AccessToken
}

function TestInstalledModules{
    #Test if the required module is installed, if not exit the script and print a help message
    if(-not (Get-InstalledModule -Name MSAL.PS -MinimumVersion 4.37.0.0)) {
        Write-Host "This script requires 'MSAL.PS' module with minimum version 4.37.0.0" -ForegroundColor Red
        Write-Host "Please install the required 'MSAL.PS' module from the PSGallery repository by running command:" -ForegroundColor Red
        Write-Host "Install-Module -Name MSAL.PS -MinimumVersion '4.37.0.0' -Repository:PSGallery" -ForegroundColor Red
        exit
    }
    if($ImpersonationCheck){
        if(-not(Get-InstalledModule -Name ExchangeOnlineManagement -MinimumVersion 3.4.0 -ErrorAction Ignore)){
            Log "This script requires 'ExchangeOnlineManagement' module with minimum version 3.4.0" Red
            Write-Host "Please install the required 'ExchangeOnlineManagement' module from the PSGallery repository by running command" -ForegroundColor Red
            Write-Host "Install-Module -Name ExchangeOnlineManagement -MinimumVersion 3.4.0" -ForegroundColor Red
            exit
        }
    }
}

function SendGraphRequest{
    param(
        [Parameter(Mandatory=$true, HelpMessage="The Uri parameter specifies the request uri.")] [string] $Uri,
        [Parameter(Mandatory=$false, HelpMessage="The HttpMethod parameter specifies the method for the request.")] [string] $HttpMethod="GET"
    )
    # Verify token hasn't expired before sending request
    if($Script:OAuthToken.ExpiresOn -lt (Get-Date)) {
        Write-Host "Acquiring new token..." -ForegroundColor Cyan
        $Script:Token = GetOAuthToken
        #Update the headers with the new auth token
        $Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Script:Token"
        }
    }
    else {
        $Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Script:Token"
        }
    }
    $MessageParams = @{
        "URI"         = $Uri
        "Headers"     = $Headers
        "Method"      = $HttpMethod
        "ContentType" = "application/json"
        "UseBasicParsing" = $null
        }

        $Results = ""
        $StatusCode = ""
        # Send the request and look for either throttling or timeout errors
        do {
            try {
                $Results = Invoke-RestMethod @Messageparams
                $StatusCode = $Results.StatusCode
            }
            catch {
                $StatusCode = $_.Exception.Response.StatusCode.value__
                if ($StatusCode -eq 429) {
                    Write-Warning "Request being throttled. Sleeping for 50 seconds..."
                    Start-Sleep -Seconds 50
                }
                elseif ($StatusCode -eq 504) {
                    Write-Warning "Request received timeout error. Retrying in 20 seconds..."
                    Start-Sleep -Seconds 20
                }
                else {
                    Write-Error $_.Exception
                }
            }
        }
        while ($StatusCode -eq 429)
        return $Results

}

function GetAzureAdApplications{
    Log "Getting a list of Azure AD Applications" Green
    $Script:AadApplications = New-Object System.Collections.ArrayList
    $AadApplicationResults = SendGraphRequest -Uri "https://graph.microsoft.com/v1.0/applications?`$select=appId,createdDateTime,displayName,description,notes,requiredResourceAccess"
    foreach($application in $AadApplicationResults.Value){
        $Script:AadApplications.Add($application) | Out-Null
    }
    # Check if response includes more results link
    while($AadApplicationResults.'@odata.nextLink' -ne $null){
        $AadApplicationResults = SendGraphRequest -Uri $AadApplicationResults.'@odata.nextLink'
        foreach($application in $AadApplicationResults.Value){
            $Script:AadApplications.Add($application) | Out-Null
        }
    }
}

function GetAzureAdServicePrincipals{
    Log "Getting a list of allAzure AD service applications" Green
    $Script:ServicePrincipals = New-Object System.Collections.ArrayList
    $ServicePrincipalsResults = SendGraphRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=id,appDisplayName,appDescription,appId,createdDateTime,displayName,servicePrincipalType,appRoles,oauth2PermissionScopes"
    foreach($ServicePrincipal in $ServicePrincipalsResults.Value){
        $Script:ServicePrincipals.Add($ServicePrincipal) | Out-Null
    }
    # Check if response includes more results link
    while($ServicePrincipalsResults.'@odata.nextLink' -ne $null){
        $ServicePrincipalsResults = SendGraphRequest -Uri $ServicePrincipalsResults.'@odata.nextLink'
        foreach($ServicePrincipal in $ServicePrincipalsResults.Value){
            $Script:ServicePrincipals.Add($ServicePrincipal) | Out-Null
        }
    }
}

function GetAppsByApi {
    $Script:ApiPermissions = New-Object System.Collections.ArrayList
foreach($application in $Script:AadApplications) {
    foreach($RequiredResourceAccess in $application.requiredResourceAccess) {
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
                        $Script:AppPermissions = [PSCustomObject]@{
                                'ApplicationDisplayName'  = $application.displayName
                                'ApplicationID'           = $application.appId
                                'PermissionType'          = "Delegate"
                                'PermissionValue'         = $AppScope.Value
                                "ResourceDisplayName"     = $RequiredResourceAccess.ResourceAppId
                            }
                            switch ($Api) {
                                'EWS' {
                                    if($appScope.value -in $Script:EWSPermissions) {
                                        $Script:ApiPermissions.Add($Script:AppPermissions) | Out-Null
                                    }
                                }
                                'OutlookRESTv2' {
                                    if($appScope.value -in $Script:Delegated_OutlookRESTPermissions -and $ExchangeOnlineAccess) {
                                        $Script:ApiPermissions.Add($Script:AppPermissions) | Out-Null
                                    }
                                    elseif($appScope.value -in $Script:EWSPermissions -and $GraphAccess) {
                                        $Script:ApiPermissions.Add($Script:AppPermissions) | Out-Null
                                    }
                                }
                            }
                        
                    }
                    elseif ($($ResourceAccess.Type) -eq "Role") {
                        $AppRole = $sp.appRoles | Where-Object {$_.id -eq "$($ResourceAccess.id)"}
                        $Script:AppPermissions = [PSCustomObject]@{
                            'ApplicationDisplayName'  = $application.displayName
                            'ApplicationID'           = $application.appId
                            'PermissionType'          = "Application"
                            'PermissionValue'         = $AppRole.Value
                            "ResourceDisplayName"     = $RequiredResourceAccess.ResourceAppId
                        }
                        switch ($Api) {
                            'EWS' {
                                if($appRole.value -in $Script:EWSPermissions -and $ExchangeOnlineAccess) {
                                    $Script:ApiPermissions.Add($Script:AppPermissions) | Out-Null
                                }
                            }
                            'OutlookRESTv2' {
                                if($appRole.value -in $Script:Application_OutlookRESTPermissions -and $ExchangeOnlineAccess) {
                                    $Script:ApiPermissions.Add($Script:AppPermissions) | Out-Null
                                }
                            }
                        }
                    }
                    
                }
                
            }
    }
}

}

function GetEwsSignIns{
    $ApplicationPermissions = $Script:ApiPermissions | Sort-Object -Property ApplicationId -Unique
    $NumberOfApps = $ApplicationPermissions.Count
    $AppsCompleted = 0
    $StartDate = (Get-Date).AddDays(-$NumberOfDays)
    $TempDate = [datetime]$StartDate
    $TempDate = $TempDate.ToUniversalTime()
    $SearchStartDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempDate

    Log "Searching for EWS sign-in attempts." Green
    foreach($App in $ApplicationPermissions) {
        Write-Progress -Activity "Searching for EWS sign-in attempts" -Status "Checking $($App.ApplicationDisplayName)" -PercentComplete ((($AppsCompleted)/$NumberOfApps)*100)
        $SignIns = SendGraphRequest -Uri "https://graph.microsoft.com/beta/auditLogs/signIns?`$filter=appid eq '$($App.ApplicationId)' and signInEventTypes/any(t: t eq 'interactiveUser' or t eq 'nonInteractiveUser' or t eq 'servicePrincipal' or t eq 'managedIdentity') and CreatedDateTime ge $SearchStartDate"
        $SignIns.value | Select-Object id, createdDateTime, appId, appDisplayName, correlationId, clientCredentialType, resourceDisplayName, resourceId, servicePrincipalId , userDisplayName, userPrincipalName, @{Name='SignInEventTypes';Expression={$_.signInEventTypes -join '; ' } } | Export-Csv -Path $Script:FileName -NoTypeInformation -NoClobber -Append
        $AppsCompleted++
    }
}


#Define variables
$Date = [DateTime]::Now
$Script:StartTime = '{0:MM/dd/yyyy HH:mm:ss}' -f $Date
$Script:FileName = "$OutputPath\Ews-Applications-$('{0:MMddyyyyHHmms}' -f $Date).csv"
$Script:Scope = "https://graph.microsoft.com/.default"
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
TestInstalledModules

# Call function to obtain OAuth token
$Script:Token = GetOAuthToken

# Call function to obtain list of app registrations from Entra
GetAzureADApplications
$Script:AadApplications | Export-Csv "$OutputPath\AppRegistrations-$('{0:MMddyyyyHHmms}' -f $Date).csv" -NoTypeInformation

# Call function to obtain list of service principals from Entra
GetAzureAdServicePrincipals
$Script:ServicePrincipals | Export-Csv "$OutputPath\AadServicePrincipals-$('{0:MMddyyyyHHmms}' -f $Date).csv" -NoTypeInformation

# Call function to Filter app registrations using the selected API
GetAppsByApi
$Script:ApiPermissions | Format-Table -AutoSize

# Call function to obtain sign-in logs for app registrations using the selected API
GetEwsSignIns
Log "Script complete" Green

# Find users with EWS impersonation rights
if($ImpersonationCheck){
    Log "Checking for users with the ApplicationImpersonation role" Green

}

<#


#>
# SIG # Begin signature block
# MIInwQYJKoZIhvcNAQcCoIInsjCCJ64CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDjZI5RNTZ2UnXu
# By+9heCYQ/dRI9m87Wh6zwEf73vSKaCCDXYwggX0MIID3KADAgECAhMzAAADrzBA
# DkyjTQVBAAAAAAOvMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMxMTE2MTkwOTAwWhcNMjQxMTE0MTkwOTAwWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDOS8s1ra6f0YGtg0OhEaQa/t3Q+q1MEHhWJhqQVuO5amYXQpy8MDPNoJYk+FWA
# hePP5LxwcSge5aen+f5Q6WNPd6EDxGzotvVpNi5ve0H97S3F7C/axDfKxyNh21MG
# 0W8Sb0vxi/vorcLHOL9i+t2D6yvvDzLlEefUCbQV/zGCBjXGlYJcUj6RAzXyeNAN
# xSpKXAGd7Fh+ocGHPPphcD9LQTOJgG7Y7aYztHqBLJiQQ4eAgZNU4ac6+8LnEGAL
# go1ydC5BJEuJQjYKbNTy959HrKSu7LO3Ws0w8jw6pYdC1IMpdTkk2puTgY2PDNzB
# tLM4evG7FYer3WX+8t1UMYNTAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQURxxxNPIEPGSO8kqz+bgCAQWGXsEw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMTgyNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAISxFt/zR2frTFPB45Yd
# mhZpB2nNJoOoi+qlgcTlnO4QwlYN1w/vYwbDy/oFJolD5r6FMJd0RGcgEM8q9TgQ
# 2OC7gQEmhweVJ7yuKJlQBH7P7Pg5RiqgV3cSonJ+OM4kFHbP3gPLiyzssSQdRuPY
# 1mIWoGg9i7Y4ZC8ST7WhpSyc0pns2XsUe1XsIjaUcGu7zd7gg97eCUiLRdVklPmp
# XobH9CEAWakRUGNICYN2AgjhRTC4j3KJfqMkU04R6Toyh4/Toswm1uoDcGr5laYn
# TfcX3u5WnJqJLhuPe8Uj9kGAOcyo0O1mNwDa+LhFEzB6CB32+wfJMumfr6degvLT
# e8x55urQLeTjimBQgS49BSUkhFN7ois3cZyNpnrMca5AZaC7pLI72vuqSsSlLalG
# OcZmPHZGYJqZ0BacN274OZ80Q8B11iNokns9Od348bMb5Z4fihxaBWebl8kWEi2O
# PvQImOAeq3nt7UWJBzJYLAGEpfasaA3ZQgIcEXdD+uwo6ymMzDY6UamFOfYqYWXk
# ntxDGu7ngD2ugKUuccYKJJRiiz+LAUcj90BVcSHRLQop9N8zoALr/1sJuwPrVAtx
# HNEgSW+AKBqIxYWM4Ev32l6agSUAezLMbq5f3d8x9qzT031jMDT+sUAoCw0M5wVt
# CUQcqINPuYjbS1WgJyZIiEkBMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGaEwghmdAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAOvMEAOTKNNBUEAAAAAA68wDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIIQshildIvrTUD4A9pAgyuOr
# LPBQlqcyRmYa1KufwGXSMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQAchr1aoSSZS1HBqVq2+BpnmeczHt0me5e2jJXxVstvJBdt+lm97eR7
# O9/5oAs7Gkkekduk4ghbJaN7GCZ/3eNPTx+iQuIX/dEP1ysQUaJ0i/i5iRVFFnPa
# GQMd543eszpP/dCpeZ4Bngbx/KBWER8rTNsdnlJei/iumgGJCycJfguuEfGDRUSW
# BzYsfeUGU+Q30RbegZ9hK3/5di/Vqq1N0qjnsUp3CDH/bxN3J7VZs1fVJ2Q+DU2h
# 8yAvStoyiwWCRkr/JKcO7niJzugFI7+pPaEUuW6agsVa1rF9bMUUTeDBLiwpH7wX
# UeR0B3cE+8xZHoullgV+DyDvoAn17XwCoYIXKTCCFyUGCisGAQQBgjcDAwExghcV
# MIIXEQYJKoZIhvcNAQcCoIIXAjCCFv4CAQMxDzANBglghkgBZQMEAgEFADCCAVkG
# CyqGSIb3DQEJEAEEoIIBSASCAUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIFCS0otFdtMtgCRMsFNbD9TbylIy/uYUOgVVGb1NeUofAgZl1eZa
# t/EYEzIwMjQwMzA3MTQyNTExLjk2NlowBIACAfSggdikgdUwgdIxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBF
# U046MDg0Mi00QkU2LUMyOUExJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2WgghF4MIIHJzCCBQ+gAwIBAgITMwAAAdqO1claANERsQABAAAB2jAN
# BgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0y
# MzEwMTIxOTA2NTlaFw0yNTAxMTAxOTA2NTlaMIHSMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBP
# cGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjA4NDIt
# NEJFNi1DMjlBMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAk5AGCHa1UVHWPyNADg0N
# /xtxWtdI3TzQI0o9JCjtLnuwKc9TQUoXjvDYvqoe3CbgScKUXZyu5cWn+Xs+kxCD
# bkTtfzEOa/GvwEETqIBIA8J+tN5u68CxlZwliHLumuAK4F/s6J1emCxbXLynpWzu
# wPZq6n/S695jF5eUq2w+MwKmUeSTRtr4eAuGjQnrwp2OLcMzYrn3AfL3Gu2xgr5f
# 16tsMZnaaZffvrlpLlDv+6APExWDPKPzTImfpQueScP2LiRRDFWGpXV1z8MXpQF6
# 7N+6SQx53u2vNQRkxHKVruqG/BR5CWDMJCGlmPP7OxCCleU9zO8Z3SKqvuUALB9U
# aiDmmUjN0TG+3VMDwmZ5/zX1pMrAfUhUQjBgsDq69LyRF0DpHG8xxv/+6U2Mi4Zx
# 7LKQwBcTKdWssb1W8rit+sKwYvePfQuaJ26D6jCtwKNBqBiasaTWEHKReKWj1gHx
# DLLlDUqEa4frlXfMXLxrSTBsoFGzxVHge2g9jD3PUN1wl9kE7Z2HNffIAyKkIabp
# Ka+a9q9GxeHLzTmOICkPI36zT9vuizbPyJFYYmToz265Pbj3eAVX/0ksaDlgkkIl
# cj7LGQ785edkmy4a3T7NYt0dLhchcEbXug+7kqwV9FMdESWhHZ0jobBprEjIPJId
# g628jJ2Vru7iV+d8KNj+opMCAwEAAaOCAUkwggFFMB0GA1UdDgQWBBShfI3JUT1m
# E5WLMRRXCE2Avw9fRTAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBf
# BgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmww
# bAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0El
# MjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUF
# BwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAuYNV1O24jSMA
# S3jU7Y4zwJTbftMYzKGsavsXMoIQVpfG2iqT8g5tCuKrVxodWHa/K5DbifPdN04G
# /utyz+qc+M7GdcUvJk95pYuw24BFWZRWLJVheNdgHkPDNpZmBJxjwYovvIaPJauH
# vxYlSCHusTX7lUPmHT/quz10FGoDMj1+FnPuymyO3y+fHnRYTFsFJIfut9psd6d2
# l6ptOZb9F9xpP4YUixP6DZ6PvBEoir9CGeygXyakU08dXWr9Yr+sX8KGi+SEkwO+
# Wq0RNaL3saiU5IpqZkL1tiBw8p/Pbx53blYnLXRW1D0/n4L/Z058NrPVGZ45vbsp
# t6CFrRJ89yuJN85FW+o8NJref03t2FNjv7j0jx6+hp32F1nwJ8g49+3C3fFNfZGE
# xkkJWgWVpsdy99vzitoUzpzPkRiT7HVpUSJe2ArpHTGfXCMxcd/QBaVKOpGTO9Kd
# ErMWxnASXvhVqGUpWEj4KL1FP37oZzTFbMnvNAhQUTcmKLHn7sovwCsd8Fj1QUvP
# iydugntCKncgANuRThkvSJDyPwjGtrtpJh9OhR5+Zy3d0zr19/gR6HYqH02wqKKm
# Hnz0Cn/FLWMRKWt+Mv+D9luhpLl31rZ8Dn3ya5sO8sPnHk8/fvvTS+b9j48iGanZ
# 9O+5Layd15kGbJOpxQ0dE2YKT6eNXecwggdxMIIFWaADAgECAhMzAAAAFcXna54C
# m0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZp
# Y2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMy
# MjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51
# yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY
# 6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9
# cmmvHaus9ja+NSZk2pg7uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN
# 7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDua
# Rr3tpK56KTesy+uDRedGbsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74
# kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2
# K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5
# TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZk
# i1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9Q
# BXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3Pmri
# Lq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUC
# BBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJl
# pxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9y
# eS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUA
# YgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIw
# MTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/yp
# b+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulm
# ZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM
# 9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECW
# OKz3+SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4
# FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3Uw
# xTSwethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPX
# fx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVX
# VAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGC
# onsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU
# 5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEG
# ahC0HVUzWLOhcGbyoYIC1DCCAj0CAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNO
# OjA4NDItNEJFNi1DMjlBMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQBCoh8hiWMdRs2hjT/COFdGf+xIDaCBgzCB
# gKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBBQUA
# AgUA6ZQozzAiGA8yMDI0MDMwNzE5NTE0M1oYDzIwMjQwMzA4MTk1MTQzWjB0MDoG
# CisGAQQBhFkKBAExLDAqMAoCBQDplCjPAgEAMAcCAQACAh1aMAcCAQACAhE1MAoC
# BQDplXpPAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
# AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEFBQADgYEAG7QR3EaFFEXhRLjq
# H4W7xBITC6+S7LjP4C/A+4+NNTioqXN/zfnmgUUEuFnOYufhnb8fcfteQ79O4KAb
# nzmgsdkXXGbmUoZY4MyrySvo5U3foiYYjGWZLLpeY1+f9AZUqWQIxMzb997MhIWe
# UXYWvtB0KZbvNa9zg0p4ewRAD0IxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdqO1claANERsQABAAAB2jANBglghkgBZQME
# AgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJ
# BDEiBCAYFhNqNPIxI1tt1Wp9S9vVd5npCm2O+bBZrXZp/F5N7jCB+gYLKoZIhvcN
# AQkQAi8xgeowgecwgeQwgb0EICKlo2liwO+epN73kOPULT3TbQjmWOJutb+d0gI7
# GD3GMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHa
# jtXJWgDREbEAAQAAAdowIgQgdnUxQSBSNsbJUeSJ0ChO6pBPg/CUJQuS/8i7hz31
# REEwDQYJKoZIhvcNAQELBQAEggIAWdU/f7vR5W+m7g6ITvb+/M6gKl0H61aUmgbn
# jOxAX7nbW0/hvoMlSGBc3q0ENg9i01niZTFMKUydSL/8F80NIyU0gD2cNTOqj+e4
# z5fp3rdJioh2wkQ2WkrnI/Xlc6WJleseR+rfdGphIZBd4CoyZGbsFXxO7tkoQxYY
# B+hypLZlgDn2I2AJK42jHKPzP4PqCztCWtUB4Syw1CzQZC6+PBo/M7jtxFSA5SbS
# 1m4ZyB/JHpdUn/VEmIDTuakZ7AYJzxOvWzRAYmvp1WH7CEcaQemvVWKpIgSyWFS8
# HkXkM61ggvo0Fk5YwiAXiZqLu6tTqGy42N+T5WOH9rSbA6kB76QPZXcM7GZi7LWu
# IjrDCLNdvGc2Np8gFaYBnYfdNNTk++Vk/RFb/gPvcLH9x5wrpof08EPfbc1HgD9r
# aupfFbXnNf6Rde3KEXnESBZbLBvWhVpxIFdePE4eMfGKnc5okzZ38erKxsPsQpfg
# FSbRogouwjvvPgf6fBFYosedEOiKqnBWgosUq/okZQb9AK/UExLPVTGpUskfO95e
# gaAl+4Img+MfG9GhcSXQfR3KYRhoEZrEtUri3jjOCQx3m0DVOir6vbH7kIZNahxb
# W0EFlP4Xpo1BtPnwT+Oq0FQVZ/EVDZqPfF6XnEZu1IIh8Golezt0sVQsp/2ssBJa
# IWdTFxY=
# SIG # End signature block
