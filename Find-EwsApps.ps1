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

# Version 20250418.0953
param (
    [ValidateScript({ Test-Path $_ })]
    [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the output data.")]
    [string] $OutputPath,

    [ValidateScript({ Test-Path $_ })]
    [Parameter(Mandatory = $true, HelpMessage="The EwsUsageReportPath parameter specifies the path for the EWS usage report.")]
    [string] $EwsUsageReportPath="C:\Temp\Output\EWSWeeklyUsage_4_11_2025_12_10_14.csv"
)

#Download the latest CSV of first party apps
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/merill/microsoft-info/main/_info/MicrosoftApps.csv" -OutFile "$($OutputPath)\MicrosoftApps.csv"
$MicrosoftApps = Import-Csv -Path "$($OutputPath)\MicrosoftApps.csv"
#Create a hashtable of first party apps
$AppsToCheck = @{}
foreach($app in $MicrosoftApps) {$AppsToCheck[$app.AppId] = $app.appDisplayName}
#Obtain list of EWS apps in tenant from usage report
$EwsApps = Import-Csv $EwsUsageReportPath
$EwsApps = $EwsApps | Sort-Object appid -Unique
#Get app name for all non-third-party apps
$Applications = New-Object System.Collections.ArrayList
$AppIdsNotFound = New-Object System.Collections.ArrayList
foreach($app in $EwsApps) {
    if(-not($AppsToCheck.ContainsKey($app.AppId))) { 
        Write-Host "Getting app name for $($app.AppId)"
        try{
            $Application = Get-MgApplication -Filter "AppId eq '$($app.AppId)'"
            if(-not([string]::IsNullOrEmpty($Application.DisplayName))){
                #Add found application to list of applications
                $appObject = [PSCustomObject]@{
                    AppId = $Application.Id
                    DisplayName = $Application.DisplayName
                }
                $Applications.Add($appObject) | Out-Null
            }
            else{
                Write-Warning "AppId not found"
                $appObject = [PSCustomObject]@{
                    AppId = $app.AppId
                }
                $AppIdsNotFound.Add($appObject) | Out-Null
            }
        }
        catch{
            Write-Error $Error[0].Exception.Message
            break
        }
    }
    else{
        #Add found application to list of applications
        $appObject = [PSCustomObject]@{
            AppId = $app.AppId
            DisplayName = $AppsToCheck[$app.AppId]
        }
        $Applications.Add($appObject) | Out-Null
    }
}

$Applications | Out-GridView  -Title "EWS Applications"
$Applications | Export-Csv -Path "$($OutputPath)\EwsApps.csv" -NoTypeInformation -Force
$AppIdsNotFound | Export-Csv -Path "$($OutputPath)\EwsAppsNotFound.csv" -NoTypeInformation -Force