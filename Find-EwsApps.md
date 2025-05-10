# Find-EwsApps

Get the application names for appIDs found in the EWS usage report.

## Ews Usage Report
The EWS usage report provides a list of active appIds within a tenant. This script downloads a list of first-party Microsoft apps to check against the EWS usage report. It also queries Microsoft Entra ID for all other appIDs to provide a list of application names. Two CSV files are generated: one containing a list of found applications and another with a list of unknown apps.

### Using the script
```powershell
.\Find-EwsApps.ps1 -OutputPath C:\Temp\Output\ -EwsUsageReportPath C:\Temp\Output\EWSWeeklyUsage_4_11_2025_12_10_14.csv
```

## Parameters

**OutputPath** - The OutputPath parameter specifies the path for the output files.

**EwsUsageReportPath** - The path to the EWS usage report that was exported from the M365 Admin Center.