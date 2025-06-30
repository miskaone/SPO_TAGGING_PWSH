<#
.SYNOPSIS
    Applies classifications to SharePoint Online sites with advanced targeting, reporting, and notifications.

.DESCRIPTION
    This script connects to SharePoint Online using secure app-only authentication and can operate in several modes:
    - Default: Processes all SharePoint sites
    - CsvInclusion/CsvExclusion: Processes or skips sites based on a CSV file
    - HubSite/Template: Processes sites associated with a specific Hub or template
    
    Generates detailed HTML/CSV reports and optionally sends email notifications.
    Designed for secure, unattended execution in Azure Automation.

.PREREQUISITES
    - PnP.PowerShell module
    - Azure AD App Registration with 'Sites.FullControl.All' SharePoint application permission
    - Certificate for authentication stored securely
    - SMTP server access for email notifications (if enabled)

.PARAMETER Mode
    Operational mode. Default: 'Default'
    Valid values: 'Default', 'CsvInclusion', 'CsvExclusion', 'HubSite', 'Template'

.PARAMETER TenantAdminUrl
    SharePoint Admin Center URL (e.g., "https://{SP-Domain}-admin.sharepoint.com")

.PARAMETER TenantName
    Tenant name (e.g., "{SP-Domain}.onmicrosoft.com")

.PARAMETER SensitivityLabel
    Classification/sensitivity label to apply to sites

.PARAMETER CsvPath
    [Required for CsvInclusion/CsvExclusion] Path to CSV file with 'SiteUrl' column

.PARAMETER TargetValue
    [Required for HubSite/Template] Hub Site ID or Template Name (e.g., "STS#3")

.PARAMETER ReportPath
    Directory path for saving HTML and CSV reports

.PARAMETER SendEmailReport
    Switch to enable email report sending

.PARAMETER EmailTo
    [Required with -SendEmailReport] Recipient email addresses

.PARAMETER EmailFrom
    [Required with -SendEmailReport] Sender email address

.PARAMETER SmtpServer
    [Required with -SendEmailReport] SMTP server address

.PARAMETER SmtpPort
    SMTP server port (default: 25)

.PARAMETER UseSSL
    Use SSL for SMTP connection

.PARAMETER DryRun
    Preview changes without applying them

.EXAMPLE
    # Process all sites
    .\SharePoint-Governance.ps1 -Mode Default -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" `
        -TenantName "{SP-Domain}.onmicrosoft.com" -SensitivityLabel "Confidential" -ReportPath "C:\Reports"

.EXAMPLE
    # Exclude sites from CSV with email report
    .\SharePoint-Governance.ps1 -Mode CsvExclusion -CsvPath "C:\temp\exclude.csv" `
        -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" -TenantName "{SP-Domain}.onmicrosoft.com" `
        -SensitivityLabel "Internal" -ReportPath "C:\Reports" -SendEmailReport `
        -EmailTo "admin@{SP-Domain}.com" -EmailFrom "noreply@{SP-Domain}.com" -SmtpServer "smtp.{SP-Domain}.com"

.EXAMPLE
    # Process with advanced options: batch processing, throttling, and custom filters
    $filters = @{
        StorageUsageCurrent = { $_ -gt 1GB }
        Template = { $_ -notlike "*CHANNEL*" }
    }
    .\SharePoint-Governance.ps1 -Mode Default -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" `
        -TenantName "{SP-Domain}.onmicrosoft.com" -SensitivityLabel "Public" -ReportPath "C:\Reports" `
        -BatchSize 25 -ThrottleDelayMs 500 -MaxRetries 5 -CustomFilters $filters `
        -LogLevel Detailed -GenerateChangeLog

.EXAMPLE
    # Run during maintenance window only with additional metadata collection
    .\SharePoint-Governance.ps1 -Mode HubSite -TargetValue "hub-guid-here" `
        -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" -TenantName "{SP-Domain}.onmicrosoft.com" `
        -SensitivityLabel "Restricted" -ReportPath "C:\Reports" `
        -AllowedHours 22,23,0,1,2,3 -AdditionalProperties "Owners","LastModified" `
        -ExportFailedSites -DryRun

.NOTES
    Version: 3.0
    Last Modified: 2025-06-30
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Default', 'CsvInclusion', 'CsvExclusion', 'HubSite', 'Template')]
    [string]$Mode = 'Default',

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[\w-]+-admin\.sharepoint\.com$')]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('\.onmicrosoft\.com$')]
    [string]$TenantName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SensitivityLabel,

    [Parameter()]
    [ValidateScript({
        if ($_ -and !(Test-Path $_)) { throw "CSV file not found: $_" }
        return $true
    })]
    [string]$CsvPath,

    [Parameter()]
    [string]$TargetValue,
    
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (!(Test-Path $_ -PathType Container)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
        }
        return $true
    })]
    [string]$ReportPath,

    [Parameter()]
    [switch]$SendEmailReport,

    [Parameter()]
    [ValidatePattern('^[\w\.-]+@[\w\.-]+\.\w+$')]
    [string[]]$EmailTo,

    [Parameter()]
    [ValidatePattern('^[\w\.-]+@[\w\.-]+\.\w+$')]
    [string]$EmailFrom,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SmtpServer,

    [Parameter()]
    [ValidateRange(1, 65535)]
    [int]$SmtpPort = 25,

    [Parameter()]
    [switch]$UseSSL,

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [ValidateRange(1, 100)]
    [int]$BatchSize = 50,

    [Parameter()]
    [int]$ThrottleDelayMs = 100,

    [Parameter()]
    [int]$MaxRetries = 3,

    [Parameter()]
    [switch]$ExportFailedSites,

    [Parameter()]
    [ValidateSet('Basic', 'Detailed', 'Verbose')]
    [string]$LogLevel = 'Basic',

    [Parameter()]
    [string[]]$AdditionalProperties,

    [Parameter()]
    [switch]$SkipCertificateCheck,

    [Parameter()]
    [hashtable]$CustomFilters,

    [Parameter()]
    [ValidateScript({
        if ($_ -lt 0 -or $_ -gt 23) { throw "Hour must be between 0 and 23" }
        return $true
    })]
    [int[]]$AllowedHours,

    [Parameter()]
    [switch]$GenerateChangeLog
)

#region Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug', 'Verbose')]
        [string]$Level = 'Info'
    )
    
    # Check log level filtering
    $shouldLog = switch ($script:LogLevel) {
        'Verbose' { $true }
        'Detailed' { $Level -ne 'Verbose' }
        'Basic' { $Level -in 'Info', 'Warning', 'Error', 'Success' }
    }
    
    if (!$shouldLog) { return }
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Info'    { Write-Host $logEntry -ForegroundColor White }
        'Warning' { Write-Host $logEntry -ForegroundColor Yellow }
        'Error'   { Write-Host $logEntry -ForegroundColor Red }
        'Success' { Write-Host $logEntry -ForegroundColor Green }
        'Debug'   { Write-Host $logEntry -ForegroundColor Gray }
        'Verbose' { Write-Host $logEntry -ForegroundColor DarkGray }
    }
    
    # Also write to log file
    $logFile = Join-Path $ReportPath "SharePoint-Governance-$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $logFile -Value $logEntry
}

function Test-MaintenanceWindow {
    <#
    .SYNOPSIS
        Checks if current time is within allowed hours
    #>
    if (!$AllowedHours) { return $true }
    
    $currentHour = (Get-Date).Hour
    if ($currentHour -notin $AllowedHours) {
        Write-Log "Current hour ($currentHour) is outside allowed hours: $($AllowedHours -join ', ')" -Level Warning
        return $false
    }
    return $true
}

function Start-ThrottledOperation {
    <#
    .SYNOPSIS
        Implements retry logic with exponential backoff
    #>
    param(
        [scriptblock]$Operation,
        [string]$OperationName,
        [int]$MaxRetries = $script:MaxRetries
    )
    
    $retryCount = 0
    $completed = $false
    $result = $null
    
    while (!$completed -and $retryCount -lt $MaxRetries) {
        try {
            $result = & $Operation
            $completed = $true
        }
        catch {
            $retryCount++
            if ($retryCount -ge $MaxRetries) {
                throw
            }
            
            $delay = [Math]::Pow(2, $retryCount) * 1000  # Exponential backoff
            Write-Log "Operation failed, attempt $retryCount of $MaxRetries. Waiting $($delay)ms..." -Level Warning
            Start-Sleep -Milliseconds $delay
        }
    }
    
    return $result
}

function Get-SiteMetadata {
    <#
    .SYNOPSIS
        Retrieves additional site properties if requested
    #>
    param(
        [object]$Site
    )
    
    $metadata = @{
        Url = $Site.Url
        Title = $Site.Title
        Template = $Site.Template
        StorageUsed = $Site.StorageUsageCurrent
        LastContentModified = $Site.LastContentModifiedDate
        SiteOwners = @()
        CreatedDate = $null
    }
    
    if ($AdditionalProperties) {
        Write-Log "Retrieving additional properties for $($Site.Url)" -Level Verbose
        
        try {
            Connect-PnPOnline -Url $Site.Url -ClientId $script:credentials.ClientId `
                             -Thumbprint $script:credentials.CertThumbprint `
                             -Tenant $TenantName -ErrorAction Stop
            
            $web = Get-PnPWeb -Includes Created, SiteUsers
            $metadata.CreatedDate = $web.Created
            
            if ('Owners' -in $AdditionalProperties) {
                $owners = Get-PnPSiteCollectionAdmin
                $metadata.SiteOwners = $owners | ForEach-Object { $_.Email }
            }
            
            if ('LastModified' -in $AdditionalProperties) {
                $metadata.LastContentModified = $web.LastItemModifiedDate
            }
        }
        catch {
            Write-Log "Failed to retrieve additional properties: $_" -Level Warning
        }
    }
    
    return $metadata
}

function Export-FailedSites {
    <#
    .SYNOPSIS
        Exports failed sites for retry processing
    #>
    param(
        [array]$FailedSites
    )
    
    if ($FailedSites.Count -eq 0) { return }
    
    $failedPath = Join-Path $ReportPath "Failed-Sites-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $FailedSites | Select-Object SiteUrl, @{N='FailureReason';E={$_.Message}} | 
        Export-Csv -Path $failedPath -NoTypeInformation
    
    Write-Log "Exported $($FailedSites.Count) failed sites to: $failedPath" -Level Warning
}

function New-ChangeLogEntry {
    <#
    .SYNOPSIS
        Creates detailed change log for audit purposes
    #>
    param(
        [object]$Site,
        [string]$OldValue,
        [string]$NewValue,
        [string]$Status
    )
    
    if (!$GenerateChangeLog) { return }
    
    $changeLog = @{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        UserPrincipal = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        SiteUrl = $Site.Url
        SiteId = $Site.Id
        PropertyChanged = 'Classification'
        OldValue = $OldValue
        NewValue = $NewValue
        Status = $Status
        ScriptVersion = '3.0'
        ExecutionId = $script:executionId
    }
    
    $logPath = Join-Path $ReportPath "ChangeLog-$(Get-Date -Format 'yyyyMMdd').json"
    $changeLog | ConvertTo-Json -Compress | Add-Content -Path $logPath
}

function Test-CustomFilters {
    <#
    .SYNOPSIS
        Applies custom filter logic to sites
    #>
    param(
        [object]$Site
    )
    
    if (!$CustomFilters) { return $true }
    
    foreach ($key in $CustomFilters.Keys) {
        $expectedValue = $CustomFilters[$key]
        $actualValue = $Site.$key
        
        if ($expectedValue -is [scriptblock]) {
            if (!(& $expectedValue $actualValue)) {
                Write-Log "Site $($Site.Url) filtered out by custom filter: $key" -Level Debug
                return $false
            }
        }
        elseif ($actualValue -ne $expectedValue) {
            return $false
        }
    }
    
    return $true
}

function Get-SecureCredentials {
    <#
    .SYNOPSIS
        Retrieves credentials from secure storage
    #>
    Write-Log "Retrieving secure credentials..." -Level Info
    
    # Production: Use Azure Key Vault or Automation Variables
    if ($env:AUTOMATION_ASSET_ACCOUNTID) {
        # Running in Azure Automation
        try {
            $clientId = Get-AutomationVariable -Name 'SharePoint-AppID'
            $certThumbprint = Get-AutomationVariable -Name 'SharePoint-CertThumbprint'
        }
        catch {
            Write-Log "Failed to retrieve automation variables: $_" -Level Error
            throw
        }
    }
    else {
        # Local development - check for environment variables
        $clientId = $env:SHAREPOINT_APP_ID
        $certThumbprint = $env:SHAREPOINT_CERT_THUMBPRINT
        
        if (!$clientId -or !$certThumbprint) {
            Write-Log "Credentials not found in environment variables. Please set SHAREPOINT_APP_ID and SHAREPOINT_CERT_THUMBPRINT" -Level Error
            throw "Missing credentials"
        }
    }
    
    return @{
        ClientId = $clientId
        CertThumbprint = $certThumbprint
    }
}

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Validates script prerequisites
    #>
    Write-Log "Validating prerequisites..." -Level Info
    
    # Check PnP module
    if (!(Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Log "PnP.PowerShell module not found. Install with: Install-Module PnP.PowerShell" -Level Error
        return $false
    }
    
    # Validate mode-specific parameters
    if ($Mode -in 'CsvInclusion', 'CsvExclusion' -and !$CsvPath) {
        Write-Log "Mode '$Mode' requires -CsvPath parameter" -Level Error
        return $false
    }
    
    if ($Mode -in 'HubSite', 'Template' -and !$TargetValue) {
        Write-Log "Mode '$Mode' requires -TargetValue parameter" -Level Error
        return $false
    }
    
    # Validate email parameters
    if ($SendEmailReport) {
        if (!$EmailTo -or !$EmailFrom -or !$SmtpServer) {
            Write-Log "Email notification requires -EmailTo, -EmailFrom, and -SmtpServer parameters" -Level Error
            return $false
        }
    }
    
    return $true
}

function Get-TargetSites {
    <#
    .SYNOPSIS
        Retrieves sites based on selected mode
    #>
    param(
        [array]$AllSites
    )
    
    Write-Log "Filtering sites based on mode: $Mode" -Level Info
    
    switch ($Mode) {
        'Default' {
            return $AllSites
        }
        
        'CsvInclusion' {
            $inclusionList = @((Import-Csv -Path $CsvPath).SiteUrl)
            Write-Log "Loaded $($inclusionList.Count) sites from inclusion list" -Level Info
            return $AllSites | Where-Object { $_.Url -in $inclusionList }
        }
        
        'CsvExclusion' {
            $exclusionList = @((Import-Csv -Path $CsvPath).SiteUrl)
            Write-Log "Loaded $($exclusionList.Count) sites from exclusion list" -Level Info
            return $AllSites | Where-Object { $_.Url -notin $exclusionList }
        }
        
        'HubSite' {
            Write-Log "Retrieving detailed site information for Hub filtering..." -Level Info
            $detailedSites = Get-PnPTenantSite -IncludeOneDriveSites:$false -Detailed
            return $detailedSites | Where-Object { $_.HubSiteId -eq $TargetValue }
        }
        
        'Template' {
            return $AllSites | Where-Object { $_.Template -eq $TargetValue }
        }
    }
}

function New-HtmlReport {
    <#
    .SYNOPSIS
        Generates HTML report with styling
    #>
    param(
        [array]$ReportData,
        [string]$FilePath
    )
    
    $stats = @{
        Total = $ReportData.Count
        Success = ($ReportData | Where-Object { $_.Status -eq 'Success' }).Count
        Failed = ($ReportData | Where-Object { $_.Status -eq 'Failed' }).Count
        Skipped = ($ReportData | Where-Object { $_.Action -eq 'Skipped' }).Count
    }
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>SharePoint Classification Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0078d4;
            border-bottom: 2px solid #0078d4;
            padding-bottom: 10px;
        }
        .summary {
            display: flex;
            gap: 20px;
            margin: 20px 0;
        }
        .stat-card {
            flex: 1;
            padding: 15px;
            border-radius: 5px;
            text-align: center;
            color: white;
        }
        .stat-card.total { background-color: #0078d4; }
        .stat-card.success { background-color: #107c10; }
        .stat-card.failed { background-color: #d83b01; }
        .stat-card.skipped { background-color: #ffb900; }
        .stat-number { font-size: 2em; font-weight: bold; }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th {
            background-color: #0078d4;
            color: white;
            padding: 12px;
            text-align: left;
            position: sticky;
            top: 0;
        }
        td {
            padding: 10px;
            border-bottom: 1px solid #e0e0e0;
        }
        tr:hover {
            background-color: #f8f8f8;
        }
        .status-success { color: #107c10; font-weight: bold; }
        .status-failed { color: #d83b01; font-weight: bold; }
        .status-skipped { color: #ffb900; }
        .metadata {
            margin-top: 20px;
            padding: 10px;
            background-color: #f0f0f0;
            border-radius: 5px;
            font-size: 0.9em;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>SharePoint Classification Report</h1>
        
        <div class="summary">
            <div class="stat-card total">
                <div class="stat-number">$($stats.Total)</div>
                <div>Total Sites</div>
            </div>
            <div class="stat-card success">
                <div class="stat-number">$($stats.Success)</div>
                <div>Successful</div>
            </div>
            <div class="stat-card failed">
                <div class="stat-number">$($stats.Failed)</div>
                <div>Failed</div>
            </div>
            <div class="stat-card skipped">
                <div class="stat-number">$($stats.Skipped)</div>
                <div>Skipped</div>
            </div>
        </div>
        
        <table>
            <thead>
                <tr>
                    <th>Timestamp</th>
                    <th>Site URL</th>
                    <th>Status</th>
                    <th>Action</th>
                    <th>Old Label</th>
                    <th>New Label</th>
                    <th>Message</th>
                </tr>
            </thead>
            <tbody>
"@
    
    foreach ($item in $ReportData) {
        $statusClass = "status-$($item.Status.ToLower())"
        $html += @"
                <tr>
                    <td>$($item.Timestamp)</td>
                    <td><a href="$($item.SiteUrl)" target="_blank">$($item.SiteUrl)</a></td>
                    <td class="$statusClass">$($item.Status)</td>
                    <td>$($item.Action)</td>
                    <td>$($item.OldLabel)</td>
                    <td>$($item.NewLabel)</td>
                    <td>$($item.Message)</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
        
        <div class="metadata">
            <strong>Report Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>
            <strong>Mode:</strong> $Mode<br>
            <strong>Target Label:</strong> $SensitivityLabel<br>
            $(if ($DryRun) { '<strong>Note:</strong> This was a dry run - no changes were applied.<br>' })
        </div>
    </div>
</body>
</html>
"@
    
    Set-Content -Path $FilePath -Value $html -Encoding UTF8
    Write-Log "HTML report saved to: $FilePath" -Level Success
}
#endregion

#region Main Script
try {
    # Validate prerequisites
    if (!(Test-Prerequisites)) {
        throw "Prerequisites validation failed"
    }
    
    # Initialize report data collection
    $script:reportData = @()
    $script:executionId = [guid]::NewGuid().ToString()
    $script:credentials = $null
    $script:LogLevel = $LogLevel
    $script:MaxRetries = $MaxRetries
    $startTime = Get-Date
    
    Write-Log "=== SharePoint Governance Script Started ===" -Level Info
    Write-Log "Execution ID: $($script:executionId)" -Level Info
    Write-Log "Mode: $Mode | Target Label: $SensitivityLabel" -Level Info
    
    # Check maintenance window
    if (!$(Test-MaintenanceWindow)) {
        Write-Log "Script execution blocked - outside maintenance window" -Level Error
        return
    }
    
    if ($DryRun) {
        Write-Log "DRY RUN MODE - No changes will be applied" -Level Warning
    }
    
    # Get credentials
    $script:credentials = Get-SecureCredentials
    
    # Connect to SharePoint
    Write-Log "Connecting to SharePoint tenant..." -Level Info
    $connectParams = @{
        Url = $TenantAdminUrl
        ClientId = $script:credentials.ClientId
        Thumbprint = $script:credentials.CertThumbprint
        Tenant = $TenantName
        ErrorAction = 'Stop'
    }
    
    if ($SkipCertificateCheck) {
        $connectParams['SkipTenantAdminCheck'] = $true
    }
    
    try {
        Connect-PnPOnline @connectParams
        Write-Log "Successfully connected to SharePoint" -Level Success
    }
    catch {
        Write-Log "Failed to connect to SharePoint: $_" -Level Error
        throw
    }
    
    # Retrieve all sites
    Write-Log "Retrieving SharePoint sites..." -Level Info
    $allSites = @(Get-PnPTenantSite -IncludeOneDriveSites:$false -ErrorAction Stop)
    Write-Log "Found $($allSites.Count) total sites" -Level Info
    
    # Filter sites based on mode
    $sitesToProcess = @(Get-TargetSites -AllSites $allSites)
    
    # Apply custom filters if provided
    if ($CustomFilters) {
        $sitesToProcess = $sitesToProcess | Where-Object { Test-CustomFilters -Site $_ }
    }
    
    Write-Log "Filtered to $($sitesToProcess.Count) sites for processing" -Level Info
    
    if ($sitesToProcess.Count -eq 0) {
        Write-Log "No sites match the specified criteria" -Level Warning
    }
    else {
        # Process sites in batches
        $batches = [Math]::Ceiling($sitesToProcess.Count / $BatchSize)
        Write-Log "Processing sites in $batches batches of $BatchSize" -Level Info
        
        $progressCount = 0
        $failedSites = @()
        
        for ($batchNum = 0; $batchNum -lt $batches; $batchNum++) {
            $batchStart = $batchNum * $BatchSize
            $batchEnd = [Math]::Min($batchStart + $BatchSize, $sitesToProcess.Count)
            $batch = $sitesToProcess[$batchStart..($batchEnd - 1)]
            
            Write-Log "Processing batch $($batchNum + 1) of $batches" -Level Info
            
            foreach ($site in $batch) {
                $progressCount++
                $percentComplete = ($progressCount / $sitesToProcess.Count) * 100
                
                Write-Progress -Activity "Processing SharePoint Sites" `
                              -Status "Site $progressCount of $($sitesToProcess.Count)" `
                              -PercentComplete $percentComplete `
                              -CurrentOperation $site.Url
                
                $result = [PSCustomObject]@{
                    Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    SiteUrl = $site.Url
                    Status = 'Success'
                    Action = ''
                    Message = ''
                    OldLabel = $site.Classification
                    NewLabel = $SensitivityLabel
                    ExecutionId = $script:executionId
                }
                
                try {
                    Write-Log "Processing: $($site.Url)" -Level Info
                    
                    # Get additional metadata if requested
                    if ($AdditionalProperties) {
                        $metadata = Get-SiteMetadata -Site $site
                        $result | Add-Member -NotePropertyName 'Metadata' -NotePropertyValue $metadata
                    }
                    
                    if ($site.Classification -eq $SensitivityLabel) {
                        $result.Action = 'Skipped'
                        $result.Message = 'Site already has the target classification'
                        Write-Log "  └─ Skipped: Already classified as '$SensitivityLabel'" -Level Info
                    }
                    else {
                        if ($DryRun) {
                            $result.Action = 'DryRun'
                            $result.Message = "Would change from '$($site.Classification)' to '$SensitivityLabel'"
                            Write-Log "  └─ DryRun: Would apply '$SensitivityLabel'" -Level Warning
                        }
                        elseif ($PSCmdlet.ShouldProcess($site.Url, "Apply classification '$SensitivityLabel'")) {
                            # Use retry logic for the actual update
                            Start-ThrottledOperation -OperationName "Set-Classification" -Operation {
                                Set-PnPTenantSite -Identity $site.Url `
                                                 -Classification $SensitivityLabel `
                                                 -ErrorAction Stop
                            }
                            
                            $result.Action = 'Applied'
                            $result.Message = "Successfully changed from '$($site.Classification)' to '$SensitivityLabel'"
                            Write-Log "  └─ Success: Applied '$SensitivityLabel'" -Level Success
                            
                            # Create change log entry
                            New-ChangeLogEntry -Site $site -OldValue $site.Classification `
                                             -NewValue $SensitivityLabel -Status 'Success'
                        }
                        else {
                            $result.Action = 'Cancelled'
                            $result.Message = 'Operation cancelled by user'
                        }
                    }
                }
                catch {
                    $result.Status = 'Failed'
                    $result.Action = 'Error'
                    $result.Message = $_.Exception.Message
                    Write-Log "  └─ Error: $_" -Level Error
                    $failedSites += $result
                }
                
                $script:reportData += $result
                
                # Throttle between operations
                if ($ThrottleDelayMs -gt 0) {
                    Start-Sleep -Milliseconds $ThrottleDelayMs
                }
            }
            
            # Pause between batches
            if ($batchNum -lt ($batches - 1)) {
                Write-Log "Pausing between batches..." -Level Debug
                Start-Sleep -Seconds 2
            }
        }
        
        Write-Progress -Activity "Processing SharePoint Sites" -Completed
        
        # Export failed sites if requested
        if ($ExportFailedSites -and $failedSites.Count -gt 0) {
            Export-FailedSites -FailedSites $failedSites
        }
    }
    
    # Generate reports
    if ($script:reportData.Count -gt 0) {
        Write-Log "Generating reports..." -Level Info
        
        $reportBaseName = "SharePoint-Classification-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        $csvPath = Join-Path $ReportPath "$reportBaseName.csv"
        $htmlPath = Join-Path $ReportPath "$reportBaseName.html"
        
        # CSV Report
        $script:reportData | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Log "CSV report saved to: $csvPath" -Level Success
        
        # HTML Report
        New-HtmlReport -ReportData $script:reportData -FilePath $htmlPath
        
        # Send email if requested
        if ($SendEmailReport) {
            Write-Log "Sending email report..." -Level Info
            
            $stats = @{
                Success = ($script:reportData | Where-Object { $_.Status -eq 'Success' -and $_.Action -eq 'Applied' }).Count
                Failed = ($script:reportData | Where-Object { $_.Status -eq 'Failed' }).Count
                Skipped = ($script:reportData | Where-Object { $_.Action -eq 'Skipped' }).Count
            }
            
            $subject = "SharePoint Classification Report - $($stats.Success) Applied, $($stats.Failed) Failed"
            
            $emailParams = @{
                To = $EmailTo
                From = $EmailFrom
                Subject = $subject
                Body = Get-Content $htmlPath -Raw
                BodyAsHtml = $true
                SmtpServer = $SmtpServer
                Port = $SmtpPort
                UseSsl = $UseSSL
                Attachments = @($csvPath, $htmlPath)
                ErrorAction = 'Stop'
            }
            
            try {
                Send-MailMessage @emailParams
                Write-Log "Email sent successfully to: $($EmailTo -join ', ')" -Level Success
            }
            catch {
                Write-Log "Failed to send email: $_" -Level Error
            }
        }
    }
    
    # Summary
    $duration = (Get-Date) - $startTime
    Write-Log "=== Script Completed ===" -Level Success
    Write-Log "Total execution time: $($duration.ToString('hh\:mm\:ss'))" -Level Info
    
    # Display summary
    $summary = $script:reportData | Group-Object Status, Action | ForEach-Object {
        [PSCustomObject]@{
            Status = $_.Group[0].Status
            Action = $_.Group[0].Action
            Count = $_.Count
        }
    } | Sort-Object Status, Action
    
    Write-Log "Summary:" -Level Info
    $summary | ForEach-Object {
        Write-Log "  $($_.Status) - $($_.Action): $($_.Count) sites" -Level Info
    }
}
catch {
    Write-Log "Script failed with error: $_" -Level Error
    throw
}
finally {
    # Cleanup
    if (Get-PnPConnection) {
        Disconnect-PnPOnline
        Write-Log "Disconnected from SharePoint" -Level Info
    }
}
#endregion
