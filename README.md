# SPO_TAGGING_PWSH
Powershell Script for reviewing and Tagging SharePoint Sites for Governance

# SharePoint Governance Automation Script

A comprehensive PowerShell script for automating SharePoint Online site classification and governance at scale. This enterprise-ready solution provides advanced targeting, batch processing, detailed reporting, and robust error handling.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Parameters](#parameters)
- [Filtering Modes](#filtering-modes)
- [Advanced Features](#advanced-features)
- [Examples](#examples)
- [Reporting](#reporting)
- [Troubleshooting](#troubleshooting)
- [Best Practices](#best-practices)
- [Security Considerations](#security-considerations)

## Features

### Core Capabilities

- **Automated Classification**: Apply sensitivity labels to SharePoint sites in bulk
- **Multiple Targeting Modes**: Process all sites, specific sites from CSV, hub sites, or template-based filtering
- **Batch Processing**: Handle thousands of sites efficiently with configurable batch sizes
- **Retry Logic**: Automatic retry with exponential backoff for transient failures
- **Detailed Reporting**: Generate HTML and CSV reports with visual statistics
- **Email Notifications**: Send automated reports to stakeholders
- **Audit Trail**: Create detailed change logs for compliance

### Advanced Features

- **Custom Filtering**: Apply complex business rules using PowerShell scriptblocks
- **Maintenance Windows**: Restrict execution to specific hours
- **Throttling**: Configurable delays to avoid rate limiting
- **Dry Run Mode**: Preview changes before applying them
- **Additional Metadata**: Collect extra site information (owners, last modified dates)
- **Failed Site Export**: Save failed sites for retry processing

## Prerequisites

### Required Components

1. **PowerShell 5.1 or later**

2. **PnP.PowerShell Module**

   ```powershell
   Install-Module -Name PnP.PowerShell -Force
   ```

3. **Azure AD App Registration** with:

   - Certificate-based authentication
   - SharePoint API permission: `Sites.FullControl.All`
   - Admin consent granted

4. **SharePoint Admin Access** to your tenant

### Optional Components

- **SMTP Server** (for email notifications)
- **Azure Automation Account** (for scheduled execution)
- **Azure Key Vault** (for secure credential storage)

## Installation

1. **Download the Script**

   ```powershell
   # Save the script to your preferred location
   $scriptPath = "C:\Scripts\SharePoint-Governance.ps1"
   ```

2. **Install PnP.PowerShell Module**

   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
   Import-Module PnP.PowerShell
   ```

3. **Create Azure AD App Registration**

   ```powershell
   # Generate a self-signed certificate
   $cert = New-SelfSignedCertificate -Subject "CN=SharePointGovernance" `
                                     -CertStoreLocation "Cert:\CurrentUser\My" `
                                     -KeyExportPolicy Exportable `
                                     -KeySpec Signature `
                                     -KeyLength 2048 `
                                     -HashAlgorithm SHA256
   
   # Export the certificate
   Export-Certificate -Cert $cert -FilePath ".\SharePointGovernance.cer"
   ```

4. **Configure App Registration in Azure Portal**

   - Navigate to Azure AD > App registrations
   - Create new registration
   - Upload certificate under "Certificates & secrets"
   - Add API permission: SharePoint > Application permissions > Sites.FullControl.All
   - Grant admin consent

## Configuration

### Environment Variables (Recommended for Development)

```powershell
# Set credentials as environment variables
$env:SHAREPOINT_APP_ID = "your-app-id-guid"
$env:SHAREPOINT_CERT_THUMBPRINT = "your-certificate-thumbprint"
```

### Azure Automation Variables (Production)

- Create automation variables:
  - `SharePoint-AppID`
  - `SharePoint-CertThumbprint`

### Azure Key Vault (Most Secure)

```powershell
# Store secrets in Key Vault
Set-AzKeyVaultSecret -VaultName "YourKeyVault" `
                     -Name "SharePoint-AppID" `
                     -SecretValue (ConvertTo-SecureString "your-app-id" -AsPlainText -Force)
```

## Usage

### Basic Usage

```powershell
# Apply classification to all sites
.\SharePoint-Governance.ps1 `
    -Mode Default `
    -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" `
    -TenantName "{SP-Domain}.onmicrosoft.com" `
    -SensitivityLabel "Internal" `
    -ReportPath "C:\Reports"
```

### Dry Run (Preview Changes)

```powershell
# See what would happen without making changes
.\SharePoint-Governance.ps1 `
    -Mode Default `
    -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" `
    -TenantName "{SP-Domain}.onmicrosoft.com" `
    -SensitivityLabel "Confidential" `
    -ReportPath "C:\Reports" `
    -DryRun
```

## Parameters

### Required Parameters

| Parameter           | Type   | Description                                                  |
| ------------------- | ------ | ------------------------------------------------------------ |
| `-Mode`             | String | Targeting mode: Default, CsvInclusion, CsvExclusion, HubSite, Template |
| `-TenantAdminUrl`   | String | SharePoint Admin Center URL (e.g., `https://{SP-Domain}-admin.sharepoint.com`) |
| `-TenantName`       | String | Tenant name (e.g., `{SP-Domain}.onmicrosoft.com`)            |
| `-SensitivityLabel` | String | Classification label to apply                                |
| `-ReportPath`       | String | Directory path for saving reports                            |

### Mode-Specific Parameters

| Parameter      | Type   | Required For               | Description                           |
| -------------- | ------ | -------------------------- | ------------------------------------- |
| `-CsvPath`     | String | CsvInclusion, CsvExclusion | Path to CSV file containing site URLs |
| `-TargetValue` | String | HubSite, Template          | Hub Site ID or Template name          |

### Optional Parameters

| Parameter               | Type      | Default | Description                                        |
| ----------------------- | --------- | ------- | -------------------------------------------------- |
| `-BatchSize`            | Int       | 50      | Number of sites to process per batch               |
| `-ThrottleDelayMs`      | Int       | 100     | Milliseconds to wait between operations            |
| `-MaxRetries`           | Int       | 3       | Maximum retry attempts for failed operations       |
| `-LogLevel`             | String    | Basic   | Logging verbosity: Basic, Detailed, Verbose        |
| `-AllowedHours`         | Int[]     | None    | Hours (0-23) when script can run                   |
| `-AdditionalProperties` | String[]  | None    | Extra properties to collect (Owners, LastModified) |
| `-CustomFilters`        | Hashtable | None    | Custom filtering logic                             |
| `-DryRun`               | Switch    | False   | Preview mode without making changes                |
| `-ExportFailedSites`    | Switch    | False   | Export failed sites to CSV                         |
| `-GenerateChangeLog`    | Switch    | False   | Create detailed audit log                          |
| `-SendEmailReport`      | Switch    | False   | Send report via email                              |
| `-EmailTo`              | String[]  | None    | Email recipients (required with -SendEmailReport)  |
| `-EmailFrom`            | String    | None    | Sender email (required with -SendEmailReport)      |
| `-SmtpServer`           | String    | None    | SMTP server (required with -SendEmailReport)       |
| `-SmtpPort`             | Int       | 25      | SMTP port                                          |
| `-UseSSL`               | Switch    | False   | Use SSL for SMTP                                   |

## Filtering Modes

### 1. Default Mode

Process all SharePoint sites in the tenant (excluding OneDrive).

```powershell
-Mode Default
```

### 2. CSV Inclusion

Process only sites listed in a CSV file.

```powershell
-Mode CsvInclusion -CsvPath "C:\IncludeSites.csv"
```

CSV Format:

```csv
SiteUrl
https://{SP-Domain}.sharepoint.com/sites/Marketing
https://{SP-Domain}.sharepoint.com/sites/Sales
```

### 3. CSV Exclusion

Process all sites except those in the CSV file.

```powershell
-Mode CsvExclusion -CsvPath "C:\ExcludeSites.csv"
```

### 4. Hub Site

Process only sites associated with a specific hub.

```powershell
-Mode HubSite -TargetValue "hub-site-guid"
```

### 5. Template

Process only sites created from a specific template.

```powershell
-Mode Template -TargetValue "STS#3"
```

Common Templates:

- `STS#3` - Modern Team Site
- `SITEPAGEPUBLISHING#0` - Communication Site
- `GROUP#0` - Microsoft 365 Group-connected Site
- `TEAMCHANNEL#0` - Teams Channel Site

## Advanced Features

### Custom Filtering

Apply complex business rules using PowerShell scriptblocks:

```powershell
$filters = @{
    # Sites larger than 1GB
    StorageUsageCurrent = { $_ -gt 1GB }
    # Exclude Teams channels
    Template = { $_ -notlike "*CHANNEL*" }
    # Recently modified sites only
    LastContentModifiedDate = { $_ -gt (Get-Date).AddDays(-90) }
}

.\SharePoint-Governance.ps1 `
    -Mode Default `
    -CustomFilters $filters `
    # ... other parameters
```

### Maintenance Windows

Restrict execution to specific hours:

```powershell
# Only run between 10 PM and 3 AM
.\SharePoint-Governance.ps1 `
    -AllowedHours 22,23,0,1,2,3 `
    # ... other parameters
```

### Batch Processing with Throttling

Handle large environments efficiently:

```powershell
.\SharePoint-Governance.ps1 `
    -BatchSize 25 `           # Process 25 sites at a time
    -ThrottleDelayMs 500 `    # Wait 500ms between sites
    -MaxRetries 5 `           # Retry up to 5 times
    # ... other parameters
```

## Examples

### Example 1: Basic Classification

```powershell
# Apply "Internal" label to all sites
.\SharePoint-Governance.ps1 `
    -Mode Default `
    -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" `
    -TenantName "{SP-Domain}.onmicrosoft.com" `
    -SensitivityLabel "Internal" `
    -ReportPath "C:\Reports"
```

### Example 2: Department-Specific Processing

```powershell
# Create CSV with department sites
@"
SiteUrl
https://{SP-Domain}.sharepoint.com/sites/Finance
https://{SP-Domain}.sharepoint.com/sites/Legal
https://{SP-Domain}.sharepoint.com/sites/HR
"@ | Out-File "C:\DeptSites.csv"

# Process only these sites with email notification
.\SharePoint-Governance.ps1 `
    -Mode CsvInclusion `
    -CsvPath "C:\DeptSites.csv" `
    -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" `
    -TenantName "{SP-Domain}.onmicrosoft.com" `
    -SensitivityLabel "Confidential" `
    -ReportPath "C:\Reports" `
    -SendEmailReport `
    -EmailTo "admin@{SP-Domain}.com","security@{SP-Domain}.com" `
    -EmailFrom "noreply@{SP-Domain}.com" `
    -SmtpServer "smtp.{SP-Domain}.com"
```

### Example 3: Production Deployment with All Features

```powershell
# Advanced configuration for production
$customFilters = @{
    StorageUsageCurrent = { $_ -gt 500MB }
    Template = { $_ -notlike "*PERSONAL*" }
}

.\SharePoint-Governance.ps1 `
    -Mode Default `
    -TenantAdminUrl "https://{SP-Domain}-admin.sharepoint.com" `
    -TenantName "{SP-Domain}.onmicrosoft.com" `
    -SensitivityLabel "Public" `
    -ReportPath "D:\SharePointReports" `
    -BatchSize 20 `
    -ThrottleDelayMs 1000 `
    -MaxRetries 5 `
    -CustomFilters $customFilters `
    -LogLevel Detailed `
    -AllowedHours 23,0,1,2,3,4 `
    -AdditionalProperties "Owners","LastModified" `
    -ExportFailedSites `
    -GenerateChangeLog `
    -SendEmailReport `
    -EmailTo "sp-admins@{SP-Domain}.com" `
    -EmailFrom "sharepoint-automation@{SP-Domain}.com" `
    -SmtpServer "smtp.office365.com" `
    -SmtpPort 587 `
    -UseSSL
```

## Reporting

### HTML Report

The script generates a professional HTML report with:

- Visual statistics cards (Total, Success, Failed, Skipped)
- Detailed table of all processed sites
- Clickable site URLs
- Execution metadata
- Responsive design with modern styling

### CSV Report

Detailed CSV export containing:

- Timestamp
- Site URL
- Status (Success/Failed)
- Action taken
- Old and new classification values
- Error messages (if any)
- Execution ID for correlation

### Change Log (Optional)

JSON-formatted audit log with:

- Precise timestamps
- User principal
- Site details
- Before/after values
- Execution tracking

### Failed Sites Export (Optional)

CSV file containing only failed sites for easy retry:

- Site URL
- Failure reason

## Troubleshooting

### Common Issues

1. **Authentication Failures**
   - Verify certificate is installed in personal certificate store
   - Check app registration has correct permissions
   - Ensure admin consent is granted

2. **Rate Limiting**
   - Increase `ThrottleDelayMs` value
   - Reduce `BatchSize`
   - Use `AllowedHours` to run during off-peak times

3. **Permission Errors**
   - Verify account has SharePoint admin rights
   - Check site-specific permissions
   - Ensure app has Sites.FullControl.All permission

### Debug Mode

Enable verbose logging for troubleshooting:

```powershell
.\SharePoint-Governance.ps1 `
    -LogLevel Verbose `
    # ... other parameters
```

### Log Files

Check logs in the report directory:

- `SharePoint-Governance-YYYYMMDD.log` - Execution log
- `ChangeLog-YYYYMMDD.json` - Audit trail (if enabled)

## Best Practices

### 1. Testing

- Always run with `-DryRun` first
- Test on a small subset using CSV inclusion
- Validate in non-production environment

### 2. Performance

- Use batch processing for large environments
- Implement throttling to avoid rate limits
- Schedule during maintenance windows

### 3. Security

- Store credentials in Azure Key Vault
- Use certificate-based authentication
- Enable audit logging for compliance
- Regularly rotate certificates

### 4. Monitoring

- Enable email notifications for critical runs
- Review failed sites and investigate patterns
- Monitor execution times and adjust batch sizes

### 5. Maintenance

- Regularly update PnP.PowerShell module
- Review and update filtering criteria
- Archive old reports and logs

## Security Considerations

1. **Credential Management**
   - Never hardcode credentials in scripts
   - Use Azure Key Vault for production
   - Implement least-privilege access

2. **Network Security**
   - Run from secure, managed devices
   - Use encrypted connections (HTTPS/TLS)
   - Implement IP restrictions if possible

3. **Audit and Compliance**
   - Enable change logging
   - Retain reports per compliance requirements
   - Review logs regularly

4. **Access Control**
   - Limit who can execute the script
   - Use separate app registrations per environment
   - Implement approval workflows for production changes

## Support and Contribution

For issues, questions, or contributions:

1. Check existing documentation
2. Review troubleshooting section
3. Enable verbose logging for detailed diagnostics
4. Contact your SharePoint administrator

---

**Version:** 3.0  
**Last Updated:** June 30, 2025  
**PowerShell Compatibility:** 5.1+  
**Module Dependencies:** PnP.PowerShell
