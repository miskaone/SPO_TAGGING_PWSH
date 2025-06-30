# SPO_TAGGING_PWSH
Powershell Script for reviewing and Tagging SharePoint Sites for Governance

# SharePoint Tenant-Wide Site Classification

This repository contains a PowerShell script designed to automate the application of a classification (e.g., a sensitivity label) to all SharePoint Online sites across an entire tenant. This ensures consistent governance and labeling, even as new sites are created.

The script leverages the **PnP.PowerShell module** and uses secure, app-only authentication with a certificate to connect to your SharePoint tenant.

## Features

- **Tenant-Wide Application**: Scans and applies a classification to every SharePoint site.
- **Automation Ready**: Designed to be run as an automated task (e.g., in an Azure Automation Runbook) to classify new sites as they are created.
- **Consistent Governance**: Enforces a baseline security and compliance posture across your SharePoint estate.
- **Secure Authentication**: Uses modern, app-only authentication with a certificate, avoiding the need to store user credentials.

## Prerequisites

Before running this script, you must have the following in place:

1. **PnP.PowerShell Module**: The script requires the latest version of the PnP.PowerShell module. You can install it with:
   ```powershell
   Install-Module -Name PnP.PowerShell
   ```
2. **Azure AD App Registration**: An App Registration in Azure Active Directory is required for authentication.
3. **API Permissions**: The App Registration must be granted the following application permission for SharePoint:
   - `Tenant.ReadWrite.All`
4. **Certificate**: You need to create a self-signed certificate and upload the public key (`.cer` file) to your Azure AD App Registration. The script will use the certificate's thumbprint for authentication.

## Setup & Configuration

1. **Clone the Repository**:

   ```bash
   git clone [your-repository-url]
   ```

2. **Update Script Parameters**: Open the PowerShell script and update the following variables with your tenant's information:

   ```powershell
   # Parameters
   $TenantAdminUrl = "https://contoso-admin.sharepoint.com" # Your SharePoint admin center URL
   $ClientId       = "YOUR-APP-ID"                         # The Application (client) ID from your Azure AD App
   $Tenant         = "contoso.onmicrosoft.com"             # Your tenant domain name
   $CertThumbprint = "ABC123DEF456"                        # The thumbprint of the certificate
   $SensitivityLabel = "Confidential"                    # The classification/label to apply
   ```

## PowerShell Script

```powershell
# ==================================================================================
# SharePoint Tenant-Wide Site Classification
#
# This script applies a specified classification (sensitivity label) to all
# SharePoint sites in the tenant.
# ==================================================================================

# Parameters
$TenantAdminUrl = "https://contoso-admin.sharepoint.com"
$ClientId       = "YOUR-APP-ID"
$Tenant         = "contoso.onmicrosoft.com"
$CertThumbprint = "ABC123DEF456"
$SensitivityLabel = "Confidential"

# Connect using certificate-based app-only authentication
try {
    Connect-PnPOnline `
      -Tenant $Tenant `
      -ClientId $ClientId `
      -Thumbprint $CertThumbprint `
      -Url $TenantAdminUrl `
      -ErrorAction Stop
    Write-Host "Successfully connected to tenant." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to the tenant. Please check parameters and permissions." -ForegroundColor Red
    throw
}


# Iterate through all tenant sites (excluding OneDrive sites) and apply the classification
Get-PnPTenantSite -IncludeOneDriveSites:$false | ForEach-Object {
    try {
        Write-Host "Processing site:" $_.Url
        # Check if the site already has the desired classification
        if ($_.Classification -ne $SensitivityLabel) {
            Write-Host "  -> Applying classification '$SensitivityLabel'..."
            Set-PnPTenantSite `
              -Identity $_.Url `
              -Classification $SensitivityLabel

            Write-Host "  -> Successfully tagged site:" $_.Url -ForegroundColor Cyan
        } else {
            Write-Host "  -> Site is already classified correctly. Skipping." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  -> Failed to tag site: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "Script execution completed." -ForegroundColor Green
Disconnect-PnPOnline
```

## Usage

You can run this script manually from a machine that has the PnP.PowerShell module installed.

### Recommended Usage: Automation

For best results, automate this script using a serverless platform like **Azure Automation**.

1. **Create an Azure Automation Account**.
2. **Import Modules**: Add the `PnP.PowerShell` module to your Automation Account from the gallery.
3. **Add Certificate**: Add the certificate (`.pfx` file) to the Automation Account's "Certificates" asset store.
4. **Create Runbook**: Create a new PowerShell Runbook and paste the script content.
5. **Schedule**: Set a recurring schedule (e.g., daily) to ensure new sites are automatically classified.
6. **Use Variables**: For better security, store parameters like `$ClientId` and `$Tenant` in the Automation Account's "Variables" assets instead of hardcoding them in the script.

## Security Considerations

- **Permissions**: The `Tenant.ReadWrite.All` permission is highly privileged. Ensure that the App Registration and its credentials (the certificate) are kept secure.
- **Secrets Management**: For automation, it is highly recommended to store the certificate in **Azure Key Vault** and grant the Azure Automation's managed identity access to it. This is more secure than storing the certificate directly in the Automation Account.
