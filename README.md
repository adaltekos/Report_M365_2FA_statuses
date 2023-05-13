# Report M365 2FA statuses
## Description
This PowerShell script retrieves a list of all users in the Microsoft 365 tenant and their status of Enable, Disable, Enforced 2FA Status. It exports the information to an Excel file stored in SharePoint Online.

## Prerequisites
- PowerShell version 5.1 or later
- Installed SharePointPnPPowerShellOnline module
- Installed MSOnline module
- Installed ImportExcel module
- Installed CredentialManager module

## Configuration
Before running the script, you need to modify the following variables in the script:
- `$filename`: Complete with filename (e.g. Raport_M365_2FA_Status.xlsx)
- `$localPath`: Complete with local path (e.g. C:\Raporty\)
- `$onlinePath`: Complete with path where the file is located on SharePoint (e.g. Shared Documents/Global/)
- `$cred`: Complete with the name of the stored credential in the credential manager
- `Url`: Complete with the URL of the SharePoint site (e.g. https://company.sharepoint.com/sites/it-dep)
- `Tenant`: Complete with the tenant name (e.g. company.onmicrosoft.com)
- `ClientId`: Complete with the ClientId (which is the ID of the application registered in Azure AD)
- `Thumbprint`: Complete with the Thumbprint (which is the certificate thumbprint)
- 
