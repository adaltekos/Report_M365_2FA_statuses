#Install-Module SharePointPnPPowerShellOnline
#Install-Module MSOnline
#Install-Module ImportExcel
#Install-Module CredentialManager

$filename 		= "" #Complete with filename (ex. Raport_M365_2FA_Status.xlsx)
$localPath 		= "" #Complete with local path (ex. C:\Raporty\)
$onlinePath		= "" #Complete with path where file is on sharepoint (ex. Shared Documents/Global/)
$cred = Get-StoredCredential -Target '' #Complete with name of stored credential in credential manager

$pnpConnectParams  = @{
    Url				= "" #Complete with Url site (ex. https://company.sharepoint.com/sites/it-dep)
    Tenant			= "" #Complete with tenant name (ex. company.onmicrosoft.com)
    ClientId		= "" #Complete with ClientId (which is ID of application registered in Azure AD)
    Thumbprint		= "" #Complete with Thumbprint (which is certificate thumbprint)
}
Connect-PnPOnline @pnpConnectParams

$getPnPFileParams = @{
    Url				= ($onlinePath + $filename)
    Path			= $localPath
    Filename		= $filename
    AsFile			= $true
    Force			= $true
}
Get-PnPFile @getPnPFileParams

Start-Sleep -s 3

#Clear Excel file
$excel = Open-ExcelPackage -Path ($localPath + $filename)
$excel.new.Cells["A2:E500"].Clear()
Close-ExcelPackage -ExcelPackage $excel

Start-Sleep -s 3

#Connect and get list of all users and their status of Enable, Disable, Enforced 2FA Status
Connect-MsolService -Credential $cred
Get-MsolUser -all | select UserPrincipalName,DisplayName,@{N="directorySynced"; E={if ($_.ImmutableId -ne $null){"True"} else {"False"}}},@{N="Licenses"; E={if ($_.Licenses.AccountSkuId -like "DUONDystrybucja:O365_BUSINESS_PREMIUM"){"Microsoft 365 Business Standard"} elseif ($_.Licenses.AccountSkuId -like "DUONDystrybucja:O365_BUSINESS_ESSENTIALS") {"Microsoft 365 Business Basic"}}},@{N="MFA Status"; E={ if( $_.StrongAuthenticationRequirements.State -ne $null) {$_.StrongAuthenticationRequirements.State} else { "Disabled"}}} | Export-Excel -Path ($onlinePath + $filename) -WorkSheetname new -AutoSize

$addPnPFileParams  = @{
    Folder		= $onlinePath	
    Path		= ($localPath + $filename)
}
Add-PnPFile @addPnPFileParams