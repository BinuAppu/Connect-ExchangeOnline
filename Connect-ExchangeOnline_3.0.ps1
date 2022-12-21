<#
Author - Binu Balan
Created - Sep 2022
Purpose - When your machine is restricted to run only Signed Scripts. Below lines of code will help in loading O365 Module in signed mode.
Pre-Requsite - Code Signing Certificate
Supported for - Exchange Online 3.0
#>

$erroractionpreference = 'SilentlyContinue'

# if you are connecting to O365 provide below details. If you are connecting using normal authentication, set the below Values to $null
$CertThumbprint = "0D231187EF0BB1A7676CA240578C559B6CD394BE"
$AppID = "6dd8f2bc-f084-421b-af25-be5a97971660"
$Org = "appultd.onmicrosoft.com"
Import-Module ExchangeOnlineManagement
if ($null -eq $CertificateThumbPrint) {
    Connect-ExchangeOnline 
} 
Else {
    Connect-ExchangeOnline -CertificateThumbPrint $CertThumbprint -AppID $AppID -Organization $Org -UseRPSSession
}
$DownladedModule = $null
$DownladedModule = Get-ChildItem $env:TEMP\tmp* -Directory | Sort-Object -Property CreationTime | Select-Object -Last 1
Set-AuthenticodeSignature -FilePath $DownladedModule\*.psm1 -Certificate (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert)
Import-Module $DownladedModule
$erroractionpreference = 'Continue'