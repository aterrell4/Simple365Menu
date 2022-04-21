Write-Host -Foregroundcolor Yellow "Setting Permissions"
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
Set-ExecutionPolicy bypass -Force
Write-Host -Foregroundcolor Cyan "Installing Dependencies"
Install-Module -Name MSOnline
Install-Module -Name AzureAD
Install-Module -Name ExchangeOnlineManagement
Write-Host -ForegroundColor Red "###################################################################"
Write-Host -Foregroundcolor Red "#Provide valid Office365 tenant credentials. MFA OTP supported    #"
Write-Host -ForegroundColor Red "###################################################################"
Connect-MSolService
$option = ""
Write-Host "1 - Microsoft Office 365 Online"
Write-Host "2 - Microsoft Azure Active Directory"
Write-Host "3 - Microsoft Exchange Online (Email)"
Write-Host "4 - Microsoft Security and Compliance"
Write-Host "0 - Cancel and Exit"
while($option -notmatch '^1$|^2$|^3$|^4$|^0$' )
{
Write-Host "Specify Which Office365 Module You Would Like To Connect To:"
$option = Read-Host
}
if($option -eq "0") {
Write-Host "Exiting...."
pause
exit
}elseif ($option -eq "1") {
Write-Host "Connecting to Office 365 Online"
Connect-MsolService
}elseif ($option -eq "2") {
Write-Host "Connecting to Azure Active Directory"
Connect-AzureAD
}elseif ($option -eq "3") {
Write-Host "Connecting to Exchange Online (Email)"
Connect-ExchangeOnline
}elseif ($option -eq "4") {
Write-Host "Connecting to Microsoft Security and Compliance"
Connect-IPPSSession
}
