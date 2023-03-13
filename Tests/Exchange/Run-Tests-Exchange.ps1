# Reference to Test Module
$ProjectPath = "C:\Users\John\Documents\Projects\Exchange Module\main"
$TempDir = "c:\temp"
$ModuleFileFullPath = "$ProjectPath\Source Code\Exchange.ps1"

# (re)Import Test Module
$ModuleName = [io.path]::GetFileNameWithoutExtension($ModuleFileFullPath)
Get-Module | Where-Object -Property "Name" -Like "*$ModuleName*" | Remove-Module
Import-Module $ModuleFileFullPath

# Credentials
$username = "john.smith@contoso.com"

# Recipients
$ToRecipient2 = "jack.martin@contoso.com"
$CcRecipient2 = "mike.hammer@contoso.com"
$BccRecipient2 = "watch@contoso.com"

# Meeting dates
$MeetingA_Start1 = '2021-08-26T17:00:00'
$MeetingA_End1 = '2021-08-26T17:30:00'
$MeetingA_Start2 = '2021-08-27T11:00:00'
$MeetingA_End2 = '2021-08-27T11:30:00'
$MeetingB_Start1 = '2021-08-28T12:30:00'
$MeetingB_End1 = '2021-08-28T14:00:00'

# Url / Dll
$ExchServerUrl = "https://owa.contoso.com/EWS/Exchange.asmx"
$ExchDllPath = "C:\Program Files\PackageManagement\NuGet\Packages\Exchange.WebServices.Managed.Api.2.2.1.2\lib\net35\Microsoft.Exchange.WebServices.dll"

if ($null -eq $SecurePassword) {
    $SecurePassword = Read-Host "Enter password of $username" -AsSecureString
}

$EnteredValue = Read-Host "Custom variables edited ? Folders and meetings cleared ? [Enter Y to continue]"
If ( [string]::IsNullOrEmpty($EnteredValue) -or ($EnteredValue.ToUpper() -ne "Y")) {exit}

. $ProjectPath\Tests\Exchange\Tests-Exchange.ps1