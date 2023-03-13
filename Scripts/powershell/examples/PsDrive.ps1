$SourceFile = "c:\documents\report.xls"
$ReportingPath = "\\shared\prod\50004\reporting"

$Username = "CONTOSO\srvMonitoring"
$SecurePwd= Read-Host "Enter password of $Username" -AsSecureString
$Credentials= New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePwd

New-PSDrive -Name ReportingSharedDir -PSProvider FileSystem -Root $ReportingPath -Credential $Credentials #-Persist
Copy-Item -Path $SourceFile -Destination "ReportingSharedDir:\report.xls"
Remove-PSDrive -Name ReportingSharedDir