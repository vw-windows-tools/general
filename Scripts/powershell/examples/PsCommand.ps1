$computername = $env:COMPUTERNAME

$username = "DOMAIN\username"
$securepassword = Read-Host "Enter password of $username" -AsSecureString
$credentials= New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securepassword

Invoke-Command -Authentication Credssp -Credential $credentials -ComputerName $computername -ScriptBlock {

    whoami | out-file 'c:\temp\whoami.txt'

}