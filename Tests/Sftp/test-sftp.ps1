﻿function TestReceive-Winscp {

	# 1) Run with zero file on server
	# 2) Run with one or more files on server
	# 3) Run with one or more file matching filemask on server
	# 4) Run with wrong password
	# 5) Run with wrong port
	
    $workdir = "c:\0\w-tools"
    $tempdir = "c:\0\tmp"

    #Set-ExecutionPolicy -Scope "CurrentUser" -ExecutionPolicy "Unrestricted"
    Import-Module "$workdir\Source code"

    # Load General Config & Parameters
    $JSONConfigPath = "$tempdir\sftp_winscp_tests.json"
    $Config = (Get-Content $JSONConfigPath) | ConvertFrom-Json
    $serverkey = $Config.GENERAL.serverkey
    $server = $Config.GENERAL.server
    $port = $Config.GENERAL.port
    $user = $Config.GENERAL.user
    $filemask = $Config.GENERAL.filemask
    $destination = $Config.GENERAL.destination
    $path = $Config.GENERAL.path

    # Method 1 : direct input
    #$SecurePassword = Read-Host -AsSecureString

    # Method 2 : plain text (not recommended)
    #$Password = "P@ssw0rd"
    #$SecurePassword = ($Password | ConvertTo-SecureString -asPlainText -Force)

    # Method 3 : encrypted key (preferred)
    #-PasswordFile $tempdir\server_pwd.txt -KeyFile $tempdir\master-key.txt
    $key = Get-Content $Config.GENERAL.'encrypted-keyfile'
    $encpassword = $Config.GENERAL.'encrypted-password'
    $SecurePassword = $encpassword | ConvertTo-SecureString -Key $key

    # RUN DOWNLOAD TEST
    Receive-Winscp -Server $server -Port $port -Fingerprint $serverkey -Path $path -FileMask $filemask -Destination $destination -Username $user -SecurePassword $SecurePassword

}
TestReceive-Winscp