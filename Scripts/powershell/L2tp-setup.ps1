﻿$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

Import-Module "$ScriptDir\stdio.ps1"
Import-Module "$ScriptDir\Windows.ps1"

function SetupVpnFromConfig
{

    param (
        [Parameter(Mandatory=$true)][string]$ConfigFile
    )

    $Config = Get-Content $ConfigFile | ConvertFrom-Json
    $VpnConName = $Config.PARAMETERS.VpnConName
    $VpnServerAddress = $Config.PARAMETERS.VpnServerAddress
    $PreSharedKey = $Config.PARAMETERS.PreSharedKey
    $CaCert = $Config.PARAMETERS.CaCert
    $VpnType = $Config.PARAMETERS.VpnType
    $DestinationNetworks = $Config.PARAMETERS.DestinationNetworks

    Switch($VpnType)
    { 
        "L2tp" {
            
            New-L2tpPskVpn -VpnConName $VpnConName -VpnServerAddress $VpnServerAddress -DestinationNetwork $DestinationNetworks -PreSharedKey $PreSharedKey
            
        }
        Default {Write-Host "$VpnType : not supported ($VpnConName)"} 
    }

}

# Clients Windows 10 Home
Set-EncapsulationContextPolicy

# Cerbere (keep default gateway)
SetupVpnFromConfig -ConfigFile "$ScriptDir\L2tp (custom routes).json"

# Cerbere (route all packets)
SetupVpnFromConfig -ConfigFile "$ScriptDir\L2tp.json"

# Reboot recommended for Windows 10 Home
Read-String -Value "Y" -Message "Do you want to reboot now ? [Y/N]" -Action "Restart-Computer"