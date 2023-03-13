function Convert-XmlElementToString
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $xml
    )

    $sw = [System.IO.StringWriter]::new()
    $xmlSettings = [System.Xml.XmlWriterSettings]::new()
    $xmlSettings.ConformanceLevel = [System.Xml.ConformanceLevel]::Fragment
    $xmlSettings.Indent = $true
    $xw = [System.Xml.XmlWriter]::Create($sw, $xmlSettings)
    $xml.WriteTo($xw)
    $xw.Close()
    return $sw.ToString()
}

<#
.SYNOPSIS
Convert Base64 string to Plain text

.DESCRIPTION
Converts Base64 string to Plain text, with (optionnally) specified encoding

.PARAMETER Base64String
Source Base64String to convert to plain text

.PARAMETER Encoding
Converted output string format. Default is UTF8

.EXAMPLE
Convert-FromBase64 -Base64String 'SGVsbG8gV29ybGQ=' -Encoding "Unicode"

.EXAMPLE
Get-content $DataFilePath -raw | Convert-FromBase64 | ConvertFrom-Json
#>
function Convert-FromBase64 {

    param (
        [parameter(Mandatory=$True, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]

        [string]$Base64String,

        [parameter(Mandatory=$False)]
        [ValidateSet("Ascii","Default","Oem","Unicode","UTF32","UTF7","UTF8")]
        [Alias('e','f')]
        [string]$Encoding="Utf8"
    )

    try {
        $PlainText = [Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($Base64String))
    }
    catch {
        Throw
    }

    Return $PlainText

}

<#
.SYNOPSIS
Convert plain text to Base64

.DESCRIPTION
Convert plain text to Base64 string, with (optionnally) specified encoding

.PARAMETER PlainText
Source string to convert to Base64 string

.PARAMETER Encoding
Converted output string format. Default is UTF8

.EXAMPLE
Convert-ToBase64 -PlainText 'Hello World' -Encoding "utf8"

.EXAMPLE
'Hello World' | Convert-ToBase64
#>
function Convert-ToBase64 {

    param (
        [parameter(Mandatory=$True, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$PlainText,

        [parameter(Mandatory=$False)]
        [ValidateSet("Ascii","Default","Oem","Unicode","UTF32","UTF7","UTF8")]
        [Alias('e','f')]
        [string]$Encoding="Utf8"
    )

    try {
        $Base64String = [System.Convert]::ToBase64String([System.Text.Encoding]::$($Encoding).GetBytes($PlainText))
    }
    catch {
        Throw
    }

    Return $Base64String

}

function IsValidUri {
    param (
        [parameter(Mandatory=$True, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri
    )

    Return ($Uri -as [System.URI]).AbsoluteURI -ne $null

}

function IsValidUrl {
    param (
        [parameter(Mandatory=$True, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$Url
    )

    $Uri = $Url -as [System.URI]
	Return $null -ne $Uri.AbsoluteURI -and $Uri.Scheme -match '[http|https]'

}

function IsValidEmail
{
    param (
    [parameter(Mandatory=$True, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$Address
    )

        $EmailRegex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'

        Return ($Address -match $EmailRegex)

}

function IsValidADUsername
{
    param (
    [parameter(Mandatory=$True, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$UserName
    )

    Return $UserName -match "[$([Regex]::Escape('/\[:;|=,+*?<>') + '\]' + '\"')]"

}

function IsValidPhoneNumber
{
    param (
    [parameter(Mandatory=$True, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$Number
    )

        $PhoneRegex = '^(\+\d{8,20})|(\d{8,20})$'

        Return ($Number -match $PhoneRegex)
}