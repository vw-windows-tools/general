# Custom values
$username = "CONTOSO\John"
$clearpassword = "myp@55w0rd"
$server='test-srv03.contoso.com'
$source='\\shared\downloads\example-one.cfg'
$destination = "d:\configs\examples"

# Create credentials
$securePassword = ConvertTo-SecureString -String $ClearPassword -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential($username,$SecurePassword)

# Create remote session
$RemoteSession = New-PSSession -ComputerName $server -Credential $mycreds -ErrorAction Stop

# Create remote directory using a scriptblock
Invoke-command -Session $RemoteSession -ScriptBlock {

    $source = "$using:source"
    $destination = "$using:destination"

    $Error.Clear()
    New-Item -Path $destination -ItemType "directory" -ErrorAction SilentlyContinue

    If ($Null -ne $Error[0]) {
        If ($Error[0].CategoryInfo.Category -eq "ResourceExists") {
            Write-Warning "Directory already exists"
        }
        Else {
            Throw $error[0]
        }
    }

}

# Copy file using Copy-Item + 'ToSession' parameter
Copy-Item -Path $source -Destination $destination -ToSession $RemoteSession -Force -ErrorAction Stop