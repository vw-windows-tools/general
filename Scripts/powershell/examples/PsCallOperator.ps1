# Set path to executable
$ExePath = "C:\Some\Program\File.exe"
$ArchiveName = "C:\Backup Directory\Monday.7z"
$SourceFiles = "C:\Important Data\*.*"

# Set parameters for executable
[System.Collections.ArrayList]$AllArgs = @('a', '-y', $ArchiveName, $SourceFiles)

# Add parameter if necessary
if ($Delete) {
    $AllArgs.Add("-del")
    Write-Verbose "Delete mode ON"
}

& $ExePath $AllArgs