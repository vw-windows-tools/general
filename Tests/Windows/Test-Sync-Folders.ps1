#Region User-defined parameters
$ProjectPath = "C:\Users\John\Documents\Projects\myProject"
$ModuleFileFullPath = "$ProjectPath\Source Code\Windows.ps1"
$ReferencePath = "\\some\serverpath"
$TargetPath = "e:\destpath"
$ExcludeFolders = "Z001,Z002,Z010,Z011,Z037,DUMPZ"
$DebugPreference="Continue"
#Endregion

# (re)Import Test Module
$ModuleName = [io.path]::GetFileNameWithoutExtension($ModuleFileFullPath)
Get-Module | Where-Object -Property "Name" -Like "*$ModuleName*" | Remove-Module
Import-Module $ModuleFileFullPath

# Call function
Sync-Folders -ErrorAction Stop -ReferenceDirPath $ReferencePath -TargetDirPath $TargetPath -ParentFoldersExcludeList $ExcludeFolders # -delete -whatif