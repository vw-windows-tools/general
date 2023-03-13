#Region Modules
# Using string.ps1 function(s) : IsValidUrl, IsValidEmail, IsValidADUsername
$ModuleFileFullPath = "$PSScriptRoot\string.ps1"
$ModuleName = [io.path]::GetFileNameWithoutExtension($ModuleFileFullPath)
Get-Module | Where-Object -Property "Name" -Like "*$ModuleName*" | Remove-Module
Import-Module $ModuleFileFullPath

# Using string.ps1 function(s) : IsValidWindowsFile
$ModuleFileFullPath = "$PSScriptRoot\Windows.ps1"
$ModuleName = [io.path]::GetFileNameWithoutExtension($ModuleFileFullPath)
Get-Module | Where-Object -Property "Name" -Like "*$ModuleName*" | Remove-Module
Import-Module $ModuleFileFullPath
#Endregion

#Region Functions
<#
.SYNOPSIS
Create new EWS ExchangeService object

.DESCRIPTION
Create new ExchangeService object using Exchange Web Service API

.PARAMETER WebServiceUrl
Full Url to Exchange web service. If not specified, Office 365 web service URL is used

.PARAMETER WebServiceDll
Full path to Microsoft.Exchange.WebServices.dll. If not specified, function will try to find it using a default list

.PARAMETER UserName
Exchange user name (e.g "john@contoso.com")

.PARAMETER ClearPassword
Exchange user password, as plain text (not recommended)

.PARAMETER SecurePassword
Exchange user password, as securestring (recommended)

.EXAMPLE
$exchServ = New-ExchangeService -WebServiceUrl https://owa.contoso.com/EWS/Exchange.asmx -WebServiceDll $EWS_DllPath -UserName "john@contoso.com" -SecurePassword $SecPass

.OUTPUTS
Microsoft.Exchange.WebServices.Data.ExchangeServiceBase type object
#>
function New-ExchangeService
{

    Param(

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if ((IsValidUrl -Url $_) -eq $True) {
                $True
            }
            else {
                Throw "WebServiceUrl must be a valid URL"
            }
        })]
        [Alias('url')]
        [string]$WebServiceUrl,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (IsValidWindowsFile -Path $_) {
                $True
            }
            else {
                Throw "DLL not found $($_)"
            }
        })]
        [Alias('dll')]
        [string]$WebServiceDll,

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (((IsValidEmail -Address $_) -eq $False) -and ((IsValidADUsername -UserName $_) -eq $False)) {
                Throw "Must be a valid email address of Active Directory user name"
            }
            else {
                $True
            }
        })]
        [Alias('u')]
        [string]$UserName,

        [parameter(Mandatory=$True, ParameterSetName='ClearPassword')]
        [ValidateNotNullOrEmpty()]
        [Alias('cp')]
        [string]$ClearPassword,

        [parameter(Mandatory=$True, ParameterSetName='SecurePassword')]
        [ValidateNotNullOrEmpty()]
        [Alias('sp')]
        [SecureString]$SecurePassword

    )

    Try{

        # Set secure string when clear password is supplied
        if ($ClearPassword) {
            Write-Warning "Using ClearPassword parameter. It is recommended to use SecurePassword instead, otherwise it potentially exposes sensitive information"
            $SecurePassword = $ClearPassword | ConvertTo-SecureString -AsPlainText -Force
        }

        if ([string]::IsNullOrEmpty($WebServiceDll)){
            Write-Verbose "Path to EWS Managed Api library not specified, trying to find it..."
            if (Test-Path -Path 'C:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll') {$WebServiceDll = 'C:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll'}
            elseif (Test-Path -Path 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll') {$WebServiceDll = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'}
            elseif (Test-Path -Path 'C:\Program Files\PackageManagement\NuGet\Packages\Exchange.WebServices.Managed.Api.2.2.1.2\lib\net35\Microsoft.Exchange.WebServices.dll') {$WebServiceDll = 'C:\Program Files\PackageManagement\NuGet\Packages\Exchange.WebServices.Managed.Api.2.2.1.2\lib\net35\Microsoft.Exchange.WebServices.dll'}
            else {Throw "Path to EWS Managed Api library not found, please specify it using -WebServiceDll parameter"}
        }
        
        # Load Exchange Web Services API
        Import-Module $WebServiceDll

        # Create EWS object
        $ExchangeService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

        # Credentials
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
        $ExchangeService.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($cred)

        # Set default Office 365 Url for Exchange Web Service if not specified
        if ([string]::IsNullOrEmpty($WebServiceUrl)){$WebServiceUrl = "https://outlook.office365.com/EWS/Exchange.asmx"}
        $ExchangeService.Url= new-object Uri($WebServiceUrl)

    }
    catch {
        Throw
    }

    return $ExchangeService

}

<#
.SYNOPSIS
Create new mail folder

.DESCRIPTION
Creates new Exchange mail folder using EWS Managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER NewFolderDisplayName
Name of the folder to create

.PARAMETER ParentFolderPath
Full path to parent folder into which to create the new folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"

.PARAMETER ParentFolderObject
Exchange.WebServices.Data.Folder type object can by specified instead of ParentFolderPath

.EXAMPLE
$newFolderId = New-ExchangeMailFolder -ExchangeService $exchService -NewFolderDisplayName "folder01" -ParentFolderPath "inbox\tests"

.OUTPUTS
Unique Id of the successfully created folder
#>
function New-ExchangeMailFolder {

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('name')]
        [string]$NewFolderDisplayName,
        
        [Parameter(ParameterSetName='Id', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('i','id')]
        [string]$ParentFolderId,
        
        [Parameter(ParameterSetName='Path', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('p', 'path')]
        [string]$ParentFolderPath,
        
        [Parameter(ParameterSetName='Object', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('o', 'obj', 'object')]
        $ParentFolderObject

    )

    try {

        If ($ParentFolderId) {
            Write-Verbose "Get Parent folder object from Id"
            $ParentFolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $ParentFolderId
        }elseif ($ParentFolderPath) {
            Write-Verbose "Get Parent folder object from Path"
            $ParentFolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $ParentFolderPath
        }elseif (($ParentFolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied Parent object is not a Folder"
        }

        Write-Verbose "Create new folder object"
        $NewFolderObject = [Microsoft.Exchange.WebServices.Data.Folder]::new($ExchangeService)
        $NewFolderObject.DisplayName = $NewFolderDisplayName

        if ($PSCmdlet.ShouldProcess("folder '$($ParentFolderObject.DisplayName)'", "Save new folder '$NewFolderDisplayName'")) {
            Write-Verbose "Save new folder '$NewFolderDisplayName' under folder '$($ParentFolderObject.DisplayName)'"
            $NewFolderObject.Save($ParentFolderObject.id)
        }
        
    }
    catch {
        Throw
    }
    
    Return $NewFolderObject.Id.UniqueId

}

<#
.SYNOPSIS
Get EWS "Folder" type object

.DESCRIPTION
Returns an object of the Exchange.WebServices.Data.Folder type, for specified path, using Exchange Web Service API

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER FolderId
Exchange Id of folder.

.PARAMETER FolderPath
Full path to folder. Could be supplied instead of folder Exchange Id. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"

.PARAMETER ParentFolder
Optional Exchange.WebServices.Data.Folder type object. If supplied, FolderPath will be read starting from this folder instead of Root folder

.EXAMPLE
$archives = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath "inbox\Archives"

.OUTPUTS
Exchange.WebServices.Data.Folder type object
#>
function Get-ExchangeMailFolder
{

    Param(

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [Parameter(Mandatory=$True, ParameterSetName='Id')]
        [ValidateNotNullOrEmpty()]
        [Alias('i','id')]
        [string]$FolderId,

        [parameter(Mandatory=$True, ParameterSetName='Path')]
        [ValidateNotNullOrEmpty()]
        [Alias('p','path')]
        [string]$FolderPath,

        [parameter(Mandatory=$False, ParameterSetName='Path')]
        [ValidateNotNullOrEmpty()]
        [object]$ParentFolder

    )

    try {

        if ($PSCmdlet.ParameterSetName -eq 'Id') {
            Write-Verbose "Get Exchange mail folder by Id $($FolderId)"
            $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService, (New-Object Microsoft.Exchange.WebServices.Data.FolderId($FolderId)))
            Return $Folder
        }
        
        # Split path
        [System.Collections.ArrayList]$FolderPathArray = $FolderPath.split("\")

        # Get parent folder
        $FolderView = New-object Microsoft.Exchange.WebServices.Data.FolderView -ArgumentList 100

        # Get first level folder name
        $FolderDisplayName = $FolderPathArray[0]

        If ($Null -eq $ParentFolder) {
            
            Write-Verbose "No parent folder supplied, trying $FolderDisplayName"
            try {

                # Get Root Folder id
                $RootFolderId = ([Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)).id
                Write-Verbose "MsgFolderRoot id = '$($RootFolderId.UniqueId)'"
    
                # Get main subfolders
                $SubFolders = $ExchangeService.FindFolders($RootFolderId, $FolderView)
                Write-Verbose "Main sub-folders count : $($SubFolders.TotalCount)"

            }
            catch {
                Throw "Error getting top level folders : $($_.exception.message)"
            }
        }elseif (($ParentFolder.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied object is not a Folder"
        }
        else {
            Write-Verbose "Parent folder supplied, trying to get subfolders"
            $SubFolders = $ExchangeService.FindFolders($ParentFolder.Id, $FolderView)
        }

        # Get searched folder
        $Folder = $SubFolders | Where-Object {$_.DisplayName -eq $FolderDisplayName}

        # Check result
        if ($Null -eq $Folder) {
            Throw "Folder not found : $FolderDisplayName"
        }
        Write-Verbose "'$($FolderDisplayName)' id : $($Folder.id.UniqueId)"

        if ($FolderPathArray.count -eq 1) {
            # FolderPath contains no subfolders, returning supplied folder name at top level of mailbox
            Return $Folder
        }
        else {
            # Get next subfolder
            $FolderPathArray.RemoveAt(0) # remove top folder from list
            $NextFolderPath = $FolderPathArray -Join "\" # reassemble FolderPath parameter string
            Write-Verbose "Recursive call for next folders : '$($NextFolderPath)'"
            try {
                $Folder = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $NextFolderPath -ParentFolder $Folder
            }
            catch {
                Throw "Recursive call error : $($_.exception.message)"
            }
            Return $Folder
        }

    }
    catch {
        Throw
    }

}

<#
.SYNOPSIS
Get Exchange Mail folder's sub-folders

.DESCRIPTION
Gets all sub-folders for specified Exchange Mail folder using EWS Managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER FolderId
Exchange Id of folder. Incompatible with FolderPath and FolderObject parameters

.PARAMETER FolderPath
Full path to source folder to move. Incompatible with FolderId and FolderObject parameters. Separate folders with Antislashes ("\"). E.g : "Inbox\Tests"

.PARAMETER FolderObject
Exchange.WebServices.Data.Folder type object. Incompatible with FolderId and FolderPath parameters.

.EXAMPLE
$subfolders = Get-ExchangeMailSubFolders -ExchangeService $exchserv -FolderId "AAMkAGQ5MWNkN2Q3LWE5N..."

.EXAMPLE
$subfolders = Get-ExchangeMailSubFolders -ExchangeService $exchserv -FolderPath "deleted items"

.EXAMPLE
$subfolders = Get-ExchangeMailSubFolders -ExchangeService $exchserv -FolderObject (Get-ExchangeMailFolder -ES $TestExchService -FolderPath inbox)

.OUTPUTS
Array of Microsoft.Exchange.WebServices.Data.Folder objects, or $null if no subfolder found
#>
function Get-ExchangeMailSubFolders
{

    Param(
  
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [Parameter(Mandatory=$True, ParameterSetName='Id')]
        [ValidateNotNullOrEmpty()]
        [Alias('i','id')]
        [string]$FolderId,

        [parameter(Mandatory=$True, ParameterSetName='Path')]
        [ValidateNotNullOrEmpty()]
        [Alias('p','path')]
        [string]$FolderPath,

        [parameter(Mandatory=$True, ValueFromPipeline, ParameterSetName='Object')]
        [Alias('o', 'obj', 'object')]
        [ValidateNotNullOrEmpty()]
        $FolderObject

    )

    If ($FolderId) {
        Write-Verbose "Get folder object from Id"
        $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $FolderId
    }elseif ($FolderPath) {
        Write-Verbose "Get folder object from Path"
        $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $FolderPath
    }elseif (($FolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
        Throw "Supplied object is not a Folder"
    }

    try {
        $FolderView = New-object Microsoft.Exchange.WebServices.Data.FolderView -ArgumentList 100
        $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        $SubFolders = ($ExchangeService.FindFolders($FolderObject.id, $FolderView))
    }
    catch {
        Throw
    }
    
    Return $SubFolders

}

<#
.SYNOPSIS
Move Exchange folder

.DESCRIPTION
Moves source exchange folder into (under) destination folder.

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER SourceFolderId
Exchange Id of folder. Incompatible with SourceFolderPath and SourceFolderObject parameters

.PARAMETER SourceFolderPath
Full path to source folder to move. Incompatible with SourceFolderId and SourceFolderObject parameters. Separate folders with Antislashes ("\"). E.g : "Inbox\Tests"

.PARAMETER SourceFolderObject
Exchange.WebServices.Data.Folder type object. Incompatible with SourceFolderId and SourceFolderPath parameters.

.PARAMETER DestinationFolderId
Exchange Id of folder. Incompatible with DestinationFolderPath and DestinationFolderObject parameters

.PARAMETER DestinationFolderPath
Full path to Destination folder to move. Incompatible with DestinationFolderId and DestinationFolderObject parameters. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"

.PARAMETER DestinationFolderObject
Exchange.WebServices.Data.Folder type object. Incompatible with DestinationFolderId and DestinationFolderPath parameters.

.EXAMPLE
Move-ExchangeMailFolder -SourceFolderPath "Inbox\Tests" -DestinationFolderPath "Inbox\Archives"-ExchangeService $exchService

.EXAMPLE
Move-ExchangeMailFolder -SourceFolderObject $SourceFolder -DestinationFolderId "AAMkAGQ5MWNkN2Q3LWE5N..." -ExchangeService $exchService

.OUTPUTS
$True if folder is successfully moved to new location
#>
function Move-ExchangeMailFolder
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [Alias('service','es')]
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        ## 'SourceFolder' group (SourceFolderId, SourceFolderPath and SourceFolderObject are mutually exclusive)

        [Parameter(ParameterSetName='SourceId-DestId', Mandatory=$True)]
        [Parameter(ParameterSetName='SourceId-DestPath', Mandatory=$True)]
        [Parameter(ParameterSetName='SourceId-DestObject', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('sid','sourceid')]
        [string]$SourceFolderId,

        [Parameter(ParameterSetName='SourcePath-DestId', Mandatory=$True)]
        [Parameter(ParameterSetName='SourcePath-DestPath', Mandatory=$True)]
        [Parameter(ParameterSetName='SourcePath-DestObject', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('spath','sourcepath')]
        [string]$SourceFolderPath,

        [Parameter(ParameterSetName='SourceObject-DestId', ValueFromPipeline, Mandatory=$True)]
        [Parameter(ParameterSetName='SourceObject-DestPath', ValueFromPipeline, Mandatory=$True)]
        [Parameter(ParameterSetName='SourceObject-DestObject', ValueFromPipeline, Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('so','sourceobj','sourceobject')]
        [object]$SourceFolderObject,

        ## 'DestinationFolder' group (DestinationFolderId, DestinationFolderPath and DestinationFolderObject are mutually exclusive)

        [Parameter(ParameterSetName='SourceId-DestId', Mandatory=$True)]
        [Parameter(ParameterSetName='SourcePath-DestId', Mandatory=$True)]
        [Parameter(ParameterSetName='SourceObject-DestId', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('did','destid')]
        [string]$DestinationFolderId,

        [Parameter(ParameterSetName='SourceId-DestPath', Mandatory=$True)]
        [Parameter(ParameterSetName='SourcePath-DestPath', Mandatory=$True)]
        [Parameter(ParameterSetName='SourceObject-DestPath', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('dpath','destpath')]
        [string]$DestinationFolderPath,

        [Parameter(ParameterSetName='SourceId-DestObject', Mandatory=$True)]
        [Parameter(ParameterSetName='SourcePath-DestObject', Mandatory=$True)]
        [Parameter(ParameterSetName='SourceObject-DestObject', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('do','destobj','destobject')]
        [object]$DestinationFolderObject

    )

    try {

        If ($SourceFolderId) {
            Write-Verbose "Get source folder object from Id"
            $SourceFolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $SourceFolderId
        }elseif ($SourceFolderPath) {
            Write-Verbose "Get source folder object from Path"
            $SourceFolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $SourceFolderPath
        }elseif (($SourceFolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied source object is not a Folder"
        }

        If ($DestinationFolderId) {
            Write-Verbose "Get destination folder object from Id"
            $DestinationFolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $DestinationFolderId
        }elseif ($DestinationFolderPath) {
            Write-Verbose "Get destination folder object from Path"
            $DestinationFolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $DestinationFolderPath
        }elseif (($DestinationFolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied destination object is not a Folder"
        }

        if ($PSCmdlet.ShouldProcess("$($DestinationFolderObject.DisplayName)", "Move folder $($SourceFolderObject.DisplayName)")) {
            Write-Verbose "Move folder $($SourceFolderObject.DisplayName) to $($DestinationFolderObject.DisplayName)"
            $SourceFolderObject.Move($DestinationFolderObject.Id) | Out-Null # Out-Null used here not to go into pipeline
        }

    }
    catch {
        Throw
    }

    Return $True

}

<#
.SYNOPSIS
Rename Exchange mail folder

.DESCRIPTION
Renames Exchange mail folder using EWS Managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER FolderNewName
New name of the folder to update

.PARAMETER FolderId
Exchange Id of folder. Incompatible with FolderPath and FolderObject parameters

.PARAMETER FolderPath
Full path to folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests". Incompatible with FolderId and FolderId parameters

.PARAMETER FolderObject
Exchange.WebServices.Data.Folder type object. Incompatible with FolderPath and FolderId parameters

.EXAMPLE
$Result = Rename-ExchangeMailFolder -ExchangeService $exchService -FolderPath $FolderPath -FolderNewName $NewName

.OUTPUTS
$True if folder is successfully renamed
#>
function Rename-ExchangeMailFolder {

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [parameter(Mandatory=$True)]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('n', 'newname')]
        [string]$FolderNewName,

        [Parameter(ParameterSetName='Id', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('i', 'id')]
        [string]$FolderId,

        [Parameter(ParameterSetName='Path', Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('p', 'path')]
        [string]$FolderPath,

        [Parameter(ParameterSetName='Object', ValueFromPipeline, Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('o', 'obj','object')]
        [object]$FolderObject

    )

    try {

        If ($FolderId) {
            Write-Verbose "Get folder object from Id"
            $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $FolderId
        }elseif ($FolderPath) {
            Write-Verbose "Get folder object from Path"
            $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $FolderPath
        }elseif (($FolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied object is not a Folder"
        }

        Write-Verbose "Rename folder '$($FolderObject.DisplayName)' to '$FolderNewName'"
        $FolderObject.DisplayName = $FolderNewName
        
        if ($PSCmdlet.ShouldProcess("'$FolderNewName'", "Rename folder '$($FolderObject.DisplayName)'")) {
            $FolderObject.Update() | Out-Null # Out-Null used here not to go into pipeline
        }
        
    }
    catch {
        Throw
    }

    Return $True

}

<#
.SYNOPSIS
Delete Exchange Folder

.DESCRIPTION
Deletes Exchange folder using Exchange Web Services API. Full path (e.g "inbox\archives") or EWS Folder object can be specified.

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER FolderId
Exchange Id of folder. Incompatible with FolderPath and FolderObject parameters

.PARAMETER FolderPath
Full path to folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests". Incompatible with FolderId and FolderId parameters

.PARAMETER FolderObject
Exchange.WebServices.Data.Folder type object. Incompatible with FolderPath and FolderId parameters

.EXAMPLE
Remove-ExchangeMailFolder -ExchangeService $exs -FolderId "AAMkAGQ5MWNkN2Q3LWE5N..."

.EXAMPLE
Remove-ExchangeMailFolder -ExchangeService $exs -FolderPath "inbox\archives\john" -Hard

.EXAMPLE
Remove-ExchangeMailFolder -ExchangeService $exs -FolderObject $myFolder -Soft

.OUTPUTS
$True if folder(s) removed successfully

.NOTES
The standard and -Hard options are transactional, which means that by the time a web service call completes, the database has moved the item to the Deleted Items folder or permanently removed the item from the Exchange database. The -Soft delete option works differently for different target versions of Exchange Server. Soft Delete for Exchange 2007 sets a bit on the item that indicates to the Exchange database that the item will be moved to the dumpster folder at an indeterminate time in the future. Soft Delete for versions of Exchange starting with Exchange 2010, including Exchange Online, immediately moves the item to the dumpster. Soft Delete is not an option for folder deletion. Soft Delete traversal searches for items and folders will not return any results. On a side note, a folder that previously contained an item (folder, mail) that was not hard deleted (so it's still in deleted items or dumpster) can only be deleted using the "Hard" delete mode.
#>
function Remove-ExchangeMailFolder
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [Parameter(Mandatory=$True, ParameterSetName='Id')]
        [ValidateNotNullOrEmpty()]
        [Alias('i', 'id')]
        [string]$FolderId,

        [Parameter(Mandatory=$True, ParameterSetName='Path')]
        [ValidateNotNullOrEmpty()]
        [Alias('p', 'path')]
        [string]$FolderPath,

        [Parameter(Mandatory=$True, ValueFromPipeline, ParameterSetName='Object')]
        [ValidateNotNullOrEmpty()]
        [Alias('o','obj','object')]
        [object]$FolderObject,

        [ValidateSet("Hard","Soft","Default")]
        [ValidateNotNullOrEmpty()]
        [Alias('d', 'dm')]
        [string]$DeleteMode="Default",

        [parameter(Mandatory=$False)]
        [Alias('r')]
        [switch]$Recurse

    )

    try {

        If ($FolderId) {
            Write-Verbose "Get folder object from Id"
            $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $FolderId
        }elseif ($FolderPath) {
            Write-Verbose "Get folder object from Path"
            $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $FolderPath
        }elseif (($FolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied object is not a Folder"
        }

        if ($Recurse) {

            Write-Verbose "Parameter -Recurse specified, getting all additional subfolders to delete for folder '$($FolderObject)'"
            $SubFolders = Get-ExchangeMailSubFolders -ExchangeService $ExchangeService -FolderObject $FolderObject
            $AllFolders = New-Object System.Collections.ArrayList($null)
            $AllFolders.Add($SubFolders) # Add all subfolders to delete list
            $AllFolders.Add($FolderObject) # Add specified folder to delete list

            $preventInfiniteLoop = 0
            $maxLoopCount = 50
            while ((($AllFolders | Measure-Object).count -gt 0) -and ($preventInfiniteLoopCounter -lt $maxLoopCount)) {
                # Loop on all AllFolders. A folder can't be deleted as long as it's got old AllFolders in it.
                # Hierarchy not being taken into account, the purpose of this loop is to repeat delete tries until AllFolders list is empty.
                $preventInfiniteLoop++
                foreach ($folder in $AllFolders) {
                    try {
                        if ($PSCmdlet.ShouldProcess("Folder '$($folder.DisplayName)'", "Delete item (mode = '$($DeleteMode)'")) {
                            Write-Verbose "Delete folder '$($folder.DisplayName)' (mode = '$($DeleteMode)')"
                            Remove-ExchangeItem -ExchangeService $ExchangeService -ExchangeItem $folder -DeleteMode $DeleteMode | Out-Null # Out-Null used here not to go into pipeline $folder
                        }
                        $AllFolders = $AllFolders | Where-Object {$_ -ne $folder}
                    }
                    catch {
                        Write-Verbose "Exception caught deleting subfolder '$($folder.DisplayName)' : $($_.exception.message)"
                    } 
                }
            }

        }else {
            if ($PSCmdlet.ShouldProcess("Folder '$($FolderObject.DisplayName)'", "Delete item (mode = $DeleteMode)")) {
                Write-Verbose "Delete folder '$($FolderObject.DisplayName)' (mode = $DeleteMode)"
                Remove-ExchangeItem -ExchangeService $ExchangeService -ExchangeItem $FolderObject -DeleteMode $DeleteMode | Out-Null # Out-Null used here not to go into pipeline
            }
        }
        
    }
    catch {
        Throw
    }
    
    Return $True

}

<#
.SYNOPSIS
Clear Exchange Mail folder

.DESCRIPTION
Clears a folder using Exchange Web Services Managed Api. Erases all emails within specified folder. Applies to subfolders with -recurse parameter

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER FolderId
Exchange Id of folder. Incompatible with FolderPath and FolderObject parameters

.PARAMETER FolderPath
Full path to folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests". Incompatible with FolderId and FolderId parameters

.PARAMETER FolderObject
Exchange.WebServices.Data.Folder type object. Incompatible with FolderPath and FolderId parameters

.PARAMETER DeleteMode
Optional. "Default" behaviour is to move all erased emails to the mailbox's Deleted Items folder. "Soft" will move them to the dumpster (items in the dumpster can be recovered). "Hard" will permanently delete the emails. 

.PARAMETER Recurse
Optional. Deletes all mails in specified folder + all mails in subfolders

.EXAMPLE
Clear-ExchangeMailFolder -ExchangeService $exchService -FolderPath "inbox\archives\old"

.EXAMPLE
Clear-ExchangeMailFolder -ExchangeService $exchService -FolderPath "inbox\archives" -recurse

.EXAMPLE
Clear-ExchangeMailFolder -ExchangeService $exchService -FolderObject (Get-ExchangeMailFolder -FolderPath "inbox\archives\old")

.OUTPUTS
$True if folder(s) cleared successfully
#>
function Clear-ExchangeMailFolder {

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [Parameter(Mandatory=$True, ParameterSetName='Id')]
        [ValidateNotNullOrEmpty()]
        [Alias('i', 'id')]
        [string]$FolderId,

        [Parameter(Mandatory=$True, ParameterSetName='Path')]
        [ValidateNotNullOrEmpty()]
        [Alias('p', 'path')]
        [string]$FolderPath,

        [Parameter(Mandatory=$True, ValueFromPipeline, ParameterSetName='Object')]
        [ValidateNotNullOrEmpty()]
        [Alias('o','obj','object')]
        [object]$FolderObject,

        [ValidateSet("Hard","Soft","Default")]
        [ValidateNotNullOrEmpty()]
        [Alias('d','dm')]
        [string]$DeleteMode="Default",

        [parameter(Mandatory=$False)]
        [Alias('r')]
        [switch]$Recurse

    )

    function Clear-AllMailsInFolder {
        [CmdletBinding()]
        param (
            [Parameter()]$FolderToClear
        )
        try {
            $AllMailsInFolder = Get-ExchangeMail -ExchangeService $ExchangeService -FolderObject $FolderToClear
            foreach ($mail in $AllMailsInFolder) {
                Remove-ExchangeItem -ExchangeService $ExchangeService -ExchangeItem $mail -DeleteMode $DeleteMode
            }
        }
        catch {
            Throw
        }
    }

    try {
        
        If ($FolderId) {
            Write-Verbose "Get folder object from Id"
            $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $FolderId
        }
        elseif ($FolderPath) {
            Write-Verbose "Get folder object from Path"
            $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $FolderPath
        }elseif (($FolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied object is not a Folder"
        }

        if ($Recurse) {

            Write-Verbose "Parameter -Recurse specified, getting all additional subfolders to clear for folder '$($FolderObject)'"
            $SubFolders = Get-ExchangeMailSubFolders -ExchangeService $ExchangeService -FolderObject $FolderObject
            $AllFolders = New-Object System.Collections.ArrayList($null)
            $AllFolders.Add($SubFolders) # Add all subfolders to clearing list
            $AllFolders.Add($FolderObject) # Add specified folder to clearing list

            $preventInfiniteLoop = 0
            $maxLoopCount = 50
            while ((($AllFolders | Measure-Object).count -gt 0) -and ($preventInfiniteLoopCounter -lt $maxLoopCount)) {
                # Loop on all AllFolders. A folder can't be deleted as long as it's got old AllFolders in it.
                # Hierarchy not being taken into account, the purpose of this loop is to repeat delete tries until AllFolders list is empty.
                $preventInfiniteLoop++
                foreach ($folder in $AllFolders) {
                    try {
                        if ($PSCmdlet.ShouldProcess("folder '$($folder.DisplayName)'", "Clear all mails")) {
                            Clear-AllMailsInFolder -FolderToClear $folder | Out-Null # Out-Null used here not to go into pipeline
                        }
                        $AllFolders = $AllFolders | Where-Object {$_ -ne $folder}
                    }
                    catch {
                        Write-Warning "Exception caught trying to clear subfolder '$($folder.DisplayName)' : $($_.exception.message)"
                    } 
                }
            }

        }else {
            if ($PSCmdlet.ShouldProcess("folder '$($FolderObject.DisplayName)'", "Clear all mails")) {
                Write-Verbose "Clear all mails in folder '$($FolderObject.DisplayName)'"
                Clear-AllMailsInFolder -FolderToClear $FolderObject
            }
        }

    }
    catch {
        Throw
    }

    Return $True

}

<#
.SYNOPSIS
Send Exchange mail

.DESCRIPTION
Sends an email using Exchange Web Service Managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER To
"Outlook-like" semi-colon separated list of email addresses (optional if Cc or Bcc is specified)

.PARAMETER Cc
"Outlook-like" semi-colon separated list of email addresses (optional if To or Bcc is specified)

.PARAMETER Bcc
"Outlook-like" semi-colon separated list of email addresses (optional if To or Cc is specified)

.PARAMETER Title
"Subject" of the email (mandatory, empty string accepted)

.PARAMETER Body
Mail body (optional, empty string accepted)

.PARAMETER BodyType
"Text" or "HTML". Set body type (optional, Default is HTML).

.PARAMETER Attachments
semi-colon separated list of files full paths

.PARAMETER Importance
(optional) Email priority set by sender, high, low, or normal

.EXAMPLE
$Return = Send-ExchangeMail -ExchangeService $exchserv -To "tom@contoso.com;john@domain.com" -Title 'Hi' -Body $mailbody -BodyType "text"

.EXAMPLE
Send-ExchangeMail -ExchangeService $exchserv -Attachments "c:\docs\file1.txt;c:\pictures\file2.jpg" -Bcc "tom@contoso.com" -Title "Hello" -Body "hello<br><br>world"

.OUTPUTS
$True if mail successfully sent

.NOTES
To, Cc and Bcc are all optional parameters, but one must be specified at least.
#>
function Send-ExchangeMail
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [Parameter(Mandatory, ParameterSetName='to')]
        [Parameter(Mandatory, ParameterSetName='to-cc')]
        [Parameter(Mandatory, ParameterSetName='to-bcc')]
        [Parameter(Mandatory, ParameterSetName='to-cc-bcc')]
        [ValidateNotNullOrEmpty()]
        [string]$To,
        
        [Parameter(Mandatory, ParameterSetName='cc')]
        [Parameter(Mandatory, ParameterSetName='to-cc')]
        [Parameter(Mandatory, ParameterSetName='cc-bcc')]
        [Parameter(Mandatory, ParameterSetName='to-cc-bcc')]
        [ValidateNotNullOrEmpty()]
        [String]$Cc,

        [Parameter(Mandatory, ParameterSetName='bcc')]
        [Parameter(Mandatory, ParameterSetName='to-bcc')]
        [Parameter(Mandatory, ParameterSetName='cc-bcc')]
        [Parameter(Mandatory, ParameterSetName='to-cc-bcc')]
        [ValidateNotNullOrEmpty()]
        [string]$Bcc,
        
        [parameter(Mandatory=$True)]
        [AllowEmptyString()]
        [Alias('t','s','subject')]
        [string]$Title,

        [parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [Alias('b')]
        [string]$Body,

        [parameter(Mandatory=$False)]
        [ValidateSet("Text","Html")]
        [Alias('bt','type')]
        [string]$BodyType="Html",

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('a','f','files')]
        [string]$Attachments,

        [parameter(Mandatory=$False)]
        [ValidateSet("Normal","High","Low")]
        [Alias('i','p','priority')]
        [string]$Importance="Normal"

    )

    try {

        Write-Verbose "Create the email message and set the Subject and Body"
        $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $ExchangeService
        $message.Subject = $Title
        $message.Body = $Body + "`r`n"
        $message.Body.BodyType = $BodyType

        Write-Verbose "Body type set to '$($BodyType)'"

        Write-Verbose "Set mail priority"
        switch ($Importance.tolower()) {
            "high" { $Importance = "High" ; break }
            "low" { $Importance = "Low" ; break }
            "normal" { $Importance = "Normal" ; break }
        }
        $message.Importance = $Importance

        if ($Attachments){
            Write-Verbose "Add attachments"
            # Split attachment string into array
            $AttachmentsList = $Attachments.Split(";");
            ForEach ($file in $AttachmentsList){
                # Check file path before attaching
                if (Test-Path $file){
                    $message.Attachments.AddFileAttachment($file) | Out-Null; # Out-Null used here not to go into pipeline
                    Write-Verbose "Added attachment $file"
                }
                else {
                    Throw "File not found $file --> $($_.Exception.Message)"            
                }
            }
        }

        # Split attachment string into array
        if (![string]::IsNullOrEmpty($To)){
            Write-Verbose "Add each specified 'To' recipient"
            $ToList = $To.Split(";");
            ForEach ($Recipient in $ToList)
            {
                $message.ToRecipients.Add($Recipient) | Out-Null # Out-Null used here not to go into pipeline
            }
        }

        # Split attachment string into array
        if (![string]::IsNullOrEmpty($Cc)){
            Write-Verbose "Add each specified 'Cc' recipient"
            $CcList = $Cc.Split(";");
            ForEach ($Recipient in $CcList)
            {
                $message.CcRecipients.Add($Recipient) | Out-Null # Out-Null used here not to go into pipeline
            }
        }

        # Split attachment string into array
        if (![string]::IsNullOrEmpty($Bcc)){
            Write-Verbose "Add each specified 'Bcc' recipient"
            $BccList = $Bcc.Split(";");
            ForEach ($Recipient in $BccList)
            {
                $message.BccRecipients.Add($Recipient) | Out-Null # Out-Null used here not to go into pipeline
            }
        }

        if ($PSCmdlet.ShouldProcess("Recipients", "Send the message (copy gets saved in sent items of the user)")) {
            Write-Verbose "Send the message (copy gets saved in sent items of the user)"
            $message.SendAndSaveCopy() | Out-Null # Out-Null used here not to go into pipeline
        }
        
    }

    catch {
        Throw
    }

    return $True

}

<#
.SYNOPSIS
Send Exchange mail reply

.DESCRIPTION
Sends an email reply using Exchange Web Services Managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER MailId
Get Email message by its unique Id. Incompatible with MailObject parameter.

.PARAMETER MailObject
Exchange Web Services Email object. Incompatible with MailId parameter. Could be retrieved with function Get-ExchangeMail

.PARAMETER ReplyString
Reply message to be sent ; can be empty, plain text or Html. History of previous email(s) will be kept.

.PARAMETER AddTo
Semi-colon separated list of recipients to add in To recipients

.PARAMETER AddCc
Semi-colon separated list of recipients to add in Cc recipients

.PARAMETER AddBcc
Semi-colon separated list of recipients to add in Bcc recipients

.PARAMETER Importance
(optional) Email priority set by sender, high, low, or normal

.PARAMETER Attachments
Semi-colon separated list of full path to file(s) to be added as attachment(s)

.PARAMETER ReplyToAll
If specified, all recipients of the initial email will be kept as new recipients. If omitted, only the sender will be kept. Can be associated in both cases with To, Cc, and Bcc parameters.

.PARAMETER Forward
(optional) Create 'Forward' response mail instead if Reply mail. At least one recipient must be specified with this option, using 'AddTo' parameter

.EXAMPLE
Send-ExchangeMailReply -ExchangeService $exchService -MailObject $mail -ReplyString "Hello<br><br>World" -Attachments "C:\users\john\doc\prices.pdf;c:\users\john\doc\john.vcf"

.EXAMPLE
Send-ExchangeMailReply -ExchangeService $exchService -MailId "AAMkAGQ5MWNkN2Q3LWE5N..." -ReplyString "Thanks everyone for your cooperation" -AddCc "ceo@contoso.com" -ReplyAll

.EXAMPLE
Send-ExchangeMailReply -es $es -Forward -ReplyString "" -AddTo 'supervision@contoso.com' -MailId "AAMkAGQ5MWNkN2Q3LWE5N..."

.OUTPUTS
$True if mail was successfully sent.
#>
function Send-ExchangeMailReply
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [parameter(Mandatory=$True)]
        [AllowEmptyString()]
        [Alias('r','rep')]
        [string]$ReplyString,

        [Parameter(Mandatory=$True, ParameterSetName="Id-Reply")]
        [Parameter(Mandatory=$True, ParameterSetName="Id-Forward")]
        [ValidateNotNullOrEmpty()]
        [Alias('i','id')]
        [string]$MailId,

        [Parameter(Mandatory=$True, ValueFromPipeline, ParameterSetName="Object-Reply")]
        [Parameter(Mandatory=$True, ValueFromPipeline, ParameterSetName="Object-Forward")]
        [ValidateNotNullOrEmpty()]
        [Alias('o','obj','object')]
        [object]$MailObject,

        [Parameter(Mandatory=$False, ParameterSetName="Object-Reply")]
        [Parameter(Mandatory=$False, ParameterSetName="Id-Reply")]
        [Parameter(Mandatory=$True, ParameterSetName="Object-Forward")]
        [Parameter(Mandatory=$True, ParameterSetName="Id-Forward")]
        [ValidateNotNullOrEmpty()]
        [Alias('to')]
        [string]$AddTo,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('cc')]
        [string]$AddCc,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('bcc')]
        [string]$AddBcc,

        [parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [Alias('t','s','subject')]
        [string]$Title,

        [parameter(Mandatory=$False)]
        [ValidateSet("Normal","High","Low")]
        [Alias('p','priority')]
        [string]$Importance,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('a','files')]
        [string]$Attachments,

        [parameter(Mandatory=$False)]
        [Alias('all')]
        [switch]$ReplyToAll,

        [Parameter(Mandatory=$True, ParameterSetName="Id-Forward")]
        [Parameter(Mandatory=$True, ParameterSetName="Object-Forward")]
        [parameter(Mandatory=$False)]
        [Alias('f','fw')]
        [switch]$Forward

    )

    try {

        if ($PSCmdlet.ParameterSetName -match 'Id') {
            $MailObject = Get-ExchangeMail -ExchangeService $ExchangeService -MailId $MailId
        }elseif (($MailObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.EmailMessage") {
            Throw "Supplied object is not an Email"
        }

        if ($Forward) {
            Write-Verbose "Create Forward"
            $responseMail = $MailObject.CreateForward()
        }else {
            Write-Verbose "Create Reply"
            $responseMail = $MailObject.CreateReply([boolean]$ReplyToAll)
        }

        $responseMail.BodyPrefix = $ReplyString

        if (!$Forward) {
            # Add at least the sender as To recipient when parameter -ReplyToAll is not specified, exception when -Forward is specified
            $responseMail.ToRecipients.Add($MailObject.Sender)
        }
        
        if ($ReplyToAll) {
            # Fix CreateReply method not working with bool ReplyAll set to True (set server-side ?)
            Write-Verbose "ReplyToAll specified, add all recipients to mail reply"
            foreach ($recipient in $MailObject.ToRecipients) {
                $responseMail.ToRecipients.Add($recipient)
            }
            foreach ($recipient in $MailObject.CcRecipients) {
                $responseMail.CcRecipients.Add($recipient)
            }
        }

        if (![string]::IsNullOrEmpty($AddTo)) {

            try {
                Write-Verbose "Add To recipients"
                $AddToRecipientsList = $AddTo.Split(";")
                foreach ($recipient in $AddToRecipientsList) {
                    $responseMail.ToRecipients.Add($recipient)
                }
            }catch {
                Throw "Exception caught adding To recipients : $($_.exception.message)"
            }

        }
        
        if (![string]::IsNullOrEmpty($AddCc)) {

            try {
                
                    Write-Verbose "Add Cc recipients"
                    $AddCcRecipientsList = $AddCc.Split(";")
                    foreach ($recipient in $AddCcRecipientsList) {
                        $responseMail.CcRecipients.Add($recipient)
                    }
                
            }
        
            catch {
                Throw "Exception caught adding Cc recipients : $($_.exception.message)"
            }

        }

        if (![string]::IsNullOrEmpty($AddBcc)) {

            try {
                
                    Write-Verbose "Add Bcc recipients"
                    $AddBccRecipientsList = $AddBcc.Split(";")
                    foreach ($recipient in $AddBccRecipientsList) {
                        $responseMail.BccRecipients.Add($recipient)
                    }
                
            }
            catch {
                Throw "Exception caught adding Bcc recipients : $($_.exception.message)"
            }

        }

        if (![string]::IsNullOrEmpty($Title)) {
            
            try {
                $responseMail.Subject = $Title
            }
            catch {
                Throw "Exception caught adding subject : $($_.exception.message)"
            }
            
        }

        if (![string]::IsNullOrEmpty($Importance) -or ![string]::IsNullOrEmpty($Attachments)) {

            Write-Verbose "Save response mail"
            $savedResponseMail = $responseMail.Save() # create new EmailMessage object because ResponseMessage class has neither AddAttachments method nor Importance property

            if (![string]::IsNullOrEmpty($Importance)) {
                Write-Verbose "Set mail priority"
                switch ($Importance.tolower()) {
                    "high" { $Importance = "High" ; break }
                    "low" { $Importance = "Low" ; break }
                    "normal" { $Importance = "Normal" ; break }
                }
                $savedResponseMail.Importance = $Importance
            }
            
            if (![string]::IsNullOrEmpty($Attachments)) {

                try {
                        
                    Write-Verbose "Add attachments"
                    
                    $AddAttachmentsList = $Attachments.Split(";")
                    ForEach ($file in $AddAttachmentsList){

                        # Check file path before attaching
                        if (Test-Path $file){
                            $savedResponseMail.Attachments.AddFileAttachment($file) | Out-Null; # Out-Null used here not to go into pipeline
                        }
                        else {
                            Throw "File not found $file"
                        }

                    }
                    
                }
                catch {
                    Throw "Exception caught adding attachment : $($_.exception.message)"
                }
        
            }

            try {
                if ($PSCmdlet.ShouldProcess("Recipients", "Send email reply with attachment(s)")) {
                    Write-Verbose "Send email reply with attachment(s)..."
                    $savedResponseMail.SendAndSaveCopy()
                    Write-Verbose "Email sent."
                }
            }
            catch {
                Throw "Exception caught sending reply email : $($_.exception.message)"
            }

        }
        else {

            try {
                if ($PSCmdlet.ShouldProcess("Recipients", "Send email reply without attachements and no importance set")) {
                    Write-Verbose "Send email reply without attachements..."
                    $responseMail.SendAndSaveCopy()
                    Write-Verbose "Email sent."
                }
            }
            catch {
                Throw "Exception caught sending reply email : $($_.exception.message)"
            }

        }

    }
    catch {
        Throw
    }

    Return $True

}

<#
.SYNOPSIS
Get Exchange Mails

.DESCRIPTION
Gets Exchange Mail(s) in specified Folder using EWS Managed Api, with several filters on recipients, dates, etc

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER MailId
Get Email message by its unique Id. Appart from ExchangeService, any other parameter (folder, filters..) becomes irrelevant with this one.

.PARAMETER FolderPath
Full path to Exchange mail folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"

.PARAMETER FolderObject
Exchange.WebServices.Data.Folder type object can by specified instead of FolderPath

.PARAMETER SentBefore
(optional) Minimum send date (as datetime, or as string, format "yyyy-MM-ddTHH:mm:ss") of the email(s)

.PARAMETER SentAfter
(optional) Maximum send date (as datetime, or as string, format "yyyy-MM-ddTHH:mm:ss") of the email(s)

.PARAMETER ReceivedBefore
(optional) Minimum receive date (as datetime, or as string, format "yyyy-MM-ddTHH:mm:ss") of the email(s)

.PARAMETER ReceivedAfter
(optional) Maximum send date (as datetime, or as string, format "yyyy-MM-ddTHH:mm:ss") of the email(s)

.PARAMETER From
(optional) Exact email address of the Sender

.PARAMETER To
(optional) Matching string in "To" recipients list (can be an full/partial email address or name)

.PARAMETER Cc
(optional) Matching string in "Cc" recipients list (can be an full/partial email address or name)

.PARAMETER Bcc
(optional) Matching string in "Bcc" recipients list (can be an full/partial email address or name)

.PARAMETER DisplayTo
(optional) Matching string in "To" recipients *names* list (not addresses)

.PARAMETER DisplayCc
(optional) Matching string in "Cc" recipients *names* list (not addresses)

.PARAMETER Subject
(optional, empty string allowed) Matching string in Subject if the email(s)

.PARAMETER Body
(optional, empty string allowed) Matching string in the Body if the email(s) (use with caution for HTML formatted emails)

.PARAMETER ReadStatus
(optional) Status of the email, read or unread

.PARAMETER HasAttachments
(optional) Gets only emails with attachments

.PARAMETER Importance
(optional) Email priority set by sender, high, low, or normal

.PARAMETER Not
(optional) Reverses all other filters, result will exclude all mail that match specified filters. E.g : '-to "john@contoso.com" -subject "alert" -NOT' will exclude all mails sent to John that contain the word "alert" in Subjects

.EXAMPLE
$Mails = Get-ExchangeMail -ExchangeService $exchService -FolderPath "inbox" -Subject "alert" -SentAfter "2021-08-10T23:59:00" -From "network@contoso.com" -Importance "high"

.EXAMPLE
$MailsWithoutAttachments = Get-ExchangeMail -ExchangeService $exchService -FolderPath "inbox\archives" -HasAttachments -Not

.OUTPUTS
Array of Microsoft.Exchange.WebServices.Data.EmailMessage objects, or $null if none found
#>
function Get-ExchangeMail
{

    Param(
        
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [Parameter(Mandatory=$True, ParameterSetName='MailId')]
        [ValidateNotNullOrEmpty()]
        [Alias('mid')]
        [string]$MailId,

        [Parameter(Mandatory=$True, ParameterSetName='FolderId')]
        [ValidateNotNullOrEmpty()]
        [Alias('fid')]
        [string]$FolderId,

        [Parameter(Mandatory=$True, ParameterSetName='FolderPath')]
        [ValidateNotNullOrEmpty()]
        [Alias('fp','path')]
        [string]$FolderPath,

        [Parameter(Mandatory=$True, ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [Alias('o','obj','object')]
        [object]$FolderObject,
        
        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [Alias('n')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][switch]$Not,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $ParameterType = ($_.gettype()).Name
            $ParameterContent = $_
            switch ($ParameterType) {
                "DateTime" { 
                    Write-Verbose "'SentBefore' format is DateTime"
                    Return $True
                }
                "String" {
                    try {
                        Write-Verbose "Check 'SentBefore' date string format : ParameterContent"
                        $CheckFormat = [System.DateTime]::ParseExact($ParameterContent,'yyyy-MM-ddTHH:mm:ss',$null)
                        Return $True
                    }
                    catch [System.Management.Automation.MethodInvocationException] {
                        Throw "Invalid format ParameterContent - excepted format is DateTime or String ('yyyy-MM-ddTHH:mm:ss')"
                    }
                    catch {
                        Throw
                    }
                }
                Default { Throw "$ParameterType not supported - excepted format is DateTime or String ('yyyy-MM-ddTHH:mm:ss')"}
            }
        })]
        [Alias('sb')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')]$SentBefore,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $ParameterType = ($_.gettype()).Name
            $ParameterContent = $_
            switch ($ParameterType) {
                "DateTime" { 
                    Write-Verbose "'SentAfter' format is DateTime"
                    Return $True
                }
                "String" {
                    try {
                        Write-Verbose "Check 'SentAfter' date string format : ParameterContent"
                        $CheckFormat = [System.DateTime]::ParseExact($ParameterContent,'yyyy-MM-ddTHH:mm:ss',$null)
                        Return $True
                    }
                    catch [System.Management.Automation.MethodInvocationException] {
                        Throw "Invalid format ParameterContent - excepted format is DateTime or String ('yyyy-MM-ddTHH:mm:ss')"
                    }
                    catch {
                        Throw
                    }
                }
                Default { Throw "$ParameterType not supported - excepted format is DateTime or String ('yyyy-MM-ddTHH:mm:ss')"}
            }
        })]
        [Alias('sa')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')]$SentAfter,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $ParameterType = ($_.gettype()).Name
            $ParameterContent = $_
            switch ($ParameterType) {
                "DateTime" { 
                    Write-Verbose "'ReceivedBefore' format is DateTime"
                    Return $True
                }
                "String" {
                    try {
                        Write-Verbose "Check 'ReceivedBefore' date string format : ParameterContent"
                        $CheckFormat = [System.DateTime]::ParseExact($ParameterContent,'yyyy-MM-ddTHH:mm:ss',$null)
                        Return $True
                    }
                    catch [System.Management.Automation.MethodInvocationException] {
                        Throw "Invalid format ParameterContent - excepted format is DateTime or String ('yyyy-MM-ddTHH:mm:ss')"
                    }
                    catch {
                        Throw
                    }
                }
                Default { Throw "$ParameterType not supported - excepted format is DateTime or String ('yyyy-MM-ddTHH:mm:ss')"}
            }
        })]
        [Alias('rb')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')]$ReceivedBefore,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $ParameterType = ($_.gettype()).Name
            $ParameterContent = $_
            switch ($ParameterType) {
                "DateTime" { 
                    Write-Verbose "'ReceivedAfter' format is DateTime"
                    Return $True
                }
                "String" {
                    try {
                        Write-Verbose "Check 'ReceivedAfter' date string format : ParameterContent"
                        $CheckFormat = [System.DateTime]::ParseExact($ParameterContent,'yyyy-MM-ddTHH:mm:ss',$null)
                        Return $True
                    }
                    catch [System.Management.Automation.MethodInvocationException] {
                        Throw "Invalid format ParameterContent - excepted format is DateTime or String ('yyyy-MM-ddTHH:mm:ss')"
                    }
                    catch {
                        Throw
                    }
                }
                Default { Throw "$ParameterType not supported - excepted format is DateTime or String ('yyyy-MM-ddTHH:mm:ss')"}
            }
        })]
        [Alias('ra')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')]$ReceivedAfter,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [Alias('f')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$From,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$To,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$Cc,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$Bcc,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [Alias('dt','dto')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$DisplayTo,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateNotNullOrEmpty()]
        [Alias('dcc')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$DisplayCc,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [AllowEmptyString()]
        [Alias('s','title')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$Subject,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [AllowEmptyString()]
        [Alias('b','text')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$Body,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateSet("Read","Unread")]
        [Alias('r','rs','read')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$ReadStatus,

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateSet("Yes","No","All")]
        [Alias('a')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')][string]$HasAttachments="All",

        [Parameter(ParameterSetName='FolderId')]
        [Parameter(ParameterSetName='FolderPath')]
        [Parameter(ParameterSetName='FolderObject')]
        [ValidateSet("High","Low","Normal")]
        [Alias('i', 'p', 'priority')]
        [parameter(Mandatory=$False, ParameterSetName='Filters')]
        [string]$Importance

    )

    try {

        If ((($ExchangeService.Url).ToString()).ToLower() -eq "https://outlook.office365.com/ews/exchange.asmx") {
            Throw "This function is not yet compatible with Office 365 Exchange server"
        }

        if ($MailId) {
            Write-Verbose "Get Mail object for Id $($MailId)"
            $idResult = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($ExchangeService, (New-Object Microsoft.Exchange.WebServices.Data.ItemId($MailId)))
            Return $idResult
        }
        
        if ($FolderId) {
            Write-Verbose "Get destination folder object from Id"
            $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $FolderId
        }elseif ($FolderPath) {
            Write-Verbose "Get folder object from Path"
            $FolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $FolderPath
        }elseif (($FolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied object is not a Folder"
        }

        $itemViewPageSize = 500 # Max = 1000
        $mailItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($itemViewPageSize)
        Write-Verbose "Item View page size set to $($itemViewPageSize)"

        try {

            $SearchFilterOn = $False

            if ($Subject -or $SentBefore -or $SentAfter -or $ReceivedBefore -or $ReceivedAfter -or $DisplayTo -or $DisplayCc -or $From -or $ReadStatus -or $Importance) {

                Write-Verbose "Prepare filters..."
                $SearchFilterOn = $True
                $mailFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)

                if ($Subject) {
                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $Subject)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : Subject contains $(if ($not) {"no"}) '$($Subject)'"
                }
    
                if ($SentBefore) {

                    if (($SentBefore.gettype()).Name -eq "String") {
                        $DateTimeSentEnd=[System.DateTime]::ParseExact($SentBefore,'yyyy-MM-ddTHH:mm:ss',$null)
                    }
                    else {
                        $DateTimeSentEnd = $SentBefore
                    }
                    
                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeSent, $DateTimeSentEnd)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : $(if ($not) {"Not"}) Sent before $($DateTimeSentEnd)"

                }
    
                if ($SentAfter) {

                    if (($SentAfter.gettype()).Name -eq "String") {
                        $DateTimeSentStart=[System.DateTime]::ParseExact($SentAfter,'yyyy-MM-ddTHH:mm:ss',$null)
                    }
                    else {
                        $DateTimeSentStart = $SentAfter
                    }

                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeSent, $DateTimeSentStart)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : $(if ($not) {"Not"}) Sent after $($DateTimeSentStart)"

                }
    
                if ($ReceivedBefore) {

                    if (($ReceivedBefore.gettype()).Name -eq "String") {
                        $DateTimeReceivedEnd=[System.DateTime]::ParseExact($ReceivedBefore,'yyyy-MM-ddTHH:mm:ss',$null)
                    }
                    else {
                        $DateTimeReceivedEnd = $ReceivedBefore
                    }

                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeSent, $DateTimeReceivedEnd)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : $(if ($not) {"Not"}) Received before $($DateTimeReceivedEnd)"

                }
    
                if ($ReceivedAfter) {

                    if (($ReceivedAfter.gettype()).Name -eq "String") {
                        $DateTimeReceivedStart=[System.DateTime]::ParseExact($ReceivedAfter,'yyyy-MM-ddTHH:mm:ss',$null)
                    }
                    else {
                        $DateTimeReceivedStart = $ReceivedAfter
                    }

                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeSent, $DateTimeReceivedStart)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : $(if ($not) {"Not"}) Received after $($DateTimeReceivedStart)"

                }
    
                if ($DisplayTo) {
                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DisplayTo, $DisplayTo)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : 'To' display contains $(if ($not) {"no"}) '$($DisplayTo)'"
                }
    
                if ($DisplayCc) {
                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DisplayCc, $DisplayCc)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : 'Cc' display contains $(if ($not) {"no"}) '$($DisplayCc)'"
                }
    
                if ($From) {
                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender, $From)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : Sender address is $(if ($not) {"not"}) '$($From)'"
                }
    
                if ($ReadStatus) {
                    switch ($ReadStatus.ToLower()) {
                        "read" { $IsRead = $True ; break }
                        "unread" { $IsRead = $False ; break }
                    }
                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $($IsRead))
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : Is $(if ($not) {"not"}) read '$($IsRead)'"
                }
    
                if ($HasAttachments.tolower() -ne "All") {

                    if ($HasAttachments.tolower() -eq "yes") {
                        Write-Verbose "Creating filter : Email has attachments"
                        $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $True)
                        
                    }else {
                        Write-Verbose "Creating filter : Email has no attachments"
                        $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $False)
                    }
                    
                    $mailFilter.add($paramFilter)

                    if ($Not) {
                        $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)
                        Write-Verbose "Parameter '-not' specified : reversed attachment filter"
                    }

                }
    
                if ($Importance) {
                    switch ($Importance.tolower()) {
                        "high" { $Importance = "High" ; break }
                        "low" { $Importance = "Low" ; break }
                        "normal" { $Importance = "Normal" ; break }
                    }
                    $paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Importance,$Importance)
                    if ($Not) {$paramFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($paramFilter)}
                    $mailFilter.add($paramFilter)
                    Write-Verbose "Added filter : Email Importance is $(if ($not) {"not"}) '$($Importance)'"
                }

            }
            
        }
        catch {
            Throw "Error preparing filters : $($_.exception.message)"
        }
        
        # Find Items
        $pageCounter = 1
        Write-Verbose "Find mail items (page $($pageCounter))..."
        if ($SearchFilterOn){$fiResult = $FolderObject.FindItems($mailFilter, $mailItemView)}else {$fiResult = $FolderObject.FindItems($mailItemView)}
        while($fiResult.MoreAvailable) {
            $pageCounter++
            Write-Verbose "Find mail items (page $($pageCounter))..."
            $mailItemView.Offset += $itemViewPageSize # Set offset for next page
            if ($SearchFilterOn){$fiResult += $FolderObject.FindItems($mailFilter, $mailItemView)}else {$fiResult += $FolderObject.FindItems($mailItemView)}
        }

        Write-Verbose "Phase 1 : $(($fiResult | measure-object).count) mails found"

        # Define extended propertySet to get full values for each mail
        Write-Verbose "Get email(s) full data..."
        foreach ($mail in $fiResult) {
            $mailPropertySet = New-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
            $mailObjId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0037,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
            $mailPropertySet.Add($mailObjId)
            $mail.Load($mailPropertySet)
        }

        if ($To) {
            if ($Not) {
                $fiResult = $fiResult | Where-Object {$_.ToRecipients -notmatch $To}
            }else {
                $fiResult = $fiResult | Where-Object {$_.ToRecipients -match $To}
            }
            Write-Verbose "Filtering on : To header contains $(if ($not) {"no "})'$($To)'"
        }
        if ($Cc) {
            if ($Not){
                $fiResult = $fiResult | Where-Object {$_.CcRecipients -notmatch $Cc}
            }else {
                $fiResult = $fiResult | Where-Object {$_.CcRecipients -match $Cc}
            }
            Write-Verbose "Filtering on : Cc header contains $(if ($not) {"no "})'$($Cc)'"
        }
        if ($Bcc) {
            if ($Not){
                $fiResult = $fiResult | Where-Object {$_.BccRecipients -notmatch $Bcc}
            }else {
                $fiResult = $fiResult | Where-Object {$_.BccRecipients -match $Bcc}
            }
            Write-Verbose "Filtering on : Bcc header contains $(if ($not) {"no "})'$($Bcc)'"
        }
        if ($Body) {
            if ($Not){
                $fiResult = $fiResult | Where-Object {$_.Body.Text -notmatch $Body}
            }else {
                $fiResult = $fiResult | Where-Object {$_.Body.Text -match $Body}
            }
            Write-Verbose "Filtering on : Body contains $(if ($not) {"no "})'$($Body)'"
        }
        
    }
    catch {
        Throw
    }

    Write-Verbose "Phase 2 : $(($fiResult | measure-object).count) mails found"
    return $fiResult

}

<#
.SYNOPSIS
Move Exchange Mail to Exchange Folder

.DESCRIPTION
Moves an Exchange Email object to specified folder path or object, using EWS Managed Api. Either Ids, Object, or Path (for destination folder) can be used.

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER MailId
Get Email message by its unique Id. Appart from ExchangeService, any other parameter (folder, filters..) becomes irrelevant with this one.

.PARAMETER MailObject
Exchange Web Services Email object. Could be retrieved with function Get-ExchangeMail

.PARAMETER FolderId
Exchange Id of folder.

.PARAMETER DestinationFolderPath
Full path to destination folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"

.PARAMETER DestinationFolderObject
Exchange.WebServices.Data.Folder type object can by specified instead of DestinationFolderPath

.EXAMPLE
Move-ExchangeMail -ExchangeService $exchService -MailObject $mail -DestinationFolderPath "inbox\archives\folder01"

.EXAMPLE
Move-ExchangeMail -ExchangeService $exchService -MailId "AAMkAGQ5MWNkN2Q3LWE5N..." -DestinationFolderObject $folder

.OUTPUTS
$True if move is successful
#>
function Move-ExchangeMail
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
        
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        ## 'Mail' group (MailObject and MailId are mutually exclusive)

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName="MailObject-DestinationFolderId")]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName="MailObject-DestinationFolderPath")]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName="MailObject-DestinationFolderObject")]
        [ValidateNotNullOrEmpty()]
        [Alias('mo')]
        [object]$MailObject,

        [Parameter(Mandatory, ParameterSetName="MailId-DestinationFolderId")]
        [Parameter(Mandatory, ParameterSetName="MailId-DestinationFolderPath")]
        [Parameter(Mandatory, ParameterSetName="MailId-DestinationFolderObject")]
        [ValidateNotNullOrEmpty()]
        [Alias('mi')]
        [string]$MailId,

        ## 'DestinationFolder' group (DestinationFolderId, DestinationFolderPath and DestinationFolderObject are mutually exclusive)

        [Parameter(Mandatory, ParameterSetName="MailObject-DestinationFolderId")]
        [Parameter(Mandatory, ParameterSetName="MailId-DestinationFolderId")]
        [ValidateNotNullOrEmpty()]
        [Alias('fi','folderid')]
        [string]$DestinationFolderId,

        [Parameter(Mandatory, ParameterSetName="MailObject-DestinationFolderPath")]
        [Parameter(Mandatory, ParameterSetName="MailId-DestinationFolderPath")]
        [ValidateNotNullOrEmpty()]
        [Alias('fp','fpath','folderpath')]
        [string]$DestinationFolderPath,

        [Parameter(Mandatory, ParameterSetName="MailObject-DestinationFolderObject")]
        [Parameter(Mandatory, ParameterSetName="MailId-DestinationFolderObject")]
        [ValidateNotNullOrEmpty()]
        [Alias('fo','fobj','folderobject')]
        [object]$DestinationFolderObject

    )

    try {

        If ($MailId) {
            Write-Verbose "Get mail object from Id"
            $MailObject = Get-ExchangeMail -ExchangeService $ExchangeService -MailId $MailId
        }elseif (($MailObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.EmailMessage") {
            Throw "Supplied object is not an Email"
        }
        
        if ($DestinationFolderId) {
            Write-Verbose "Get destination folder object from Id"
            $DestinationFolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderId $DestinationFolderId
        }elseif ($DestinationFolderPath) {
            Write-Verbose "Get destination folder object from Path"
            $DestinationFolderObject = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath $DestinationFolderPath
        }elseif (($DestinationFolderObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder") {
            Throw "Supplied object is not a Folder"
        }

        if ($PSCmdlet.ShouldProcess("Folder '$($DestinationFolderObject.DisplayName)'", "Move mail $($MailObject.Id.UniqueId)")) {
            Write-Verbose "Moving mail $($MailObject.Id.UniqueId) to $($DestinationFolderObject.DisplayName)"
            $MailObject.Move($DestinationFolderObject.Id) | Out-Null # Out-Null used here not to go into pipeline
        }
        
    }
    catch {
        Throw
    }

    Return $True

}

<#
.SYNOPSIS
Save Exchange mail attachments

.DESCRIPTION
Extracts and saves mail attached files to disk or network share using EWS Managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER MailId
Email message by its unique Id.

.PARAMETER MailObject
Exchange Web Services Email object. Could be retrieved with function Get-ExchangeMail

.PARAMETER DestinationFolder
Files will be saved here ; can be either a String with full path to target directory, or System.IO.DirectoryInfo object

.PARAMETER Like
(optional) Applied as filter on attached file names

.EXAMPLE
Save-ExchangeMailAttachment -ExchangeService $es -MailObject $MailObj -DestinationFolder "c:\download"

.EXAMPLE
Save-ExchangeMailAttachment -ExchangeService $es -MailId "AAMkAGQ5MWNkN2Q3LWE5N..." -DestinationFolder (Get-Item 'd:\temp') -Like "*.txt"

.OUTPUTS
$True when all files successfully saved
#>

function Save-ExchangeMailAttachment {

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
        
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [Parameter(Mandatory=$True, ParameterSetName='MailId')]
        [ValidateNotNullOrEmpty()]
        [Alias('id')]
        [string]$MailId,

        [Parameter(Mandatory=$True, ParameterSetName='MailObject')]
        [ValidateNotNullOrEmpty()]
        [Alias('mo')]
        $MailObject,

        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $ParameterType = ($_.gettype()).Name
            $ParameterContent = $_
            If( ("DirectoryInfo","String") -contains $ParameterType) { 
                Write-Verbose "'DestinationFolder' format is $ParameterType"
                If (Test-Path $ParameterContent -PathType Container) {
                    $True
                }else {
                    Throw "Destination folder not found"
                }
            }
            else {
                Throw "DestinationFolder parameter must be 'String' (full path to target directory) or 'DirectoryInfo' (leaf) type"
            }
        })]
        [Alias('o','output','destination','dir','directory','folder')]
        [parameter(Mandatory=$True)]$DestinationFolder,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('l')]
        [string]$Like

    )

    try {

        If ($MailId) {
            Write-Verbose "Get mail object from Id"
            $MailObject = Get-ExchangeMail -ExchangeService $ExchangeService -MailId $MailId
        }elseif (($MailObject.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.EmailMessage") {
            Throw "Supplied object is not an Email"
        }    
        If (!$MailObject.HasAttachments) {Throw "No attachment found in mail $($MailObject.Id)"}
    
        if (($DestinationFolder.gettype()).Name -eq "FileInfo") {
            $DestinationFolderPath = $DestinationFolder.FullName
        }
        else {
            $DestinationFolderPath = $DestinationFolder
        }
    
        If ([string]::IsNullOrEmpty($Like)) {
            $Attachments = $MailObject.Attachments
        }else {
            $Attachments = $MailObject.Attachments | Where-Object {$_.Name -like $Like}
        }
    
        Write-Verbose "[Begin Loop on mail attachments]"
        $result = $True
        foreach ($attachment in $Attachments) {
                
            $filename = $attachment.Name
    
            if ($PSCmdlet.ShouldProcess("folder $DestinationFolderPath", "Save file $($filename)")) {
    
                try {
                    Write-Verbose "Saving file $($filename) to folder $DestinationFolderPath"
                    $attachment.Load("$($DestinationFolderPath)\$($filename)")
                    Write-Verbose "Saved file $($filename) to folder $DestinationFolderPath"
                }
                catch {
                    $result = $False
                    Write-Warning "Could not save file $($filename) --> $($_.exception.message)"
                }
                
            }
    
        }
        Write-Verbose "[End Loop on mail attachments]"

    }
    catch {
        Throw
    }

    Return $Result

}

<#
.SYNOPSIS
Create a Meeting

.DESCRIPTION
Creates a new Meeting using Exchange Web Services managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER Title
"Subject" of the meeting (empty string allowed)

.PARAMETER Body
Meeting message body (empty string allowed), text-only or HTML

.PARAMETER StartDate
Meeting start date, format yyyy-MM-ddTHH:mm:ss

.PARAMETER EndDate
Meeting end date, format yyyy-MM-ddTHH:mm:ss

.PARAMETER Location
Meeting location (optional, empty string allowed)

.PARAMETER RequiredAttendees
"Outlook-like" semi-colon separated list of email addresses

.PARAMETER OptionalAttendees
"Outlook-like" semi-colon separated list of email addresses

.PARAMETER Attachments
Semi-colon separated list of files full paths (optional)

.EXAMPLE
New-ExchangeMeeting -ExchangeService $es -Title "Presentation" -Body $body -StartDate '2021-08-10T08:30:00' -EndDate '2021-08-10T09:45:00'

.EXAMPLE
New-ExchangeMeeting -ExchangeService $es -Title $title -Body '<p>Hello. This is a test meeting.<br><br>Regards&nbsp;!</p>' -StartDate '2021-08-10T08:30:00' -EndDate '2021-08-10T09:45:00' -Attachments "c:\documents\file1.txt;c:\pictures\image file 2.jpg"

.OUTPUTS
Unique id for created meeting as a String
#>
function New-ExchangeMeeting
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [parameter(Mandatory=$True)]
        [AllowEmptyString()]
        [Alias('t','s','subject')]
        [string]$Title,

        [parameter(Mandatory=$True)]
        [AllowEmptyString()]
        [Alias('b','text')]
        [string]$Body,

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('start')]
        [string]$StartDate,

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('end')]
        [string]$EndDate,

        [parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [Alias('l')]
        [string]$Location,

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('to','req')]
        [string]$RequiredAttendees,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('cc','opt')]
        [string]$OptionalAttendees,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('a','f','files')]
        [string]$Attachments

    )

    Try {

        # setup extended property set
        $CleanGlobalObjectId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting, 0x23, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
        $psPropSet.Add($CleanGlobalObjectId);

        # Bind to the Calendar folder  
        # $folderid generates a true
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)
        $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$folderid)

        # Convert date strings to system.datetime
        $MeetingStartDatetime=[System.DateTime]::ParseExact($StartDate,'yyyy-MM-ddTHH:mm:ss',$null)
        $MeetingEndDatetime=[System.DateTime]::ParseExact($EndDate,'yyyy-MM-ddTHH:mm:ss',$null)

        # Split attendees strings
        if (-not [string]::IsNullOrEmpty($RequiredAttendees)) {$RequiredAttendeesList = $RequiredAttendees.Split(";") }
        if (-not [string]::IsNullOrEmpty($OptionalAttendees)) {$OptionalAttendeesList = $OptionalAttendees.Split(";") }

        # Create Appointment object
        $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $ExchangeService
            $appointment.Subject = $Title
            $appointment.Body = $Body
            $appointment.Start = $MeetingStartDatetime;
            $appointment.End = $MeetingEndDatetime;
            if (-not [string]::IsNullOrEmpty($Location)) {$appointment.Location = $Location;}
            foreach ($attendee in $RequiredAttendeesList) {
                $null = $appointment.RequiredAttendees.Add($attendee)
            }
            foreach ($attendee in $OptionalAttendeesList) {
                $null = $appointment.OptionalAttendees.Add($attendee)
            }

        # Add attachment(s) if specified
        If ($Attachments){
            $AttachmentsList = $Attachments.Split(";")
            ForEach ($file in $AttachmentsList){
                $appointment.Attachments.AddFileAttachment($file) | Out-Null ; # Out-Null used here not to go into pipeline
            }
        }

        #$RequiredAttendees = $row.advisorEmail;
        #if($RequiredAttendees) {$RequiredAttendees | %{[void]$appointment.RequiredAttendees.Add($_)}}

        if ($PSCmdlet.ShouldProcess("New meeting", "Save")) {
            Write-Verbose "Saving new meeting"
            $appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)

            Write-Verbose "Setting the unique id for the appointment and converting to text"
            $appointment.Load($psPropSet);
            $CalIdVal = $null;
            $appointment.TryGetProperty($CleanGlobalObjectId, [ref]$CalIdVal) | Out-Null ; # Out-Null used here not to go into pipeline
            $CalIdVal64 = [Convert]::ToBase64String($CalIdVal)
        }        

    }
    Catch{
        Write-Error "Meeting was not created --> $($_.Exception.Message)" -ErrorAction:Continue
        return $null
    }

    # Return Unique id for created meeting
    return [string]$CalIdVal64

}

<#
.SYNOPSIS
Edit a Meeting

.DESCRIPTION
Edit an existing Meeting using Exchange Web Services managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER Title
New "subject" of the meeting (empty string allowed)

.PARAMETER Body
New Meeting message body (empty string allowed), text-only or HTML

.PARAMETER StartDate
New Meeting start date, format yyyy-MM-ddTHH:mm:ss

.PARAMETER EndDate
New Meeting end date, format yyyy-MM-ddTHH:mm:ss

.PARAMETER Location
New Meeting location (empty string allowed)

.PARAMETER RequiredAttendees
New "Outlook-like" semi-colon separated list of email addresses

.PARAMETER OptionalAttendees
New "Outlook-like" semi-colon separated list of email addresses

.PARAMETER Attachments
New semi-colon separated list of files full paths

.PARAMETER MeetingId
Id of the Meeting to modify

.EXAMPLE
Edit-ExchangeMeeting -ExchangeService $es -Body "Modified meeting body message" -MeetingId "BAAAAIIA4AB0xbcQGoLgCAAAAABLH8CjP87WAQAAAAAAAAAAEAAAADA99jEdyitLqpMM8yghhMU="

.EXAMPLE
Edit-ExchangeMeeting -ExchangeService $es -Title "Postponed Meeting" -Body '<p>Hello. This is a modified test meeting.<br><br>Regards&nbsp;!</p>' -StartDate '2021-08-10T08:30:00' -EndDate '2021-08-10T08:45:00' -MeetingId $MeetingId -OptionalAttendees "georges@domain.net"

.OUTPUTS
String with Id of the created meeting, or Null + Exception message if modification failed
#>
function Edit-ExchangeMeeting
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [Alias('t','s','subject')]
        [string]$Title,

        [parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [Alias('b','text')]
        [string]$Body,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('start')]
        [string]$StartDate,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('end')]
        [string]$EndDate,

        [parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [Alias('l')]
        [string]$Location,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('to','req')]
        [string]$RequiredAttendees,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('cc','opt')]
        [string]$OptionalAttendees,

        [parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [Alias('a','f','files')]
        [string]$Attachments,

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('i','id')]
        [string]$MeetingId

    )

    Try{

        # setup extended property set
        $CleanGlobalObjectId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting, 0x23, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
        $psPropSet.Add($CleanGlobalObjectId);

        # Bind to the Calendar folder  
        # $folderid generates a true
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)     
        $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$folderid)

        # Find Item
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
        $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($CleanGlobalObjectId, $MeetingId);
        $fiResult = $Calendar.FindItems($sfSearchFilter, $ivItemView) 

        # Returns Null if no meeting found with specified Id
        if ($fiResult.TotalCount -eq 0){return $null}

        ForEach ($appointment in $fiResult) { 

            # Update Attachments if specified
            If ($Attachments){
                $AttachmentsList = $Attachments.Split(";")
                # Clear attachments collection
                $appointment.Attachments.Clear()
                # Save updated meeting without sending updates to attendees, to clear old attachments
                $appointment.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve, $True)
                # Add new attachment(s)
                ForEach ($file in $AttachmentsList){
                    $appointment.Attachments.AddFileAttachment($file) | Out-Null ; # Out-Null used here not to go into pipeline
                }
            }

            # Update required attendees if specified
            If (-not [string]::IsNullOrEmpty($RequiredAttendees)){
                # Split attendees strings
                $RequiredAttendeesList = $RequiredAttendees.Split(";")
                # Clear attendees collection
                $appointment.RequiredAttendees.Clear()
                # Add new attendee(s)
                foreach ($attendee in $RequiredAttendeesList) {
                    $null = $appointment.RequiredAttendees.Add($attendee)
                }
            }

            # Update optional attendees if specified
            If (-not [string]::IsNullOrEmpty($OptionalAttendees)){
                # Split attendees strings
                $OptionalAttendeesList = $OptionalAttendees.Split(";")
                # Clear attendees collection
                $appointment.OptionalAttendees.Clear()
                # Add new attendee(s)
                foreach ($attendee in $OptionalAttendeesList) {
                    $null = $appointment.OptionalAttendees.Add($attendee)
                }
            }

            # Update subject if specified
            If ($PSBoundParameters.ContainsKey('Title')){
                $appointment.Subject = $Title
            }

            # Update body if specified
            If ($PSBoundParameters.ContainsKey('Body')){
                $appointment.Body = $Body
            }

            # Update start date if specified
            If (-not [string]::IsNullOrEmpty($StartDate)){
                # Convert date strings to system.datetime
                $MeetingStartDatetime=[System.DateTime]::ParseExact($StartDate,'yyyy-MM-ddTHH:mm:ss',$null)
                # Set meeting start date
                $appointment.Start = $MeetingStartDatetime;
            }

            # Update end date if specified
            If (-not [string]::IsNullOrEmpty($EndDate)){
                # Convert date strings to system.datetime
                $MeetingEndDatetime=[System.DateTime]::ParseExact($EndDate,'yyyy-MM-ddTHH:mm:ss',$null)
                # Set meeting end date
                $appointment.End = $MeetingEndDatetime;
            }

            # Update location if specified
            If ($PSBoundParameters.ContainsKey('Location')){
                $appointment.Location = $Location;
            }

            if ($PSCmdlet.ShouldProcess("All attendees", "Save updated meeting and send notification")) {
                Write-Verbose "Saving updated meeting and sending notification to all attendees"
                $appointment.Update([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
            }

        }

    }

    Catch{
        Write-Error "Error during meeting update --> $($_.Exception.Message)" -ErrorAction:Continue
        return $null
    }

    return $MeetingId

}

<#
.SYNOPSIS
Delete or Cancel a Meeting

.DESCRIPTION
Deletes or Cancel a Meeting using Exchange Web Services managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER Delete
Use "-delete" to completely delete a Meeting

.PARAMETER MeetingId
Id of the Meeting to delete

.EXAMPLE
$MeetingState = Remove-ExchangeMeeting -ExchangeService $es -MeetingId "BAAAAIIA4AB0xbcQGoLgCAAAAAAp1gKGUsnWAQAAAAAAAAAAEAAAAHlrfMPoxtBGv8a7N7md0Zk="

.EXAMPLE
$MeetingState = Remove-ExchangeMeeting -MeetingId $MeetingId -ExchangeService $es -Delete $True

.OUTPUTS
String with Id of the created meeting, or Null + Exception message if creation failed
#>
function Stop-ExchangeMeeting
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,
        
        [parameter(Mandatory=$False)]
        [Alias('d')]
        [switch]$Delete,

        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('i','id')]
        [string]$MeetingId

    )

    Try{

        if ($null -eq $ExchangeService) {
            # Prepare parameters for new Exchange service
            $NewExchangeServiceParams = @{
                WebServiceUrl = $ExchangeWebServiceUrl
                WebServiceDll = $ExchangeWebServiceDll
                UserName = $ExchangeUserName
                SecurePassword = $ExchangePassword
            }
            $ExchangeService = New-ExchangeService @NewExchangeServiceParams
        }

        # setup extended property set
        $CleanGlobalObjectId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting, 0x23, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
        $psPropSet.Add($CleanGlobalObjectId);

        # Bind to the Calendar folder  
        # $folderid generates a true
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)     
        $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$folderid)

        # Find Item
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
        $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($CleanGlobalObjectId, $MeetingId);
        $fiResult = $Calendar.FindItems($sfSearchFilter, $ivItemView) 

        # Returns Null if no meeting found with specified Id
        if ($fiResult.TotalCount -eq 0){return $null}

        foreach ($appointment in $fiResult) { 
            
            # Deletes meeting or cancel it depending of "-Delete" argument
            If ($Delete){
                $appointment.Delete(0);
            }
            else {
                $appointment.CancelMeeting() | Out-Null # Out-Null used here not to go into pipeline
            }
            
        }

    }

    Catch{
        Write-Error "Error during meeting cancelation --> $($_.Exception.Message)" -ErrorAction:Continue
        return $null
    }

    return $MeetingId

}

<#
.SYNOPSIS
Remove Exchange Item

.DESCRIPTION
Removes Exchange Mail or Folder using EWS Managed Api

.PARAMETER ExchangeService
ExchangeService object. Could be retrieved with function New-ExchangeService

.PARAMETER ExchangeItem
Exchange Mail or Folder object to delete

.PARAMETER DeleteMode
Optional. "Default" behaviour is to move to the mailbox's Deleted Items folder. "Soft" will move it to the dumpster (items in the dumpster can be recovered). "Hard" will permanently delete the item. 

.EXAMPLE
Remove-ExchangeItem -ExchangeService $ExchangeService -ExchangeItem $MailObject

.EXAMPLE
Remove-ExchangeItem -ExchangeService $ExchangeService -ExchangeItem (Get-ExchangeFolder -ExchangeService $exchserv -FolderPath "inbox\archives")

.OUTPUTS
$True if item removed successfully

.NOTES
A folder that previously contained an item (folder, mail) that was not hard deleted (which means, it is still present in the deleted items or dumpster) can only be deleted using the "Hard" delete mode.
#>
function Remove-ExchangeItem
{

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
        
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Alias('service','es')]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService,

        [parameter(Mandatory=$True, ValueFromPipeline = $true)]$ExchangeItem,
        [parameter(Mandatory=$False)][ValidateSet("Hard","Soft","Default")][string]$DeleteMode="Default"

    )

    try {

        if ((($ExchangeItem.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.EmailMessage") -and (($ExchangeItem.GetType()).FullName -ne "Microsoft.Exchange.WebServices.Data.Folder")) {
            Throw "Item object must be an email or a folder"
        }

        switch ($DeleteMode.tolower()) {
            "hard" { $DeleteModeDescription = "Hard delete : item will be permanently deleted" ; $DeleteModeInt = 0 ; break }
            "soft" { $DeleteModeDescription = "Soft delete : item will be moved to the dumpster. Items in the dumpster can be recovered." ; $DeleteModeInt = 1 ; break }
            "default" { $DeleteModeDescription = "Item will be moved to the mailbox's Deleted Items folder" ; $DeleteModeInt = 2 ; break }
        }

        if ($PSCmdlet.ShouldProcess("$($ExchangeItem.Id.UniqueId)", "Remove item ($DeleteModeDescription)")) {
            Write-Verbose "Removing item $($ExchangeItem.Id.UniqueId) ($DeleteModeDescription)"
            $ExchangeItem.Delete($DeleteModeInt) | Out-Null # Out-Null used here not to go into pipeline
        }
        
    }
    catch {
        Throw
    }

    Return $True

}
#Endregion