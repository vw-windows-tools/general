# FUNCTIONS
## Function New-ExchangeService
### DESCRIPTION
*Create new ExchangeService object using Exchange Web Service API*
### SYNTAX
```powershell
New-ExchangeService [-WebServiceUrl <String>] [-WebServiceDll <String>] -UserName <String> -ClearPassword <String> [<CommonParameters>]

New-ExchangeService [-WebServiceUrl <String>] [-WebServiceDll <String>] -UserName <String> -SecurePassword <SecureString> [<CommonParameters>]
```
### PARAMETERS
WEBSERVICEURL<br>
Full Url to Exchange web service. If not specified, Office 365 web service URL is used
<br><br>WEBSERVICEDLL<br>
Full path to Microsoft.Exchange.WebServices.dll. If not specified, Function will try to find it using a default list
<br><br>USERNAME<br>
Exchange user name (e.g "john@contoso.com")
<br><br>CLEARPASSWORD<br>
Exchange user password, as plain text (not recommended)
<br><br>SECUREPASSWORD<br>
Exchange user password, as securestring (recommended)
### EXAMPLE
```powershell
$exchServ = New-ExchangeService -WebServiceUrl https://owa.contoso.com/EWS/Exchange.asmx -WebServiceDll $EWS_DllPath -UserName "john@contoso.com" -SecurePassword $SecPass
```
### OUTPUTS
Microsoft.Exchange.WebServices.Data.ExchangeServiceBase type object
<br><br>
## Function New-ExchangeMailFolder
### DESCRIPTION
*Creates new Exchange mail folder using EWS Managed Api*
### SYNTAX
```powershell
New-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -NewFolderDisplayName <String> -ParentFolderPath <String> [-WhatIf] [-Confirm] [<CommonParameters>]

New-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -NewFolderDisplayName <String> -ParentFolderObject <Object> [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>NEWFOLDERDISPLAYNAME<br>
Name of the folder to create
<br><br>PARENTFOLDERPATH<br>
Full path to parent folder into which to create the new folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"
<br><br>PARENTFOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object can by specified instead of ParentFolderPath
### EXAMPLE
```powershell
$newFolderId = New-ExchangeMailFolder -ExchangeService $exchService -NewFolderDisplayName "folder01" -ParentFolderPath "inbox\tests"
```
### OUTPUTS
Unique Id of the successfully created folder
<br><br>
## Function Get-ExchangeMailFolder
### DESCRIPTION
*Returns an object of the Exchange.WebServices.Data.Folder type, for specified path, using Exchange Web Service API*
### SYNTAX
```powershell
Get-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderId <String> [<CommonParameters>]
Get-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderPath <String> [-ParentFolder <Object>] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>FOLDERID<br>
Exchange Id of folder
<br><br>FOLDERPATH<br>
Full path to folder. Could be supplied instead of folder Exchange Id. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"
<br><br>PARENTFOLDER<br>
Optional Exchange.WebServices.Data.Folder type object. If supplied, FolderPath will be read starting from this folder instead of Root folder
### EXAMPLE
```powershell
$archives = Get-ExchangeMailFolder -ExchangeService $ExchangeService -FolderPath "inbox\Archives"
```
### OUTPUTS
Exchange.WebServices.Data.Folder type object
<br><br>
## Function Get-ExchangeMailSubFolders
### SYNTAX
```powershell
Get-ExchangeMailSubFolders -ExchangeService <ExchangeServiceBase> -FolderId <String> [<CommonParameters>]
Get-ExchangeMailSubFolders -ExchangeService <ExchangeServiceBase> -FolderPath <String> [<CommonParameters>]
Get-ExchangeMailSubFolders -ExchangeService <ExchangeServiceBase> -FolderObject <Object> [<CommonParameters>]
```
### DESCRIPTION
*Gets all sub-folders for specified Exchange Mail folder using EWS Managed Api*
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>FOLDERID<br>
Exchange Id of folder. Incompatible with FolderPath and FolderObject parameters
<br><br>FOLDERPATH<br>
Full path to source folder to move. Incompatible with FolderId and FolderObject parameters. Separate folders with Antislashes ("\"). E.g : "Inbox\Tests"
<br><br>FOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object. Incompatible with FolderId and FolderPath parameters.
### EXAMPLES
```powershell
$subfolders = Get-ExchangeMailSubFolders -ExchangeService $exchserv -FolderId "AAMkAGQ5MWNkN2Q3LWE5N..."
$subfolders = Get-ExchangeMailSubFolders -ExchangeService $exchserv -FolderPath "deleted items"
$subfolders = Get-ExchangeMailSubFolders -ExchangeService $exchserv -FolderObject (Get-ExchangeMailFolder -ES $TestExchService -FolderPath inbox)
```
### OUTPUTS
Array of Microsoft.Exchange.WebServices.Data.Folder objects, or $null if no subfolder found
<BR><BR>
## Function Move-ExchangeMailFolder
### DESCRIPTION
*Moves source exchange folder into (under) destination folder.*
### SYNTAX
```powershell
Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderId <String> -DestinationFolderObject <Object> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderId <String> -DestinationFolderPath <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderId <String> -DestinationFolderId <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderPath <String> -DestinationFolderObject <Object> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderPath <String> -DestinationFolderPath <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderPath <String> -DestinationFolderId <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderObject <Object> -DestinationFolderObject <Object> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderObject <Object> -DestinationFolderPath <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -SourceFolderObject <Object> -DestinationFolderId <String> [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>SOURCEFOLDERID<br>
Exchange Id of folder. Incompatible with SourceFolderPath and SourceFolderObject parameter
<br><br>SOURCEFOLDERPATH<br>
Full path to source folder to move. Incompatible with SourceFolderId and SourceFolderObject parameters. Separate folders with Antislashes ("\"). E.g : "Inbox\Tests"
<br><br>SOURCEFOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object. Incompatible with SourceFolderId and SourceFolderPath parameters.
<br><br>DESTINATIONFOLDERID<br>
Exchange Id of folder. Incompatible with DestinationFolderPath and DestinationFolderObject parameters
<br><br>DESTINATIONFOLDERPATH<br>
Full path to Destination folder to move. Incompatible with DestinationFolderId and DestinationFolderObject parameters. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"
<br><br>DESTINATIONFOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object. Incompatible with DestinationFolderId and DestinationFolderPath parameters.
### EXAMPLES
```powershell
Move-ExchangeMailFolder -SourceFolderPath "Inbox\Tests" -DestinationFolderPath "Inbox\Archives"-ExchangeService $exchService

Move-ExchangeMailFolder -SourceFolderObject $SourceFolder -DestinationFolderId "AAMkAGQ5MWNkN2Q3LWE5N..." -ExchangeService $exchService
```
### OUTPUTS
$True if folder is successfully moved to new location
<BR><BR>
## Function Rename-ExchangeMailFolder
### DESCRIPTION
*Renames Exchange mail folder using EWS Managed Api*
### SYNTAX
```powershell
Rename-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderNewName <String> -FolderId <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Rename-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderNewName <String> -FolderPath <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Rename-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderNewName <String> -FolderObject <Object> [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>FOLDERNEWNAME<br>
New name of the folder to update
<br><br>FOLDERID<br>
Exchange Id of folder. Incompatible with FolderPath and FolderObject parameters
<br><br>FOLDERPATH<br>
Full path to folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests". Incompatible with FolderId and FolderId parameters
<br><br>FOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object. Incompatible with FolderPath and FolderId parameters
### EXAMPLE
$Result = Rename-ExchangeMailFolder -ExchangeService $exchService -FolderPath $FolderPath -FolderNewName $NewName
### OUTPUTS
$True if folder is successfully renamed
<BR><BR>
## Function Remove-ExchangeMailFolder
### DESCRIPTION
*Deletes Exchange folder using Exchange Web Services API. Full path (e.g "inbox\archives") or EWS Folder object can be specified.*
### SYNTAX
```powershell
Remove-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderId <String> [-DeleteMode <String>] [-Recurse] [-WhatIf] [-Confirm] [<CommonParameters>]

Remove-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderPath <String> [-DeleteMode <String>] [-Recurse] [-WhatIf] [-Confirm] [<CommonParameters>]

Remove-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderObject <Object> [-DeleteMode <String>] [-Recurse] [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>FOLDERID<br>
Exchange Id of folder. Incompatible with FolderPath and FolderObject parameter
<br><br>FOLDERPATH<br>
Full path to folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests". Incompatible with FolderId and FolderId parameters
<br><br>FOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object. Incompatible with FolderPath and FolderId parameters
### EXAMPLES
```powershell
Remove-ExchangeMailFolder -ExchangeService $exs -FolderId "AAMkAGQ5MWNkN2Q3LWE5N..."

Remove-ExchangeMailFolder -ExchangeService $exs -FolderPath "inbox\archives\john" -Hard

Remove-ExchangeMailFolder -ExchangeService $exs -FolderObject $myFolder -Soft
```
### OUTPUTS
$True if folder(s) removed successfully
### NOTES
The standard and -Hard options are transactional, which means that by the time a web service call completes, the database has moved the item to the Deleted Items folder or permanently removed the item from the Exchange database.
The -Soft delete option works differently for different target versions of Exchange Server. Soft Delete for Exchange 2007 sets a bit on the item that indicates to the Exchange database that the item will be moved to the dumpster folder at an indeterminate time in the future.
Soft Delete for versions of Exchange starting with Exchange 2010, including Exchange Online, immediately moves the item to the dumpster. Soft Delete is not an option for folder deletion. Soft Delete traversal searches for items and folders will not return any results.
On a side note, a folder that previously contained an item (folder,mail) that was not hard deleted (which means, it is still present in the deleted items or dumpster) can only be deleted using the "Hard" delete mode.
<BR><BR>
## Function Clear-ExchangeMailFolder
### DESCRIPTION
*Clears a folder using Exchange Web Services Managed Api. Erases all emails within specified folder. Applies to subfolders with -recurse parameter*
### SYNTAX
```powershell
Clear-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderId <String> [-DeleteMode <String>] [-Recurse] [-WhatIf] [-Confirm] [<CommonParameters>]

Clear-ExchangeMailFolder -ExchangeService <ExchangeServiceBase> -FolderPath <String> [-DeleteMode <String>] [-Recurse] [-WhatIf] [-Confirm] [<CommonParameters>]

Clear-ExchangeMailFolder-ExchangeService <ExchangeServiceBase> -FolderObject <Object> [-DeleteMode <String>] [-Recurse] [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>FOLDERID<br>
Exchange Id of folder. Incompatible with FolderPath and FolderObject parameters
<br><br>FOLDERPATH<br>
Full path to folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests". Incompatible with FolderId and FolderId parameters
<br><br>FOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object. Incompatible with FolderPath and FolderId parameters
<br><br>DELETEMODE<br>
Optional. "Default" behaviour is to move all erased emails to the mailbox's Deleted Items folder. "Soft" will move them to the dumpster (items in the dumpster can be recovered). "Hard" will permanently delete the emails. 
<br><br>RECURSE<br>
Optional. Deletes all mails in specified folder + all mails in subfolders
### EXAMPLES
```powershell
Clear-ExchangeMailFolder -ExchangeService $exchService -FolderPath "inbox\archives\old"

Clear-ExchangeMailFolder -ExchangeService $exchService -FolderPath "inbox\archives" -recurse

Clear-ExchangeMailFolder -ExchangeService $exchService -FolderObject (Get-ExchangeMailFolder -FolderPath "inbox\archives\old")
```
### OUTPUTS
$True if folder(s) cleared successfully
<BR><BR>
## Function Send-ExchangeMail
### DESCRIPTION
*Sends an email using Exchange Web Service Managed Api*
### SYNTAX
```powershell
Send-ExchangeMail -ExchangeService <ExchangeServiceBase> -To <String> -Cc <String> -Bcc <String> -Title <String> [-Body <String>] [-BodyType <String>] [-Attachments <String>] [-Importance <String>] [-WhatIf] [-Confirm] [<CommonParameters>]

Send-ExchangeMail -ExchangeService <ExchangeServiceBase> -To <String> -Bcc <String> -Title <String> [-Body <String>] [-BodyType <String>] [-Attachments <String>] [-Importance <String>] [-WhatIf] [-Confirm] [<CommonParameters>]

Send-ExchangeMail -ExchangeService <ExchangeServiceBase> -To <String> -Cc <String> -Title <String> [-Body <String>] [-BodyType <String>] [-Attachments <String>] [-Importance <String>] [-WhatIf] [-Confirm] [<CommonParameters>]

Send-ExchangeMail -ExchangeService <ExchangeServiceBase> -To <String> -Title <String> [-Body <String>] [-BodyType <String>] [-Attachments <String>] [-Importance <String>] [-WhatIf] [-Confirm] [<CommonParameters>]

Send-ExchangeMail -ExchangeService <ExchangeServiceBase> -Cc <String> -Bcc <String> -Title <String> [-Body <String>] [-BodyType <String>] [-Attachments <String>] [-Importance <String>] [-WhatIf] [-Confirm] [<CommonParameters>]

Send-ExchangeMail -ExchangeService <ExchangeServiceBase> -Cc <String> -Title <String> [-Body <String>] [-BodyType <String>] [-Attachments <String>] [-Importance <String>] [-WhatIf] [-Confirm] [<CommonParameters>]

Send-ExchangeMail -ExchangeService <ExchangeServiceBase> -Bcc <String> -Title <String> [-Body <String>] [-BodyType <String>] [-Attachments <String>] [-Importance <String>] [-WhatIf] [-Confirm] [<CommonParameters>]PARAMETERS
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>TO<br>
"Outlook-like" semi-colon separated list of email addresses (optional if Cc or Bcc is specified)
<br><br>CC<br>
"Outlook-like" semi-colon separated list of email addresses (optional if To or Bcc is specified)
<br><br>BCC<br>
"Outlook-like" semi-colon separated list of email addresses (optional if To or Cc is specified)
<br><br>TITLE<br>
"Subject" of the email (mandatory, empty string accepted)
<br><br>BODY<br>
Mail body (optional, empty string accepted)
<br><br>BODYTYPE<br>
"Text" or "HTML". Set body type (optional, Default is HTML).
<br><br>ATTACHMENTS<br>
semi-colon separated list of files full paths
<br><br>IMPORTANCE<br>
(optional) Email priority set by sender, high, low, or normal
### EXAMPLES
```powershell
$Return = Send-ExchangeMail -ExchangeService $exchserv -To "tom@contoso.com;john@domain.com" -Title 'Hi' -Body $mailbody -BodyType "text"
Send-ExchangeMail -ExchangeService $exchserv -Attachments "c:\docs\file1.txt;c:\pictures\file2.jpg" -Bcc "tom@contoso.com" -Title "Hello" -Body "hello<br><br>world"
```
### OUTPUTS
$True if mail successfully sent
### NOTES
To, Cc and Bcc are all optional parameters, but one must be specified at least.
<BR><BR>
## Function Send-ExchangeMailReply
### DESCRIPTION
*Sends an email reply using Exchange Web Services Managed Api*
### SYNTAX
```powershell
Send-ExchangeMailReply -ExchangeService <ExchangeServiceBase> -ReplyString <String> -MailId <String> [-AddTo <String>] [-AddCc <String>] [-AddBcc <String>] [-Importance <String>] [-Attachments <String>] [-ReplyToAll] [-WhatIf] [-Confirm] [<CommonParameters>]

Send-ExchangeMailReply -ExchangeService <ExchangeServiceBase> -ReplyString <String> -MailObject <Object> [-AddTo <String>] [-AddCc <String>] [-AddBcc <String>] [-Importance <String>] [-Attachments <String>] [-ReplyToAll] [-WhatIf] [-Confirm] [<CommonParameters>]PARAMETERS
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>MAILID<br>
Get Email message by its unique Id. Incompatible with MailObject parameter.
<br><br>MAILOBJECT<br>
Exchange Web Services Email object. Incompatible with MailId parameter. Could be retrieved with Function Get-ExchangeMail
<br><br>REPLYSTRING<br>
Reply message to be sent ; can be empty, plain text or Html. History of previous email(s) will be kept.
<br><br>ADDTO<br>
Semi-colon separated list of recipients to add in To recipients
<br><br>ADDCC<br>
Semi-colon separated list of recipients to add in Cc recipients
<br><br>ADDBCC<br>
Semi-colon separated list of recipients to add in Bcc recipients
<br><br>ATTACHMENTS<br>
Semi-colon separated list of full path to file(s) to be added as attachment(s)
<br><br>REPLYTOALL<br>
If specified, all recipients of the initial email will be kept as new recipients. If omitted, only the sender will be kept. Can be associated in both cases with To, Cc, and Bcc parameters.
<br><br>IMPORTANCE<br>
(optional) Email priority set by sender, high, low, or normal
<br><br>FORWARD<br>
(optional) Create 'Forward' response mail instead if Reply mail. At least one recipient must be specified with this option, using 'AddTo' parameter
### EXAMPLES
```powershell
Send-ExchangeMailReply -ExchangeService $exchService -MailObject $mail -ReplyString "Hello<br><br>World" -Attachments "C:\users\john\doc\prices.pdf;c:\users\john\doc\john.vcf"

Send-ExchangeMailReply -ExchangeService $exchService -MailId "AAMkAGQ5MWNkN2Q3LWE5N..." -ReplyString "Thanks everyone for your cooperation" -AddCc "ceo@contoso.com" -ReplyAll

Send-ExchangeMailReply -es $es -Forward -ReplyString "" -AddTo 'supervision@contoso.com' -MailId "AAMkAGQ5MWNkN2Q3LWE5N..."
```
### OUTPUTS
$True if mail was successfully sent.
<BR><BR>
## Function Get-ExchangeMail
### DESCRIPTION
*Gets Exchange Mail(s) in specified Folder using EWS Managed Api, with several filters on recipients, dates, etc*
### SYNTAX
```powershell
Get-ExchangeMail -ExchangeService <ExchangeServiceBase> -MailId <String> [<CommonParameters>]

Get-ExchangeMail -ExchangeService <ExchangeServiceBase> -FolderId <Object> [-Not] [-SentBefore <String or DateTime>] [-SentAfter <String or DateTime>] [-ReceivedBefore <String or DateTime>] [-ReceivedAfter <String or DateTime>] [-From <String>] [-To <String>] [-Cc <String>] [-Bcc <String>] [-DisplayTo <String>] [-DisplayCc <String>] [-Subject <String>] [-Body <String>] [-ReadStatus <String>] [-HasAttachments <String>] [-Importance <String>] [<CommonParameters>]

Get-ExchangeMail -ExchangeService <ExchangeServiceBase> -FolderPath <Object> [-Not] [-SentBefore <String or DateTime>] [-SentAfter <String or DateTime>] [-ReceivedBefore <String or DateTime>] [-ReceivedAfter <String or DateTime>] [-From <String>] [-To <String>] [-Cc <String>] [-Bcc <String>] [-DisplayTo <String>] [-DisplayCc <String>] [-Subject <String>] [-Body <String>] [-ReadStatus <String>] [-HasAttachments <String>] [-Importance <String>] [<CommonParameters>]

Get-ExchangeMail -ExchangeService <ExchangeServiceBase> -FolderObject <Object> [-Not] [-SentBefore <String or DateTime>] [-SentAfter <String or DateTime>] [-ReceivedBefore <String or DateTime>] [-ReceivedAfter <String or DateTime>] [-From <String>] [-To <String>] [-Cc <String>] [-Bcc <String>] [-DisplayTo <String>] [-DisplayCc <String>] [-Subject <String>] [-Body <String>] [-ReadStatus <String>] [-HasAttachments <String>] [-Importance <String>] [<CommonParameters>]

Get-ExchangeMail -ExchangeService <ExchangeServiceBase> [-Not] [-SentBefore <String or DateTime>] [-SentAfter <String or DateTime>] [-ReceivedBefore <String>] [-ReceivedAfter <String or DateTime>] [-From <String>] [-To <String>] [-Cc <String>] [-Bcc <String>] [-DisplayTo <String>] [-DisplayCc <String>] [-Subject <String>] [-Body <String>] [-ReadStatus <String>] [-HasAttachments <String>] [-Importance <String>] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>MAILID<br>
Get Email message by its unique Id. Appart from ExchangeService, any other parameter (folder, filters..) becomes irrelevant with this one.
<br><br>FOLDERPATH<br>
Full path to Exchange mail folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"
<br><br>FOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object can by specified instead of FolderPath
<br><br>SENTBEFORE<br>
(optional) Minimum send date (as datetime, or as string, format "yyyy-MM-ddTHH:mm:ss") of the email(s)
<br><br>SENTAFTER<br>
(optional) Maximum send date (as datetime, or as string, format "yyyy-MM-ddTHH:mm:ss") of the email(s)
<br><br>RECEIVEDBEFORE<br>
(optional) Minimum receive date (as datetime, or as string, format "yyyy-MM-ddTHH:mm:ss") of the email(s)
<br><br>RECEIVEDAFTER<br>
(optional) Maximum send date (as datetime, or as string, format "yyyy-MM-ddTHH:mm:ss") of the email(s)
<br><br>FROM<br>
(optional) Exact email address of the Sender
<br><br>TO<br>
(optional) Matching string in "To" recipients list (can be an full/partial email address or name)
<br><br>CC<br>
(optional) Matching string in "Cc" recipients list (can be an full/partial email address or name)
<br><br>BCC<br>
(optional) Matching string in "Bcc" recipients list (can be an full/partial email address or name)
<br><br>DISPLAYTO<br>
(optional) Matching string in "To" recipients *names* list (not addresses)
<br><br>DISPLAYCC<br>
(optional) Matching string in "Cc" recipients *names* list (not addresses)
<br><br>SUBJECT<br>
(optional, empty string allowed) Matching string in Subject if the email(s)
<br><br>BODY<br>
(optional, empty string allowed) Matching string in the Body if the email(s) (use with caution for HTML formatted emails)
<br><br>READSTATUS<br>
(optional) Status of the email, read or unread
<br><br>HASATTACHMENTS<br>
(optional) Gets only emails with attachments
<br><br>IMPORTANCE<br>
(optional) Email priority set by sender, high, low, or normal
<br><br>NOT<br>
(optional) Reverses all other filters, result will exclude all mail that match specified filters. E.g : '-to "john@contoso.com" -subject "alert" -NOT' will exclude all mails sent to John that contain the word "alert" in Subjects
### EXAMPLES
```powershell
$Mails = Get-ExchangeMail -ExchangeService $exchService -FolderPath "inbox" -Subject "alert" -SentAfter "2021-08-10T23:59:00" -From "network@contoso.com" -Importance "high"

$MailsWithoutAttachments = Get-ExchangeMail -ExchangeService $exchService -FolderPath "inbox\archives" -HasAttachments -Not
```
### OUTPUTS
Array of Microsoft.Exchange.WebServices.Data.EmailMessage objects, or $null if none found
<BR><BR>
## Function Move-ExchangeMail
### DESCRIPTION
*Moves an Exchange Email object to specified folder path or object, using EWS Managed Api. Either Ids, Object, or Path (for destination folder) can be used.*
### SYNTAX
```powershell
Move-ExchangeMail -ExchangeService <ExchangeServiceBase> -MailObject <Object> -DestinationFolderObject <Object> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMail -ExchangeService <ExchangeServiceBase> -MailObject <Object> -DestinationFolderPath <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMail -ExchangeService <ExchangeServiceBase> -MailObject <Object> -DestinationFolderId <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMail -ExchangeService <ExchangeServiceBase> -MailId <String> -DestinationFolderObject <Object> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMail -ExchangeService <ExchangeServiceBase> -MailId <String> -DestinationFolderPath <String> [-WhatIf] [-Confirm] [<CommonParameters>]

Move-ExchangeMail -ExchangeService <ExchangeServiceBase> -MailId <String> -DestinationFolderId <String> [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>MAILID<br>
Get Email message by its unique Id. Appart from ExchangeService, any other parameter (folder, filters..) becomes irrelevant with this one.
<br><br>MAILOBJECT<br>
Exchange Web Services Email object. Could be retrieved with Function Get-ExchangeMail
<br><br>FOLDERID<br>
Exchange Id of folder.
<br><br>DESTINATIONFOLDERPATH<br>
Full path to destination folder. Separate folders with Antislashes ("\"). E.g : "Inbox\Archives\Tests"
<br><br>DESTINATIONFOLDEROBJECT<br>
Exchange.WebServices.Data.Folder type object can by specified instead of DestinationFolderPath
### EXAMPLES
```powershell
Move-ExchangeMail -ExchangeService $exchService -MailObject $mail -DestinationFolderPath "inbox\archives\folder01"

Move-ExchangeMail -ExchangeService $exchService -MailId "AAMkAGQ5MWNkN2Q3LWE5N..." -DestinationFolderObject $folder
```
### OUTPUTS
$True if move is successful
<BR><BR>
## Function Save-ExchangeMailAttachment
### DESCRIPTION
*Extracts and saves mail attached files to disk or network share using EWS Managed Api*
### SYNTAX
```powershell
Save-ExchangeMailAttachment -ExchangeService <ExchangeServiceBase> -MailId <String> -DestinationFolder <Object> [-Like <String>] [-WhatIf] [-Confirm] [<CommonParameters>]
    
Save-ExchangeMailAttachment -ExchangeService <ExchangeServiceBase> -MailObject <Object> -DestinationFolder <Object> [-Like <String>] [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>MAILID<br>
Email message by its unique Id.
<br><br>MAILOBJECT<br>
Exchange Web Services Email object. Could be retrieved with function Get-ExchangeMail
<br><br>DESTINATIONFOLDER<br>
Files will be saved here ; can be either a String with full path to target directory, or System.IO.DirectoryInfo object
<br><br>LIKE<br>
(optional) Applied as filter on attached files names
### EXAMPLES
```powershell
Save-ExchangeMailAttachment -ExchangeService $es -MailObject $MailObj -DestinationFolder "c:\download"

Save-ExchangeMailAttachment -ExchangeService $es -MailId "AAMkAGQ5MWNkN2Q3LWE5N..." -DestinationFolder (Get-Item 'd:\temp') -Like "*.txt"
```
### OUTPUTS
$True when all files successfully saved
<BR><BR>
## Function New-ExchangeMeeting
### DESCRIPTION
*Creates a new Meeting using Exchange Web Services managed Api*
### SYNTAX
```powershell
New-ExchangeMeeting [-ExchangeService] <ExchangeServiceBase> [-Title] <String> [-Body] <String> [-StartDate] <String> [-EndDate] <String> [[-Location] <String>] [-RequiredAttendees] <String> [[-OptionalAttendees] <String>] [[-Attachments] <String>] [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>TITLE<br>
"Subject" of the meeting (empty string allowed)
<br><br>BODY<br>
Meeting message body (empty string allowed), text-only or HTML
<br><br>STARTDATE<br>
Meeting start date, format yyyy-MM-ddTHH:mm:ss
<br><br>ENDDATE<br>
Meeting end date, format yyyy-MM-ddTHH:mm:ss
<br><br>LOCATION<br>
Meeting location (optional, empty string allowed)
<br><br>REQUIREDATTENDEES<br>
"Outlook-like" semi-colon separated list of email addresses
<br><br>OPTIONALATTENDEES<br>
"Outlook-like" semi-colon separated list of email addresses
<br><br>ATTACHMENTS<br>
Semi-colon separated list of files full paths (optional)
### EXAMPLES
```powershell
New-ExchangeMeeting -ExchangeService $es -Title "Presentation" -Body $body -StartDate '2021-08-10T08:30:00' -EndDate '2021-08-10T09:45:00'

New-ExchangeMeeting -ExchangeService $es -Title $title -Body '<p>Hello. This is a test meeting.<br><br>Regards&nbsp;!</p>' -StartDate '2021-08-10T08:30:00' -EndDate '2021-08-10T09:45:00' -Attachments "c:\documents\file1.txt;c:\pictures\image file 2.jpg"
```
### OUTPUTS
Unique id for created meeting as a String
<BR><BR>
## Function Edit-ExchangeMeeting
### DESCRIPTION
*Edit an existing Meeting using Exchange Web Services managed Api*
### SYNTAX
```powershell
Edit-ExchangeMeeting [-ExchangeService] <ExchangeServiceBase> [[-Title] <String>] [[-Body] <String>] [[-StartDate] <String>][[-EndDate] <String>] [[-Location] <String>] [[-RequiredAttendees] <String>] [[-OptionalAttendees] <String>] [[-Attachments] <String>]  [-MeetingId] <String> [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>TITLE<br>
New "subject" of the meeting (empty string allowed)
<br><br>BODY<br>
New Meeting message body (empty string allowed), text-only or HTML
<br><br>STARTDATE<br>
New Meeting start date, format yyyy-MM-ddTHH:mm:ss
<br><br>ENDDATE<br>
New Meeting end date, format yyyy-MM-ddTHH:mm:ss
<br><br>LOCATION<br>
New Meeting location (empty string allowed)
<br><br>REQUIREDATTENDEES<br>
New "Outlook-like" semi-colon separated list of email addresses
<br><br>OPTIONALATTENDEES<br>
New "Outlook-like" semi-colon separated list of email addresses
<br><br>ATTACHMENTS<br>
New semi-colon separated list of files full paths
<br><br>MEETINGID<br>
Id of the Meeting to modify
### EXAMPLES
```powershell
Edit-ExchangeMeeting -ExchangeService $es -Body "Modified meeting body message" -MeetingId "BAAAAIIA4AB0xbcQGoLgCAAAAABLH8CjP87WAQAAAAAAAAAAEAAAADA99jEdyitLqpMM8yghhMU="

Edit-ExchangeMeeting -ExchangeService $es -Title "Postponed Meeting" -Body '<p>Hello. This is a modified test meeting.<br><br>Regards&nbsp;!</p>' -StartDate '2021-08-10T08:30:00' -EndDate '2021-08-10T08:45:00' -MeetingId $MeetingId -OptionalAttendees "georges@domain.net"
```
### OUTPUTS
String with Id of the created meeting, or Null + Exception message if modification failed
<BR><BR>
## Function Stop-ExchangeMeeting
### DESCRIPTION
*Deletes or Cancel a Meeting using Exchange Web Services managed Api*
### SYNTAX
```powershell
Stop-ExchangeMeeting [-ExchangeService] <ExchangeServiceBase> [-Delete] [-MeetingId] <String> [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>DELETE<br>
Use "-delete" to completely delete a Meeting
<br><br>MEETINGID<br>
Id of the Meeting to delete
### EXAMPLES
```powershell
$MeetingState = Remove-ExchangeMeeting -ExchangeService $es -MeetingId "BAAAAIIA4AB0xbcQGoLgCAAAAAAp1gKGUsnWAQAAAAAAAAAAEAAAAHlrfMPoxtBGv8a7N7md0Zk="

$MeetingState = Remove-ExchangeMeeting -MeetingId $MeetingId -ExchangeService $es -Delete $True
```
### OUTPUTS
String with Id of the created meeting, or Null + Exception message if creation failed
<BR><BR>
## Function Remove-ExchangeItem
### DESCRIPTION
*Removes Exchange Mail or Folder using EWS Managed Api*
### SYNTAX
```powershell
Remove-ExchangeItem [-ExchangeService] <ExchangeServiceBase> [-ExchangeItem] <Object> [[-DeleteMode] <String>] [-WhatIf] [-Confirm] [<CommonParameters>]
```
### PARAMETERS
EXCHANGESERVICE<br>
ExchangeService object. Could be retrieved with Function New-ExchangeService
<br><br>EXCHANGEITEM<br>
Exchange Mail or Folder object to delete
<br><br>DELETEMODE<br>
Optional. "Default" behaviour is to move to the mailbox's Deleted Items folder. "Soft" will move it to the dumpster (items in the dumpster can be recovered). "Hard" will permanently delete the item. 
### EXAMPLES
```powershell
Remove-ExchangeItem -ExchangeService $ExchangeService -ExchangeItem $MailObject

Remove-ExchangeItem -ExchangeService $ExchangeService -ExchangeItem (Get-ExchangeFolder -ExchangeService $exchserv -FolderPath "inbox\archives")
```
### OUTPUTS
$True if item removed successfully
### NOTES
A folder that previously contained an item (folder, mail) that was not hard deleted (which means, it is still present in the deleted items or dumpster) can only be deleted using the "Hard" delete mode.

