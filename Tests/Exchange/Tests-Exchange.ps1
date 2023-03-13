Write-Host -ForegroundColor Green "Prepare attachments"
$attachment_path_01 = "$TempDir\Send-ExchangeMail-attachment-test_01.txt"
$attachment_path_02 = "$TempDir\Send-ExchangeMail-attachment-test_02.txt"
Write-Output "Send-ExchangeMail function attachment test 01" | Out-File $attachment_path_01
Write-Output "Send-ExchangeMail function attachment test 02" | Out-File $attachment_path_02

try {

    ## Test 1 : New-ExchangeService

    Write-Host -ForegroundColor Green "Test 1A"
    $es = New-ExchangeService -WebServiceUrl $ExchServerUrl -WebServiceDll $ExchDllPath -UserName $username -SecurePassword $SecurePassword
    if ($null -eq $es) {Throw "Test Failed"}

## Test 2 : Get-ExchangeMailFolder

    Write-Host -ForegroundColor Green "Test 2A"
    $TestParentFolderObject = Get-ExchangeMailFolder -es $es -FolderPath "inbox" # Get Folder object by Path
    if ($null -eq $TestParentFolderObject) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 2B"
    $TestParentFolderId = (Get-ExchangeMailFolder -es $es -FolderId $TestParentFolderObject.Id.UniqueId).Id.UniqueId # Get Folder object by Id
    if ($null -eq $TestParentFolderId) {Throw "Test Failed"}

## Test 3 : New-ExchangeMailFolder

    # Create new folder with parent folder Id
    Write-Host -ForegroundColor Green "Test 3A"
    $TestFolder01_id = New-ExchangeMailFolder -es $es -NewFolderDisplayName "Test-Exchange_Folder-01" -ParentFolderId $TestParentFolderId
    if ($null -eq $TestFolder01_id) {Throw "Test Failed"}

    # Create new folder with parent folder Path
    Write-Host -ForegroundColor Green "Test 3B"
    $TestFolder01B_id = New-ExchangeMailFolder -es $es -NewFolderDisplayName "Test-Exchange_Folder-01B" -ParentFolderPath "inbox\Test-Exchange_Folder-01"
    if ($null -eq $TestFolder01_id) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 3C"
    $TestFolder01B = Get-ExchangeMailFolder -es $es -FolderId $TestFolder01B_id
    if ($null -eq $TestFolder01B) {Throw "Test Failed"}

    # Create new folder with parent folder Object
    Write-Host -ForegroundColor Green "Test 3D"
    $TestFolder02_id = New-ExchangeMailFolder -es $es -NewFolderDisplayName "Test-Exchange_Folder-02" -ParentFolderObject $TestParentFolderObject
    if ($null -eq $TestFolder02_id) {Throw "Test Failed"}

    ## Test 4 : Move-ExchangeMailFolder

    Write-Host -ForegroundColor Green "Test 4A"
    $TestMoveFolder01 = Move-ExchangeMailFolder -es $es -SourceFolderId $TestFolder01B_id -DestinationFolderId $TestFolder02_id # id-id
    if ($null -eq $TestMoveFolder01) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 4B"
    $TestMoveFolder02 = Move-ExchangeMailFolder -es $es -SourceFolderId $TestFolder01B_id -DestinationFolderPath "inbox\Test-Exchange_Folder-01" # id-path
    if ($null -eq $TestMoveFolder02) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 4C"
    $TestMoveFolder03 = Move-ExchangeMailFolder -es $es -SourceFolderId $TestFolder01B_id -DestinationFolderObject $TestParentFolderObject # id-object
    if ($null -eq $TestMoveFolder03) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 4D"
    $TestMoveFolder04 = Move-ExchangeMailFolder -es $es -SourceFolderPath "inbox\Test-Exchange_Folder-01B" -DestinationFolderId $TestFolder02_id # path-id
    if ($null -eq $TestMoveFolder04) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 4E"
    $TestMoveFolder05 = $TestFolder01B | Move-ExchangeMailFolder -es $es -DestinationFolderPath "inbox\Test-Exchange_Folder-01" # pipeline, object-path
    if ($null -eq $TestMoveFolder05) {Throw "Test Failed"}
    Write-Host "Sleeping..." ; Start-Sleep 10

    ## Test 5 : Rename-ExchangeMailFolder

    Write-Host -ForegroundColor Green "Test 5A"
    $TestRename01 = Rename-ExchangeMailFolder -es $es -FolderNewName "Test-Renamed-Exchange_Folder-01" -FolderId $TestFolder01_id # Id
    if ($null -eq $TestRename01) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 5B"
    $TestRename02 = Rename-ExchangeMailFolder -es $es -FolderNewName "Test-Renamed-Exchange_Folder-02" -FolderPath "inbox\Test-Exchange_Folder-02" # Path
    if ($null -eq $TestRename02) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 5C"
    $TestRename01B = $TestFolder01B | Rename-ExchangeMailFolder -es $es -FolderNewName "Test-Renamed-Exchange_Folder-01B"  # Object, Pipeline
    if ($null -eq $TestRename01B) {Throw "Test Failed"}

    ## Test 6 : Send-ExchangeMail

    # Prepare recipients, must be username, do not change values
    $ToRecipient = $CcRecipient = $BccRecipient = $username

    # Send test emails
    Write-Host -ForegroundColor Green "Test 6A"
    $TestSendmailTo = Send-ExchangeMail -es $es -To $ToRecipient -Title "Send-ExchangeMail function test" -Body "Test To"
    if ($null -eq $TestSendmailTo) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 6B"
    $TestSendmailCc = Send-ExchangeMail -es $es -Cc $CcRecipient -Title "Send-ExchangeMail function test" -Body "Test Cc only" -BodyType 'text'
    if ($null -eq $TestSendmailCc) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 6C"
    Write-Host "Sleeping..." ; Start-Sleep 3
    $TestUnformattedDateAfter= Get-Date
    Write-Host "Sleeping..." ; Start-Sleep 3
    $TestSendmailBcc = Send-ExchangeMail -es $es -To $ToRecipient -Cc $CcRecipient -Bcc $BccRecipient -Title "Send-ExchangeMail function test" -Body "Test To + Cc + Bcc"
    if ($null -eq $TestSendmailBcc) {Throw "Test Failed"}
    Write-Host "Sleeping..." ; Start-Sleep 3
    $TestUnformattedDateBefore = Get-Date
    Write-Host "Sleeping..." ; Start-Sleep 3

    Write-Host -ForegroundColor Green "Test 6D"
    $TestSendmailWithAttachments = Send-ExchangeMail -es $es -To $ToRecipient -Cc $CcRecipient -Title "Send-ExchangeMail function test" -Body "Test To + Cc + Attachments" -Attachments "$attachment_path_01;$attachment_path_02"
    if ($null -eq $TestSendmailWithAttachments) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 6E"
    $TestSendmailImportance = Send-ExchangeMail -es $es -To $ToRecipient -Title "Send-ExchangeMail Importance Test" -Body "Test Importance" -Importance "high"
    if ($null -eq $TestSendmailTo) {Throw "Test Failed"}

    ## Test 7 : Get-ExchangeMail

    Write-Host "Sleeping..." ; Start-Sleep 10

    Write-Host -ForegroundColor Green "Test 7A"
    $TestMails01 = Get-ExchangeMail -es $es -FolderId $TestParentFolderObject.Id.UniqueId -Subject "Send-ExchangeMail func" -To $ToRecipient -Body "To +"
    If ($TestMails01.count -ne 2) {Throw "Mails count should be 2 !"}

    Write-Host -ForegroundColor Green "Test 7b"
    $TestMail02 = Get-ExchangeMail -es $es -FolderPath "inbox" -Subject "Send-ExchangeMail func" -Cc $CcRecipient -Body "Test Cc only" -HasAttachments "No"
    If ($TestMail02.count -ne 1) {Throw "Mail count should be 1 !"}

    Write-Host -ForegroundColor Green "Test 7c"
    $TestDateBefore = Get-Date $TestUnformattedDateBefore -Format "yyyy-MM-ddTHH:mm:ss"
    $TestDateAfter = Get-Date $TestUnformattedDateAfter -Format "yyyy-MM-ddTHH:mm:ss"
    $TestMail03 = Get-ExchangeMail -es $es -FolderObject $TestParentFolderObject -SentBefore $TestDateBefore -SentAfter $TestDateAfter
    If ($TestMail03.count -ne 1) {Throw "Mail count should be 1 !"}

    ## Test 8 : Send-ExchangeMailReply

    Write-Host -ForegroundColor Green "Test 8A"
    $TestReply01 = Send-ExchangeMailReply -es $es -ReplyToAll -ReplyString "Reply Test 01 : simple reply to all" -MailId $TestMail02.id.UniqueId
    if ($null -eq $TestReply01) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 8B"
    $TestReply02 = $TestMail03 | Send-ExchangeMailReply -es $es -ReplyString "Reply Test 02 : reply to another mail, but reply only to sender and add To+Cc+Bcc recipients" -AddTo $ToRecipient2 -AddCc $CcRecipient2 -AddBcc $BccRecipient2
    if ($null -eq $TestReply02) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 8C"
    $TestReply03 = $TestMail03 | Send-ExchangeMailReply -es $es -ReplyToAll -ReplyString "Reply Test 03 : reply to same mail, without subject, with attachment and additional To+Cc recipients" -AddTo $ToRecipient2 -AddCc $CcRecipient2 -Attachments "$attachment_path_01;$attachment_path_02"
    if ($null -eq $TestReply03) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 8D"
    $TestReply04 = $TestMail03 | Send-ExchangeMailReply -es $es -ReplyString "Reply Test 04 : reply to same mail, with importance" -AddTo $ToRecipient2 -Importance "low"
    if ($null -eq $TestReply04) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 8E"
    $TestReply05 = Send-ExchangeMailReply -es $es -ReplyString "" -Forward -AddTo $ToRecipient2 -MailId $TestMail02.id.UniqueId
    if ($null -eq $TestReply05) {Throw "Test Failed"}

    ## Test 9 : Move-ExchangeMail

    Write-Host -ForegroundColor Green "Test 9A"
    foreach ($mail in $TestMails01) {
        $TestMoveMail01 = $mail | Move-ExchangeMail -es $es -DestinationFolderObject $TestFolder01B
        if ($null -eq $TestMoveMail01) {Throw "Test Failed"}
    }

    Write-Host -ForegroundColor Green "Test 9B"
    $TestMoveMail02 = $TestMail02 | Move-ExchangeMail -es $es -DestinationFolderId $TestFolder02_id
    if ($null -eq $TestMoveMail02) {Throw "Test Failed"}

    ## Test 10 : Clear-ExchangeMailFolder

    Write-Host -ForegroundColor Green "Test 10A"
    $TestClearFolder01 = Clear-ExchangeMailFolder -es $es -FolderPath "inbox\Test-Renamed-Exchange_Folder-02"
    if ($null -eq $TestClearFolder01) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 10B"
    $TestClearFolder02 = Clear-ExchangeMailFolder -es $es -FolderPath "inbox\Test-Renamed-Exchange_Folder-01" -Recurse
    if ($null -eq $TestClearFolder02) {Throw "Test Failed"}

    ## Test 11 : Remove-ExchangeMailFolder

    Write-Host -ForegroundColor Green "Test 11A"
    $TestRemoveFolder01 = Remove-ExchangeMailFolder -es $es -FolderId $TestFolder02_id -DeleteMode hard
    if ($null -eq $TestRemoveFolder01) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 11B"
    $TestRemoveFolder02 = Remove-ExchangeMailFolder -es $es -FolderPath "inbox\Test-Renamed-Exchange_Folder-01" -Recurse -DeleteMode Hard
    if ($null -eq $TestRemoveFolder02) {Throw "Test Failed"}

    ## Test 12 : New-ExchangeMeeting

    Write-Host -ForegroundColor Green "Test 12A"
    $TestNewMeeting01 = New-ExchangeMeeting -es $es -Title "New-ExchangeMeeting function test A" -Body "This is a meeting" -Location "somewhere" -StartDate $MeetingA_Start1 -EndDate $MeetingA_End1 -RequiredAttendees "$ToRecipient;$ToRecipient2" -OptionalAttendees $BccRecipient2  -Attachments "$attachment_path_01;$attachment_path_02"
    if ($null -eq $TestNewMeeting01) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 12B"
    $TestNewMeeting02 = New-ExchangeMeeting -es $es -Title "New-ExchangeMeeting function test B" -Body "" -StartDate $MeetingB_Start1 -EndDate $MeetingB_End1 -Location "" -RequiredAttendees $ToRecipient2
    if ($null -eq $TestNewMeeting02) {Throw "Test Failed"}

    ## Test 13 : Edit-ExchangeMeeting

    Write-Host "Sleeping..." ; Start-Sleep 5

    Write-Host -ForegroundColor Green "Test 13A"
    $TestEditMeeting01 = Edit-ExchangeMeeting -es $es -MeetingId $TestNewMeeting01 -Title "Edit-ExchangeMeeting function test A" -Body "<b><u>New meeting body</u></b>" -StartDate $MeetingA_Start2 -EndDate $MeetingA_End2 -Location "" -RequiredAttendees "$ToRecipient2;$BccRecipient2" -OptionalAttendees $ToRecipient
    if ($null -eq $TestEditMeeting01) {Throw "Test Failed"}

    ## Test 14 : Stop-ExchangeMeeting

    Write-Host "Sleeping..." ; Start-Sleep 5

    Write-Host -ForegroundColor Green "Test 14A"
    $TestStopMeeting01 = Stop-ExchangeMeeting -es $es -MeetingId $TestNewMeeting01
    if ($null -eq $TestStopMeeting01) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 14B"
    $TestStopMeeting02 = Stop-ExchangeMeeting -es $es -MeetingId $TestNewMeeting02 -Delete
    if ($null -eq $TestStopMeeting02) {Throw "Test Failed"}

    ## Test 15 : Save-ExchangeMailAttachment

    Write-Host -ForegroundColor Green "Test 15A"
    $TestMailWithAttachments01 = Get-ExchangeMail -es $es -FolderPath "inbox" -Subject "Send-ExchangeMail function test" -HasAttachments Yes
    $TestSaveAttachments01 = Save-ExchangeMailAttachment -es $es -MailObject $TestMailWithAttachments01 -DestinationFolder $TempDir -Like "*.txt"
    if ($TestSaveAttachments01 -ne $True) {Throw "Test Failed"}

    Write-Host -ForegroundColor Green "Test 15B"
    $TestMailWithAttachments02 = Get-ExchangeMail -es $es -FolderPath "inbox" -Subject "Send-ExchangeMail function test" -HasAttachments Yes
    $TestSaveAttachments02 = Save-ExchangeMailAttachment -es $es -MailId $TestMailWithAttachments01.Id.UniqueId -DestinationFolder (Get-Item $TempDir) -Like "*"
    if ($TestSaveAttachments02 -ne $True) {Throw "Test Failed"}

}
catch {
    Throw
}


## Remove temporary attachment files
Remove-Item -path $attachment_path_01
Remove-Item -path $attachment_path_02