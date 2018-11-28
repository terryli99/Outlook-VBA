Sub SaveAttachment()
'--------------------------------------------------------------------------------
' Author: Jie Jenn
' Instruction: Select the email (one email only) you want to extract the attachments first, then
'              run the SaveAttachment macro
'--------------------------------------------------------------------------------
Dim olSelection As Selection
Dim olMailItem As MailItem
Dim olAttachments As Attachments
Dim olAttachment As Attachment
Dim FolderPath As String

If ActiveExplorer.Selection.Count <> 1 Then Exit Sub

FolderPath = "<Your directory path>"

Set olSelection = ActiveExplorer.Selection
Set olMailItem = olSelection.item(1)


For Each olAttachment In olSelection.item(1).Attachments
    olAttachment.SaveAsFile FolderPath & olAttachment.FileName
Next olAttachment

Set olMailItem = Nothing
Set olSelection = Nothing
End Sub
