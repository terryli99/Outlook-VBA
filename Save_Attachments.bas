'----------------------------------------------------------------------------------
' Author: Jie Jenn
' Comment: Selecting the email first, then run the Save Attachments macro
'----------------------------------------------------------------------------------
Sub SaveAttachments()
Dim olSelection As Selection
Dim olMail As Object
Dim olAttachments As Attachments
Dim FileCount As Long, i As Long
Dim SaveFolderPath As String

On Error GoTo errHandle
'// Where do you want to save the files?
SaveFolderPath = "<Directory where you want to save your attachments>"
Set olSelection = ActiveExplorer.Selection

'----------------------------------------------------
' Extract Attachment(s)
'----------------------------------------------------
'// Here we will iterate each email from our selection
For Each olMail In olSelection
    
    '// making sure it is an actual Outlook mail only
    If TypeName(olMail) = "MailItem" Then
    
        Set olAttachments = olMail.Attachments
        
        FileCount = olAttachments.Count
        
        If FileCount > 0 Then
            
            For i = FileCount To 1 Step -1
                
                '// Save file attachments
                olAttachments.item(i).SaveAsFile SaveFolderPath & olAttachments.item(i).FileName

            Next i
        
        End If
        
        Set olAttachments = Nothing
    
    End If
Next olMail
Exit Sub

errHandle:
MsgBox "Error: " & Err.Description, vbExclamation
End Sub
