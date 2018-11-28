'-----------------------------------------------------------------------------------------
' UDF function to return the file extension name
' Input takes either full file path or just the file name
'-----------------------------------------------------------------------------------------
Public Function Get_File_Extension_Name(ByVal Full_File_Path As String) As String
Dim fs As Object

On Error GoTo Invalid_File_Path
Set fs = CreateObject("Scripting.FileSystemObject")
Get_File_Extension_Name = fs.GetExtensionName(Full_File_Path)

Exit Function

Invalid_File_Path:
Get_File_Extension_Name = ""
End Function
