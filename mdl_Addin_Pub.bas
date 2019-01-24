Attribute VB_Name = "mdl_Addin_Pub"

Function IsFileExists(ByVal strFileName As String) As Boolean
    If Dir(strFileName, 16) <> Empty Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function
'Function IsFileExists(ByVal strFileName As String) As Boolean
'    Dim objFileSystem As Object
'    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
'    If objFileSystem.FileExists(strFileName) = True Then
'        IsFileExists = True
'    Else
'        IsFileExists = False
'    End If
'End Function
