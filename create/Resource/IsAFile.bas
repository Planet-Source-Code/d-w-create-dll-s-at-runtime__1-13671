Attribute VB_Name = "IsAFile"
Option Explicit
Public Function IsFile(FileString As String) As Boolean
Dim FileNumber As Integer 'The Dir function may be
On Error Resume Next 'in use so use this.
FileNumber = FreeFile()
Open FileString For Input As #FileNumber
If Err Then
IsFile = False
Exit Function
End If
IsFile = True
Close #FileNumber
End Function
