Attribute VB_Name = "basSystem"
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function LoggedInUser() As String
On Error GoTo ErrTrap
'To get logged on user
Dim strUserName As String
    'Create a buffer
    strUserName = String(100, Chr$(0))
    'Get the username
    GetUserName strUserName, 100
    'strip the rest of the buffer
    strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
    LoggedInUser = strUserName
    Exit Function
ErrTrap:
    Exit Function
End Function



