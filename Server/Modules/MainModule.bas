Attribute VB_Name = "MainModule"
Public Const DATA_DELIMITER = vbCrLf & "====" & vbCrLf

Public Const MAX_USERS = 99 '0 - 99 = 100 Max users allowed to connect at one time
Public User(0 To MAX_USERS) As typUser 'Declare our user type

Type typUser 'Used for the 'User' variable
    FreeSocket As Boolean
    EncryptionString As String
    HasAuthenticated As Boolean
End Type


Public nIndex As Integer
