Attribute VB_Name = "SocketModule"
Public Sub SendData(strData As String, intIndex As Integer)
    FrmMain.UserSocket(intIndex).SendData strData & DATA_DELIMITER
    DoEvents
End Sub

Public Sub DisconnectUser(intIndex As Integer)
    FrmMain.UserSocket(intIndex).Close
    User(intIndex).FreeSocket = True
    User(intIndex).HasAuthenticated = False
    User(intIndex).EncryptionString = ""
    Log "Socket ID: " & intIndex & " olan kullanýcýnýn baðlantýsý kopartýldý."
End Sub
