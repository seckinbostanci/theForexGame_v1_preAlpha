Attribute VB_Name = "LogModule"
Public Sub Log(strLog As String)
'Simply add an event to the log file, has no relevance to the authentication itself
'but simply displays information to the user
With FrmMain.LogBox
    .SelColor = vbBlue
    .SelText = "[" & Time & "] "
    .SelColor = vbWhite
    .SelText = strLog & vbCrLf
End With
WriteToLogFile (strLog)
End Sub
Public Sub LogRAW(strLog As String)
'Simply add an event to the log file, has no relevance to the authentication itself
'but simply displays information to the user

'Logs raw data received by winsock in red :)

With FrmMain.LogBox
    .SelColor = vbRed
    .SelText = "[" & Time & "] "
    .SelColor = vbWhite
    .SelText = "Socket Data: " & strLog & vbCrLf
End With
WriteToLogFile (strLog)
End Sub

Public Sub WriteToLogFile(logentry As String)
    Dim loging As Integer
    loging = FreeFile
    Open App.Path & "\" & Date & "_" & App.EXEName & ".log" For Append As loging
    Write #loging, Now & ": " & logentry
    Close #loging
End Sub
