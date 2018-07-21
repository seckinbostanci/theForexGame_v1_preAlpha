Attribute VB_Name = "TEncryptModule"
'Function created by Jeffrey C. Talum
'Avaliable at: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=8527&lngWId=1
Function TEncrypt(iString)
    On Error GoTo uhoh
    Q = ""
    a = GetRandomNumber(9) + 32
    b = GetRandomNumber(9) + 32
    c = GetRandomNumber(9) + 32
    d = GetRandomNumber(9) + 32
    Q = Chr(a) & Chr(c) & Chr(b)
    e = 1


    For X = 1 To Len(iString)
        f = Mid(iString, X, 1)
        If e = 1 Then Q = Q & Chr(Asc(f) + a)
        If e = 2 Then Q = Q & Chr(Asc(f) + c)
        If e = 3 Then Q = Q & Chr(Asc(f) + b)
        If e = 4 Then Q = Q & Chr(Asc(f) + d)
        e = e + 1
        If e > 4 Then e = 1
    Next X
    Q = Q & Chr(d)
    TEncrypt = Q
    Exit Function
uhoh:
    TEncrypt = "Error: Invalid text To Encrypt"
    Exit Function
End Function
