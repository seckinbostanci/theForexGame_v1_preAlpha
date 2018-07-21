Attribute VB_Name = "TDecryptModule"
'Function created by Jeffrey C. Talum
'Avaliable at: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=8527&lngWId=1
Function TDecrypt(iString)
    On Error GoTo uhohs
    Q = ""
    zz = Left(iString, 3)
    a = Left(zz, 1)
    b = Mid(zz, 2, 1)
    c = Mid(zz, 3, 1)
    d = Right(iString, 1)
    a = Int(Asc(a)) 'key 1
    b = Int(Asc(b)) 'key 2
    c = Int(Asc(c)) 'key 3
    d = Int(Asc(d)) 'key 4
    txt = Left(iString, Len(iString) - 1)
    txt2 = Mid(txt, 4, Len(txt)) 'encrypted text
    e = 1


    For X = 1 To Len(txt2)
        f = Mid(txt2, X, 1)
        If e = 1 Then Q = Q & Chr(Asc(f) - a)
        If e = 2 Then Q = Q & Chr(Asc(f) - b)
        If e = 3 Then Q = Q & Chr(Asc(f) - c)
        If e = 4 Then Q = Q & Chr(Asc(f) - d)
        e = e + 1
        If e > 4 Then e = 1
    Next X
    TDecrypt = Q
    Exit Function
uhohs:
    TDecrypt = "Error: Invalid text To Decrypt"
    Exit Function
End Function
