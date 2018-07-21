Attribute VB_Name = "HelperModule"
Public Function GenerateAuthString(intIndex As Integer) As String
    Dim strRandomString As String
    'Generates an authentication string, returns the auth string
    strRandomString = GetRandomString(100) 'Generate a 100 character random string
    User(intIndex).EncryptionString = strRandomString 'Set our version to plain text
    GenerateAuthString = TEncrypt(strRandomString) 'Encrypt the clients version
End Function

Public Function GetRandomString(intLength As Integer) As String
    Dim intCharpos As Integer
    Dim intStrLen As Integer
    Dim strRandString As String
    Dim strChars As String
    strChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    intStrLen = 0
    Randomize Timer

    Do Until intStrLen = intLength
        intCharpos = Int((Len(strChars) * Rnd) + 1)
        strRandString = strRandString & Mid(strChars, intCharpos, 1)
        intStrLen = intStrLen + 1
    Loop
    GetRandomString = strRandString
End Function

Function GetRandomNumber(intSeed As Integer)
    Randomize
    GetRandomNumber = Int((Val(intSeed) * Rnd) + 1)
End Function

Public Function CheckAuthentication(strAuthString As String, intIndex As Integer) As Boolean
    'Returns TRUE if this user has sent a valid auth string, FALSE if not.
    If User(intIndex).EncryptionString = strAuthString Then 'If what the user has sent back
    'is what we sent (decrypted version) then this user is authentic!
    CheckAuthentication = True
    Else
    CheckAuthentication = False 'If they do not match, this user is not authentic
    End If
End Function
