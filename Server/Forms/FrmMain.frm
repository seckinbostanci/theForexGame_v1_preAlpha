VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D8D562C3-878C-11D2-943F-444553540000}#1.0#0"; "ctlist.ocx"
Object = "{B545BF63-340D-11CF-8377-F5ABEBDFD918}#1.0#0"; "ctpush.ocx"
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ForexGame Server"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":0000
   ScaleHeight     =   10740
   ScaleWidth      =   19200
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock ServerSocket 
      Left            =   18120
      Top             =   20
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3467
   End
   Begin MSWinsockLib.Winsock UserSocket 
      Index           =   0
      Left            =   18600
      Top             =   20
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox LogBox 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   9000
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   8555394
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"FrmMain.frx":1A6ED
   End
   Begin CTLISTLibCtl.ctList lstUser 
      Height          =   1935
      Left            =   14950
      TabIndex        =   1
      Top             =   1680
      Width           =   3915
      _Version        =   65536
      _ExtentX        =   6906
      _ExtentY        =   3413
      _StockProps     =   77
      ForeColor       =   16777215
      BackColor       =   8555394
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TipsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleBackImage  =   "FrmMain.frx":1A771
      HeaderPicture   =   "FrmMain.frx":1A78D
      Picture         =   "FrmMain.frx":1A7A9
      TitleText       =   "User List"
      BorderType      =   1
      ListAlign       =   2
      FocusType       =   1
      HeaderOffset    =   4
      PreColumnWidth  =   32
      PicXOffset      =   -2
      ShowTitle       =   -1  'True
      ShowHeader      =   -1  'True
      MaskBitmap      =   -1  'True
      ScrollOnVThumb  =   0   'False
      HeaderData      =   "FrmMain.frx":1A7C5
      PicArray0       =   "FrmMain.frx":1A866
      PicArray1       =   "FrmMain.frx":1A882
      PicArray2       =   "FrmMain.frx":1A89E
      PicArray3       =   "FrmMain.frx":1A8BA
      PicArray4       =   "FrmMain.frx":1A8D6
      PicArray5       =   "FrmMain.frx":1A8F2
      PicArray6       =   "FrmMain.frx":1A90E
      PicArray7       =   "FrmMain.frx":1A92A
      PicArray8       =   "FrmMain.frx":1A946
      PicArray9       =   "FrmMain.frx":1A962
      PicArray10      =   "FrmMain.frx":1A97E
      PicArray11      =   "FrmMain.frx":1A99A
      PicArray12      =   "FrmMain.frx":1A9B6
      PicArray13      =   "FrmMain.frx":1A9D2
      PicArray14      =   "FrmMain.frx":1A9EE
      PicArray15      =   "FrmMain.frx":1AA0A
      PicArray16      =   "FrmMain.frx":1AA26
      PicArray17      =   "FrmMain.frx":1AA42
      PicArray18      =   "FrmMain.frx":1AA5E
      PicArray19      =   "FrmMain.frx":1AA7A
      PicArray20      =   "FrmMain.frx":1AA96
      PicArray21      =   "FrmMain.frx":1AAB2
      PicArray22      =   "FrmMain.frx":1AACE
      PicArray23      =   "FrmMain.frx":1AAEA
      PicArray24      =   "FrmMain.frx":1AB06
      PicArray25      =   "FrmMain.frx":1AB22
      PicArray26      =   "FrmMain.frx":1AB3E
      PicArray27      =   "FrmMain.frx":1AB5A
      PicArray28      =   "FrmMain.frx":1AB76
      PicArray29      =   "FrmMain.frx":1AB92
      PicArray30      =   "FrmMain.frx":1ABAE
      PicArray31      =   "FrmMain.frx":1ABCA
      PicArray32      =   "FrmMain.frx":1ABE6
      PicArray33      =   "FrmMain.frx":1AC02
      PicArray34      =   "FrmMain.frx":1AC1E
      PicArray35      =   "FrmMain.frx":1AC3A
      PicArray36      =   "FrmMain.frx":1AC56
      PicArray37      =   "FrmMain.frx":1AC72
      PicArray38      =   "FrmMain.frx":1AC8E
      PicArray39      =   "FrmMain.frx":1ACAA
      PicArray40      =   "FrmMain.frx":1ACC6
      PicArray41      =   "FrmMain.frx":1ACE2
      PicArray42      =   "FrmMain.frx":1ACFE
      PicArray43      =   "FrmMain.frx":1AD1A
      PicArray44      =   "FrmMain.frx":1AD36
      PicArray45      =   "FrmMain.frx":1AD52
      PicArray46      =   "FrmMain.frx":1AD6E
      PicArray47      =   "FrmMain.frx":1AD8A
      PicArray48      =   "FrmMain.frx":1ADA6
      PicArray49      =   "FrmMain.frx":1ADC2
      PicArray50      =   "FrmMain.frx":1ADDE
      PicArray51      =   "FrmMain.frx":1ADFA
      PicArray52      =   "FrmMain.frx":1AE16
      PicArray53      =   "FrmMain.frx":1AE32
      PicArray54      =   "FrmMain.frx":1AE4E
      PicArray55      =   "FrmMain.frx":1AE6A
      PicArray56      =   "FrmMain.frx":1AE86
      PicArray57      =   "FrmMain.frx":1AEA2
      PicArray58      =   "FrmMain.frx":1AEBE
      PicArray59      =   "FrmMain.frx":1AEDA
      PicArray60      =   "FrmMain.frx":1AEF6
      PicArray61      =   "FrmMain.frx":1AF12
      PicArray62      =   "FrmMain.frx":1AF2E
      PicArray63      =   "FrmMain.frx":1AF4A
      PicArray64      =   "FrmMain.frx":1AF66
      PicArray65      =   "FrmMain.frx":1AF82
      PicArray66      =   "FrmMain.frx":1AF9E
      PicArray67      =   "FrmMain.frx":1AFBA
      PicArray68      =   "FrmMain.frx":1AFD6
      PicArray69      =   "FrmMain.frx":1AFF2
      PicArray70      =   "FrmMain.frx":1B00E
      PicArray71      =   "FrmMain.frx":1B02A
      PicArray72      =   "FrmMain.frx":1B046
      PicArray73      =   "FrmMain.frx":1B062
      PicArray74      =   "FrmMain.frx":1B07E
      PicArray75      =   "FrmMain.frx":1B09A
      PicArray76      =   "FrmMain.frx":1B0B6
      PicArray77      =   "FrmMain.frx":1B0D2
      PicArray78      =   "FrmMain.frx":1B0EE
      PicArray79      =   "FrmMain.frx":1B10A
      PicArray80      =   "FrmMain.frx":1B126
      PicArray81      =   "FrmMain.frx":1B142
      PicArray82      =   "FrmMain.frx":1B15E
      PicArray83      =   "FrmMain.frx":1B17A
      PicArray84      =   "FrmMain.frx":1B196
      PicArray85      =   "FrmMain.frx":1B1B2
      PicArray86      =   "FrmMain.frx":1B1CE
      PicArray87      =   "FrmMain.frx":1B1EA
      PicArray88      =   "FrmMain.frx":1B206
      PicArray89      =   "FrmMain.frx":1B222
      PicArray90      =   "FrmMain.frx":1B23E
      PicArray91      =   "FrmMain.frx":1B25A
      PicArray92      =   "FrmMain.frx":1B276
      PicArray93      =   "FrmMain.frx":1B292
      PicArray94      =   "FrmMain.frx":1B2AE
      PicArray95      =   "FrmMain.frx":1B2CA
      PicArray96      =   "FrmMain.frx":1B2E6
      PicArray97      =   "FrmMain.frx":1B302
      PicArray98      =   "FrmMain.frx":1B31E
      PicArray99      =   "FrmMain.frx":1B33A
   End
   Begin PushLibCtl.ctPush btnServer 
      Height          =   780
      Left            =   14880
      TabIndex        =   2
      Top             =   690
      Width           =   4035
      _Version        =   65536
      _ExtentX        =   7117
      _ExtentY        =   1376
      _StockProps     =   70
      Caption         =   "Running.."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmMain.frx":1B356
      PictureDisabled =   "FrmMain.frx":1B6AB
      PictureDown     =   "FrmMain.frx":1B6C7
      BackColor       =   8555394
      ForeColor       =   16777215
      PicPosition     =   3
      ButtonHeight    =   100
      ButtonWidth     =   37
      PicBevel        =   0
      Toggle          =   -1  'True
      Caption         =   "Running.."
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnServer_Click()
    If btnServer.State = False Then    ' E�er sunucuyu durdurup ba�latacak d��menin �zerinde "&Ba�lat" yaz�yorsa;
        StartServer                             ' Sunucuyu ba�latacak olan fonksiyonumuzu �a��r�yoruz.
        btnServer.Caption = "Running..."        ' D��memizin �zerindeki yaz�y� "&Durdur" olarak de�i�tiriyoruz.
    Else                                        ' E�er sunucuyu durdurup ba�latacak d��menin �zerinde "&Ba�lat" yazm�yorsa;
        StopServer                              ' Sunucuyu durduracak olan fonksiyonumuzu �a��r�yoruz.
        btnServer.Caption = "Stoped."        ' D��memizin �zerindeki yaz�y� "&Ba�lat" olarak de�i�tiriyoruz.
    End If
End Sub

Private Sub Form_Load()
    For i = 1 To MAX_USERS          ' MainModule'de tan�l� olan en y�ksek kullan�c� say�s� kadar olu�turulmu� olan
        User(i).FreeSocket = True   ' T�m kullan�c� de�i�kenlerimizin soketlerini TRUE diyerek bo� olarak ayarl�yoruz.
        Load UserSocket(i)          ' ve her kullan�c�n�n user soketini yarat�yoruz.
    Next i
    User(0).FreeSocket = True       ' ilk kullan�c�m�z�n soketini TRYE diyerek bo� olarak ayarl�yoruz.
    StartServer                     ' Sunucuyu ba�latacak olan fonksiyonumuzu �a��r�yoruz.
    btnServer.State = False
End Sub

Private Sub Form_Terminate()
    Log "Uygulama kapat�lmaya haz�rlan�yor..."
    StopServer  ' Sunucuyu kapatacak olan fonksiyonumuzu �a��r�yoruz.
    Log "Uygulama kapat�ld�."
    Unload Me   ' Form'umuzu ram bellekten ��kar�yoruz.
    End         ' Uygulamam�z� ya�am d�ng�s�n� bitiriyoruz.
End Sub

Private Sub LogBox_Change()
    LogBox.SelStart = Len(LogBox)
End Sub

Private Sub lstUser_ItemDblClick(ByVal nIndex As Long, ByVal nColumn As Integer)
    DlgUserInfo.lblIp.Caption = lstUser.ListColumnText(nIndex, 3)
    'w_Message.Label4.Caption = ctList1.ListColumnText(nIndex, 6)
    'w_Message.Text1.Text = ctList1.ListSubText(nIndex)
    DlgUserInfo.Show 1
End Sub

Private Sub ServerSocket_ConnectionRequest(ByVal requestID As Long)
    Dim i As Integer
    Dim strAuthString As String
    
    For i = 0 To MAX_USERS
        'Yeni gelen ba�lant� iste�i i�in hangi User de�i�kenimiz uygun pozisyonda diye bak�yoruz.
        If User(i).FreeSocket = True Then
            ' Uygun olan User de�i�kenimizi bulduktan sonra
            UserSocket(i).Accept requestID  ' o User'�n ID'si ile UserSoketi'ne iste�i kab�l olarak aktar�yoruz.
            User(i).FreeSocket = False      ' ve soketinin durumunu FALSE diyerek dolu konuma getiriyoruz.
            Log "Socket ID: " & i & " ve IP: " & UserSocket(i).RemoteHostIP & " olan ba�lant� kuruldu."
            DoEvents
            strAuthString = GenerateAuthString(i)
            SendData GenerateAuthString(i), i   ' �ifrelenmi� do�rulama kodunu ba�lant�s� sa�lanan kullan�c�ya g�nderiyoruz.
            Exit Sub ' di�er User de�i�kenlere bakmaya gerek yok.
        End If
    Next i
    
    'E�er kullan�c� kotam�z doluysa kullan�c�n ba�lant�s�n� kopar�yoruz.
    Log "IP: " & ServerSocket.RemoteHostIP & " ile ba�lant� iste�i atan kullan�c� kota dolu oldu�u i�in kabul edilemedi."
    ServerSocket.Close     ' ve sunucuyu kapat�p
    ServerSocket.Listen    ' tekrar ba�lant� isteklerini dinler olarak a��yoruz.
End Sub

Private Sub UserSocket_Close(Index As Integer)
    UserSocket(Index).Close              ' Soketin kapat�ld���na b�ylelikle emin oluyoruz.
    User(Index).FreeSocket = True        ' Soketin sahibi olan User'in soket durumunu TRUE diyerek uygun duruma
    User(Index).HasAuthenticated = False ' ve do�rulamas�n�da FALSE diyerek do�rulanmam�� duruma ayarl�yoruz.
    
    Dim userIndexInList As String
    Dim i As Integer
    
    For i = 0 To lstUser.ListCount - 1
        userIndexInList = lstUser.ListColumnText(i, 1)
        If userIndexInList = Index Then
            lstUser.Selected = i
            lstUser.RemoveSelected
            Log "Socket ID: " & Index & " ba�lant�l� kullan�c�n�n ba�lant�s� kapat�ld�."
            Exit Sub
        End If
    Next i
    
    
End Sub

Private Sub UserSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String, SplitData() As String, SplitRequest() As String
    Dim strAuthString As String
    On Error GoTo errServer

    'strData kullan�c�dan gelen verinin ham halidir. SplitData ise DATA_DELIMITER ile ay�klanm�� halidir.
    'Verilerin b�l�nmesi, soket'in dizilerinin birbirine kar��t�rmas�n� engeller.
    
    UserSocket(Index).GetData strData
    LogRAW strData
    
    SplitData = Split(strData, DATA_DELIMITER)
    
    For i = 0 To UBound(SplitData) - 1          ' gelen t�m verileri bak�yoruz.
        SplitRequest = Split(SplitData(i), "|") ' pipe(|) karakteri ile alt verileri ay�r�yoruz.
        
        If User(Index).HasAuthenticated = False Then ' E�er kullan�c� do�rulanmam��sa
            ' Gelen verilerden bu kullan�c�n�n do�rulamas�n� yapmaya �al���yoruz.
            If CheckAuthentication(SplitRequest(0), Index) = True Then ' do�rulama kodu ge�erli ise
                User(Index).HasAuthenticated = True ' kullan�c�n�n do�rulanm�� oldu�unu TRUE diyerek ayarl�yoruz.
                Log "Socket ID: " & Index & " olan kullan�c�n�n do�rulamas� ba�ar�l�."
                SendData "AUTHENTICATION|GRANTED", Index ' do�rulamas�n� kabul etti�imizi kendisine bildiriyoruz.
                nIndex = lstUser.AddItem(Index & ";" & UserSocket(Index).RemoteHost & ";" & UserSocket(Index).RemoteHostIP)
            Else
                User(Index).HasAuthenticated = False ' do�rulamas� ge�ersiz ise FALSE diyerek do�rulanmam�� olarak ayarl�yoruz.
                SendData "AUTHENTICATION|DENIED", Index ' do�rulamas�n� kabul etmedi�imizi kendisine bildiriyoruz.
                Log "Socket ID: " & Index & " olan kullan�c�n�n do�rulamas� ba�ar�s�z oldu�undan ba�lant�s� kesilmi�tir."
                DisconnectUser Index 'kullan�c�n�n ba�lant�s�n� d���recek fonsiyonumuza i�i devrediyoruz.
            End If
        Else ' E�er kullan�c�m�z hali haz�rda do�rulanm�� bir kullan�c� ise
             ' pipe(|) karakteri ile alt verilerine ay�r�p
             ' program�n normal ak���n� buraya programlayaca��z.
        End If
    Next i

errServer:
End Sub

Private Sub UserSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    UserSocket_Close Index ' kullan�c�n�n soketinde bir sorun olursa direk soketini kapat�yoruz. Hi� u�ra�m�yoruz.
End Sub

Private Sub StartServer()
    ServerSocket.Listen 'Sunucu Soketimizi dinleme konumuna al�yoruz.
    Log "Sunucu ba�ar�yla ba�lat�ld�." 'Ekran�m�zda sunucu soketimizin ba�lat�ld���n� kayd�n� ge�iyoruz.
End Sub

Private Sub StopServer()
    ServerSocket.Close 'Ba�lant� isteklerini dinleyen soketi kapat�yoruz.
    For i = 0 To MAX_USERS 'Do�rulamas�n� yapm�� ba�l� olan t�m kullan�c�lar�n ba�lant�s�n� d���rece�iz.
        If User(i).FreeSocket = False Then 'E�er kullan�c�n�n ba�lant�s� bo� de�ilse
            User(i).FreeSocket = True 'Bo� konuma al�yoruz
            User(i).HasAuthenticated = False 'Do�rulamas�n� do�rulanmam�� olarak de�i�tiriyoruz.
            UserSocket(i).Close 'Soketini kapat�yoruz.
        End If
    Next i 'Bir d�ng� ile t�m kullan�c�lar i�in yukar�ki ad�mlar� ger�ekle�tiriyoruz.
    Log "Sunucu ba�ar�yla durduruldu." 'B�t�m i�lemler bittikten sonra ekan�m�zda sunucunun kapat�ld��� kayd�n� ge�iyoruz.
End Sub
