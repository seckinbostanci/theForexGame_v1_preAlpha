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
    If btnServer.State = False Then    ' Eðer sunucuyu durdurup baþlatacak düðmenin üzerinde "&Baþlat" yazýyorsa;
        StartServer                             ' Sunucuyu baþlatacak olan fonksiyonumuzu çaðýrýyoruz.
        btnServer.Caption = "Running..."        ' Düðmemizin üzerindeki yazýyý "&Durdur" olarak deðiþtiriyoruz.
    Else                                        ' Eðer sunucuyu durdurup baþlatacak düðmenin üzerinde "&Baþlat" yazmýyorsa;
        StopServer                              ' Sunucuyu durduracak olan fonksiyonumuzu çaðýrýyoruz.
        btnServer.Caption = "Stoped."        ' Düðmemizin üzerindeki yazýyý "&Baþlat" olarak deðiþtiriyoruz.
    End If
End Sub

Private Sub Form_Load()
    For i = 1 To MAX_USERS          ' MainModule'de tanýlý olan en yüksek kullanýcý sayýsý kadar oluþturulmuþ olan
        User(i).FreeSocket = True   ' Tüm kullanýcý deðiþkenlerimizin soketlerini TRUE diyerek boþ olarak ayarlýyoruz.
        Load UserSocket(i)          ' ve her kullanýcýnýn user soketini yaratýyoruz.
    Next i
    User(0).FreeSocket = True       ' ilk kullanýcýmýzýn soketini TRYE diyerek boþ olarak ayarlýyoruz.
    StartServer                     ' Sunucuyu baþlatacak olan fonksiyonumuzu çaðýrýyoruz.
    btnServer.State = False
End Sub

Private Sub Form_Terminate()
    Log "Uygulama kapatýlmaya hazýrlanýyor..."
    StopServer  ' Sunucuyu kapatacak olan fonksiyonumuzu çaðýrýyoruz.
    Log "Uygulama kapatýldý."
    Unload Me   ' Form'umuzu ram bellekten çýkarýyoruz.
    End         ' Uygulamamýzý yaþam döngüsünü bitiriyoruz.
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
        'Yeni gelen baðlantý isteði için hangi User deðiþkenimiz uygun pozisyonda diye bakýyoruz.
        If User(i).FreeSocket = True Then
            ' Uygun olan User deðiþkenimizi bulduktan sonra
            UserSocket(i).Accept requestID  ' o User'ýn ID'si ile UserSoketi'ne isteði kabül olarak aktarýyoruz.
            User(i).FreeSocket = False      ' ve soketinin durumunu FALSE diyerek dolu konuma getiriyoruz.
            Log "Socket ID: " & i & " ve IP: " & UserSocket(i).RemoteHostIP & " olan baðlantý kuruldu."
            DoEvents
            strAuthString = GenerateAuthString(i)
            SendData GenerateAuthString(i), i   ' Þifrelenmiþ doðrulama kodunu baðlantýsý saðlanan kullanýcýya gönderiyoruz.
            Exit Sub ' diðer User deðiþkenlere bakmaya gerek yok.
        End If
    Next i
    
    'Eðer kullanýcý kotamýz doluysa kullanýcýn baðlantýsýný koparýyoruz.
    Log "IP: " & ServerSocket.RemoteHostIP & " ile baðlantý isteði atan kullanýcý kota dolu olduðu için kabul edilemedi."
    ServerSocket.Close     ' ve sunucuyu kapatýp
    ServerSocket.Listen    ' tekrar baðlantý isteklerini dinler olarak açýyoruz.
End Sub

Private Sub UserSocket_Close(Index As Integer)
    UserSocket(Index).Close              ' Soketin kapatýldýðýna böylelikle emin oluyoruz.
    User(Index).FreeSocket = True        ' Soketin sahibi olan User'in soket durumunu TRUE diyerek uygun duruma
    User(Index).HasAuthenticated = False ' ve doðrulamasýnýda FALSE diyerek doðrulanmamýþ duruma ayarlýyoruz.
    
    Dim userIndexInList As String
    Dim i As Integer
    
    For i = 0 To lstUser.ListCount - 1
        userIndexInList = lstUser.ListColumnText(i, 1)
        If userIndexInList = Index Then
            lstUser.Selected = i
            lstUser.RemoveSelected
            Log "Socket ID: " & Index & " baðlantýlý kullanýcýnýn baðlantýsý kapatýldý."
            Exit Sub
        End If
    Next i
    
    
End Sub

Private Sub UserSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String, SplitData() As String, SplitRequest() As String
    Dim strAuthString As String
    On Error GoTo errServer

    'strData kullanýcýdan gelen verinin ham halidir. SplitData ise DATA_DELIMITER ile ayýklanmýþ halidir.
    'Verilerin bölünmesi, soket'in dizilerinin birbirine karýþtýrmasýný engeller.
    
    UserSocket(Index).GetData strData
    LogRAW strData
    
    SplitData = Split(strData, DATA_DELIMITER)
    
    For i = 0 To UBound(SplitData) - 1          ' gelen tüm verileri bakýyoruz.
        SplitRequest = Split(SplitData(i), "|") ' pipe(|) karakteri ile alt verileri ayýrýyoruz.
        
        If User(Index).HasAuthenticated = False Then ' Eðer kullanýcý doðrulanmamýþsa
            ' Gelen verilerden bu kullanýcýnýn doðrulamasýný yapmaya çalýþýyoruz.
            If CheckAuthentication(SplitRequest(0), Index) = True Then ' doðrulama kodu geçerli ise
                User(Index).HasAuthenticated = True ' kullanýcýnýn doðrulanmýþ olduðunu TRUE diyerek ayarlýyoruz.
                Log "Socket ID: " & Index & " olan kullanýcýnýn doðrulamasý baþarýlý."
                SendData "AUTHENTICATION|GRANTED", Index ' doðrulamasýný kabul ettiðimizi kendisine bildiriyoruz.
                nIndex = lstUser.AddItem(Index & ";" & UserSocket(Index).RemoteHost & ";" & UserSocket(Index).RemoteHostIP)
            Else
                User(Index).HasAuthenticated = False ' doðrulamasý geçersiz ise FALSE diyerek doðrulanmamýþ olarak ayarlýyoruz.
                SendData "AUTHENTICATION|DENIED", Index ' doðrulamasýný kabul etmediðimizi kendisine bildiriyoruz.
                Log "Socket ID: " & Index & " olan kullanýcýnýn doðrulamasý baþarýsýz olduðundan baðlantýsý kesilmiþtir."
                DisconnectUser Index 'kullanýcýnýn baðlantýsýný düþürecek fonsiyonumuza iþi devrediyoruz.
            End If
        Else ' Eðer kullanýcýmýz hali hazýrda doðrulanmýþ bir kullanýcý ise
             ' pipe(|) karakteri ile alt verilerine ayýrýp
             ' programýn normal akýþýný buraya programlayacaðýz.
        End If
    Next i

errServer:
End Sub

Private Sub UserSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    UserSocket_Close Index ' kullanýcýnýn soketinde bir sorun olursa direk soketini kapatýyoruz. Hiç uðraþmýyoruz.
End Sub

Private Sub StartServer()
    ServerSocket.Listen 'Sunucu Soketimizi dinleme konumuna alýyoruz.
    Log "Sunucu baþarýyla baþlatýldý." 'Ekranýmýzda sunucu soketimizin baþlatýldýðýný kaydýný geçiyoruz.
End Sub

Private Sub StopServer()
    ServerSocket.Close 'Baðlantý isteklerini dinleyen soketi kapatýyoruz.
    For i = 0 To MAX_USERS 'Doðrulamasýný yapmýþ baðlý olan tüm kullanýcýlarýn baðlantýsýný düþüreceðiz.
        If User(i).FreeSocket = False Then 'Eðer kullanýcýnýn baðlantýsý boþ deðilse
            User(i).FreeSocket = True 'Boþ konuma alýyoruz
            User(i).HasAuthenticated = False 'Doðrulamasýný doðrulanmamýþ olarak deðiþtiriyoruz.
            UserSocket(i).Close 'Soketini kapatýyoruz.
        End If
    Next i 'Bir döngü ile tüm kullanýcýlar için yukarýki adýmlarý gerçekleþtiriyoruz.
    Log "Sunucu baþarýyla durduruldu." 'Bütüm iþlemler bittikten sonra ekanýmýzda sunucunun kapatýldýðý kaydýný geçiyoruz.
End Sub
