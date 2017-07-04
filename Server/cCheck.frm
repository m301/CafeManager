VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form cCheck 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::: Cafe Details :::"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MServer.chameleonButton Command3 
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   4680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "cCheck.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MServer.chameleonButton Command5 
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   4200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Register / Get Your Own Cafe Code"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "cCheck.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MServer.chameleonButton Command2 
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Check For Update"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "cCheck.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MServer.chameleonButton Command1 
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Get Details"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "cCheck.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtResponse 
      Height          =   1095
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "cCheck.frx":0070
      Top             =   5400
      Width           =   5655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   600
      Top             =   120
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "9451430071"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1320
      TabIndex        =   1
      Text            =   "test"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "cCheck.frx":0076
      Top             =   6480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MServer.chameleonButton Command4 
      Height          =   375
      Left            =   1080
      TabIndex        =   15
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "cCheck.frx":0083
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAD Cafe Manger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   240
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   240
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address        :"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone           :"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name :"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cafe name    :"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cafe Code :                                        Pin :"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Cafe Name"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Phone Number"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Cafe Address"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
   End
End
Attribute VB_Name = "cCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bClick As Boolean
Private blnConnected, Q As Boolean
Dim strResponse As String
'**************************************************************************'
'                          URL Execute Declarations                                     '
'**************************************************************************'
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Public Function uOpen(ByVal URL As String) As Boolean 'Opens URL
Dim res As Long
' it is mandatory that the URL is prefixed with http:// or https://
If InStr(1, URL, "http", vbTextCompare) <> 1 Then
URL = "http://" & URL
End If
res = ShellExecute(0&, "open", URL, vbNullString, vbNullString, _
vbNormalFocus)
OpenBrowser = (res > 32)
End Function

Private Sub Command5_Click()
uOpen "http://madsacsoft.com/cafe/register"
End Sub

Private Sub Timer1_Timer()
Command1.Enabled = True
Command2.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
End Sub

Private Sub winsock1_Connect()
    blnConnected = True
End Sub

' this event occurs when data is arriving via winsock1
Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)

        Winsock1.GetData strResponse, vbString, bytesTotal
'MsgBox (strResponse)
    strResponse = FormatLineEndings(strResponse)
    ' we append this to the response box becuase data arrives
    ' in multiple packets
    
    'txtResponse.Text = txtResponse.Text & strResponse
    Dim aRR() As String
    Dim sPl() As String
    sPl = Split(strResponse, "|^^!MAD!^^|")
    strResponse = sPl(1)
    txtResponse.Text = txtResponse.Text & strResponse
    If bClick = True Then
    
    If txtResponse.Text = "0" Then
        MsgBox "Update not available", vbInformation + vbOKOnly, "Update"
    ElseIf txtResponse.Text = "1" Then
        MsgBox "Update available !", vbInformation + vbOKOnly, "Update"
    ElseIf txtResponse.Text = "0000" Then
        MsgBox "Incorrect Cafe ID Or PIN !", vbCritical + vbOKOnly, "Incorrect Information !"
    ElseIf Left(txtResponse.Text, 7) = "|^MAD^|" Then


        sPl = Split(txtResponse.Text, "|^END_MAD^|")
        aRR = Split(sPl(0), "|^MAD^|")
        Label1.Caption = aRR(1)
        Label2.Caption = aRR(2)
        Label3.Caption = aRR(3)
        Label4.Caption = aRR(4)
        cWinsock.cfeCode.Caption = Text2.Text
        cWinsock.CfeOwner.Caption = Label1.Caption
        cWinsock.cfeAddress.Caption = Label2.Caption
        cWinsock.cfePhone.Caption = Label3.Caption
        cWinsock.cfeName.Caption = Label4.Caption
   
        Text2.Enabled = False
        Text3.Enabled = False
        Command4.Visible = True
        Command1.Visible = False

    End If
    End If

Main.saveSetting
Command1.Enabled = True
Command2.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
End Sub

Private Sub winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock1.Close
End Sub

Private Sub winsock1_Close()
    blnConnected = False
    Winsock1.Close
End Sub


Private Sub Command1_Click()
bClick = True
Command1.Enabled = False
Command2.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Timer1.Enabled = True
SendR "http://madsac.in/cafe/info.php", "id=" & Text2.Text & "&p=" & Text3.Text

End Sub

Private Sub Command2_Click()
Command1.Enabled = False
Command2.Enabled = False
Text2.Enabled = False
Text3.Enabled = False

Timer1.Enabled = True
bClick = True
SendR "http://madsac.in/cafe/update.php", "ver=2"
End Sub

Private Sub Command3_Click()
Unload Me
Unload frmSplash
End Sub

Private Sub Command4_Click()
cWinsock.cfeCode.Caption = Text2.Text
cWinsock.cfeName.Caption = Label4.Caption
Main.saveSetting
Main.Visible = True
Main.Visible = False
Unload Me
End Sub

Private Sub Form_Load()
bClick = False
blnConnected = False
End Sub

' this function converts all line endings to Windows CrLf line endings
Private Function FormatLineEndings(ByVal str As String) As String
    Dim prevChar As String
    Dim nextChar As String
    Dim curChar As String
    
    Dim strRet As String
    
    Dim X As Long
    
    prevChar = ""
    nextChar = ""
    curChar = ""
    strRet = ""
    
    For X = 1 To Len(str)
        prevChar = curChar
        curChar = Mid$(str, X, 1)
                
        If nextChar <> vbNullString And curChar <> nextChar Then
            curChar = curChar & nextChar
            nextChar = ""
        ElseIf curChar = vbLf Then
            If prevChar <> vbCr Then
                curChar = vbCrLf
            End If
            
            nextChar = ""
        ElseIf curChar = vbCr Then
            nextChar = vbLf
        End If
        
        strRet = strRet & curChar
    Next X
    
    FormatLineEndings = strRet
End Function

' this function sends the HTTP request
Private Function SendR(To_URL As String, strData As String)
Dim eUrl As URL
Dim strHeaders As String
Dim strHTTP As String
Dim X As Integer
    Q = False
   
    If blnConnected Then Exit Function
    
    ' get the url
    eUrl = ExtractUrl(To_URL)
    
    If eUrl.Host = vbNullString Then
        Exit Function
    End If
    
    ' configure winsock1
    Winsock1.Protocol = sckTCPProtocol
    Winsock1.RemoteHost = eUrl.Host
    
    If eUrl.Scheme = "http" Then
        If eUrl.Port > 0 Then
            Winsock1.RemotePort = eUrl.Port
        Else
            Winsock1.RemotePort = 80
        End If
    ElseIf eUrl.Scheme = vbNullString Then
        Winsock1.RemotePort = 80
    Else
    End If
    
    ' build encoded data the data is url encoded in the form
    ' var1=value&var2=value
   
    If eUrl.Query <> vbNullString Then
        eUrl.URI = eUrl.URI & "?" & eUrl.Query
    End If
    
    ' check if any variables were supplied
    If strData <> vbNullString Then
            If eUrl.Query <> vbNullString Then
                eUrl.URI = eUrl.URI & "&" & strData
            Else
                eUrl.URI = eUrl.URI & "?" & strData
            End If
    End If

    txtResponse.Text = ""
    
    ' build the HTTP request in the form
    '
    ' {REQ METHOD} URI HTTP/1.0
    ' Host: {host}
    ' {headers}
    '
    ' {post data}
   ' MsgBox (eUrl.URI)
   strHTTP = "GET" & " " & eUrl.URI & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & eUrl.Host & vbCrLf
    strHTTP = strHTTP & strHeaders
    strHTTP = strHTTP & vbCrLf

       
    Winsock1.Connect
    
    ' wait for a connection
    While Not blnConnected
        DoEvents
    Wend
    
    ' send the HTTP request
    Winsock1.SendData strHTTP
End Function






