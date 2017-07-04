VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Bmet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::  Details  ::"
   ClientHeight    =   5940
   ClientLeft      =   3330
   ClientTop       =   2940
   ClientWidth     =   14565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   Begin MClient.chameleonButton Command2 
      Height          =   495
      Left            =   2880
      TabIndex        =   40
      Top             =   5280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Close"
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
      MICON           =   "bM.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MClient.chameleonButton aExit 
      Height          =   495
      Left            =   600
      TabIndex        =   39
      Top             =   5280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   2
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
      MICON           =   "bM.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MClient.chameleonButton Command4 
      Height          =   495
      Left            =   6000
      TabIndex        =   38
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Increase Time"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "bM.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MClient.chameleonButton Command3 
      Height          =   495
      Left            =   4440
      TabIndex        =   37
      Top             =   4200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Logout / Lock"
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
      MICON           =   "bM.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox usrType 
      Height          =   315
      Left            =   5640
      TabIndex        =   36
      Text            =   "Normal"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   4695
      Left            =   14040
      TabIndex        =   33
      Top             =   720
      Width           =   375
   End
   Begin SHDocVwCtl.WebBrowser wAd 
      Height          =   4695
      Left            =   8520
      TabIndex        =   31
      Top             =   720
      Width           =   5895
      ExtentX         =   10398
      ExtentY         =   8281
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   8280
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   8760
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Speed"
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "User Type"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   12
      Left            =   4320
      TabIndex        =   35
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label tNow 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      Height          =   255
      Left            =   5760
      TabIndex        =   34
      Top             =   600
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      Height          =   4935
      Left            =   8400
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "Advertisement :"
      Height          =   255
      Left            =   8280
      TabIndex        =   32
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label cFen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cafe Name"
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
      Left            =   360
      TabIndex        =   29
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label cInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Terminal Name"
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   28
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Mobile Number"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   27
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Time Remaining"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   26
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Time Used"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   25
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Total time"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   24
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Start Time"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   23
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " ID Number"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   22
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " ID Type"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   21
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Name"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   20
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label cInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "9876543210"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   19
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label cInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00 Min."
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   18
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label cInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00 Min."
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   17
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label cInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00 Min."
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   16
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label cInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label cInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ID Number"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   14
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label cInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ID Type"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   13
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label cInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Your Name"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblRecv 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      Height          =   4095
      Left            =   240
      Top             =   960
      Width           =   7815
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   5640
      X2              =   5640
      Y1              =   3000
      Y2              =   3480
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   3
      Left            =   4200
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   4200
      X2              =   7920
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   4200
      X2              =   7920
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1800
      X2              =   1800
      Y1              =   3000
      Y2              =   4920
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   2
      Left            =   360
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   360
      X2              =   4080
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   360
      X2              =   4080
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   360
      X2              =   4080
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1800
      X2              =   1800
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   1
      Left            =   360
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   360
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   360
      X2              =   4080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   360
      X2              =   4080
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   4200
      X2              =   7920
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   4200
      X2              =   7920
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   4200
      X2              =   7920
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   0
      Left            =   4200
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   5640
      X2              =   5640
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Label lblRecv2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Downloaded          "
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblInter 
      Caption         =   "Interface :::::"
      Height          =   255
      Left            =   9240
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Upload Speed  "
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Speed"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Speed"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Download Speed      "
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSent2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Uploaded        "
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Terminal Name :"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Bmet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sl As Integer
Dim r1 As Double
Dim r2 As Double
Public uClick As Boolean
Dim cx As Integer, cy As Integer, r As Single
Dim s As Double
Dim enviados, recibidos As Double
Private m_objIpHelper As CIpHelper
Private TransferRate                    As Double
Private TransferRate2                   As Double
Dim FirstTime As Boolean
Dim Pindex As Long

Private Sub aExit_Click()
Main.clAll
End Sub


Private Sub chameleonButton1_Click()

End Sub

Private Sub Command1_Click()
   enviados = m_objIpHelper.BytesSent / 1024
   recibidos = m_objIpHelper.BytesReceived / 1024
End Sub



Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Command3_Click()

   cNoti.cMessage "Are you sure to Logout ? ", "Logout ", 1
  
End Sub

Private Sub Command4_Click()
If Winsock.sock1.State <> 7 Then
cNoti.cMessage "Sorry, Server is not connected ! ", "Server Not Connected", 0
Else
skTime.Show
End If
End Sub




Private Sub Form_Load()
Bmet.wAd.Navigate (Main.vuRl3)
cFen.Caption = Winsock.Cafe_Name
If App.PrevInstance = True Then
   End
End If
DoEvents
Set m_objIpHelper = New CIpHelper
FirstTime = True
cInfo(8).Caption = Winsock.sock1.LocalHostName
Me.Width = 14600
Timer3.Enabled = True
cFen.Caption = Winsock.Cafe_Name

If Me.Width > Screen.Width Then
Shape3.Width = Shape3.Width - (Me.Width - Screen.Width) - 10
wAd.Width = Shape3.Width
Command5.Left = Command5.Left - (Me.Width - Screen.Width) - 10
Me.Width = Me.Width - (Me.Width - Screen.Width) - 10
End If
tNow.Caption = Format$(Now, "hh:mm AM/PM")
usrType.AddItem "Normal"
usrType.AddItem "Starter"
usrType.AddItem "Gamer"
usrType.AddItem "Bussiness"
usrType.AddItem "Student"
End Sub



Private Sub lblRecv_Click()
Call Command1_Click
End Sub



Private Sub lblSent_Click()
Call Command1_Click
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub UpdateInterfaceInfo()

On Error Resume Next
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
Case MIB_IF_TYPE_ETHERNET: lblInter.Caption = "Ethernet"
Case MIB_IF_TYPE_FDDI: lblInter.Caption = "FDDI"
Case MIB_IF_TYPE_LOOPBACK: lblInter.Caption = "Loopback"
Case MIB_IF_TYPE_OTHER: lblInter.Caption = "Other"
Case MIB_IF_TYPE_PPP: lblInter.Caption = "PPP"
Case MIB_IF_TYPE_SLIP: lblInter.Caption = "SLIP"
Case MIB_IF_TYPE_TOKENRING: lblInter.Caption = "TokenRing"
End Select
If FirstTime Then
   FirstTime = False
   enviados = m_objIpHelper.BytesSent / 1024
   recibidos = m_objIpHelper.BytesReceived / 1024
End If
lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived / 1024 - recibidos, "###,###,###,###,##0"))
lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent / 1024 - enviados, "###,###,###,###,##0"))
Set st_objInterface = objInterface
'---------------
blnIsRecv = (m_objIpHelper.BytesReceived / 1024 > lngBytesRecv / 1024)
blnIsSent = (m_objIpHelper.BytesSent / 1024 > lngBytesSent / 1024)
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent
DoEvents

End Sub

Private Sub Timer2_Timer()

On Error Resume Next
'DoEvents
'Dim XX As Double
'Dim YY As Double
'Dim XXX As Double
'Dim YYY As Double
'YYY = CInt(Label6.Caption)
'YY = CInt(Label5.Caption)
'DoEvents
'XX = Me.lblRecv.Caption - YY
'XXX = Me.lblSent.Caption - YYY
'DoEvents
'TransferRate = Format(Int(XX), "00.00")
'DoEvents
'TransferRate2 = Format(Int(XXX), "00.00")
'DoEvents
'Label10.Caption = TransferRate2 & " Kb/s"
'DoEvents
'Label9.Caption = TransferRate & " Kb/s"
'DoEvents
'DoEvents
'Label5.Caption = Me.lblRecv.Caption
'Label6.Caption = Me.lblSent.Caption
On Error Resume Next
Call UpdateInterfaceInfo
DoEvents
Me.lblSent2.Caption = Me.lblSent.Caption & " Kb"
Me.lblRecv2.Caption = Me.lblRecv.Caption & " Kb"
End Sub



Private Sub Timer3_Timer()
Bmet.Hide
Timer3.Enabled = False
End Sub

Private Sub usrType_Change()
Main.RefURl
End Sub

Private Sub wAd_TitleChange(ByVal Text As String)
If weRr >= 50 Then
ElseIf UCase(wAd.Document.Title) = "" Then
ElseIf UCase(wAd.Document.Title) = "MAD TIPS & TRICKS" Then
ElseIf UCase(wAd.Document.Title) = "MAD AD" Then
ElseIf UCase(wAd.Document.Title) = "AD" Then
ElseIf UCase(wAd.Document.Title) = "MADSACSOFT" Then
ElseIf UCase(wAd.Document.Title) = "MAD CAFE MANAGER" Then
Else

If Dir(App.Path & "\ad_3.html") = "" Then
Open App.Path & "\ad_3.html" For Output As #19
Print #19, "<html><body><img src=ad_3.gif width=99% height=99%> </body></html>"

Close #19
End If

wAd.Navigate (App.Path & "\ad_3.html")
weRr = weRr + 1
End If

cFen.Caption = Winsock.Cafe_Name
End Sub
