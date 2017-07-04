VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cRDP 
   Caption         =   "Remote Desktop - MAD Cafe Manager"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14130
   ClipControls    =   0   'False
   Icon            =   "cRDP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   14130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Left            =   10320
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   10680
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   9480
      Top             =   120
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   60
         SelStart        =   10
         Value           =   10
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   375
         Left            =   6960
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh Interval :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   11055
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Max             =   100
         SelStart        =   75
         Value           =   75
      End
      Begin MSComctlLib.ProgressBar PgB1 
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quality :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Image wAd 
      Height          =   5775
      Left            =   120
      Stretch         =   -1  'True
      Top             =   840
      Width           =   10815
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   9480
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "cRDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'// PROGRAM SETTINGS
Dim TCP_IP As String
Const TCP_PORT              As Long = 1003

Private ReceiveData         As String
Private FileBytes           As Long
Dim weRR As Integer
Dim dComp As Boolean

' Note that this file is 2M, so you might want to try with something simpler

Public Function wRRef()
If Winsock1.State = 7 Then
cWinsock.wSend "rrdp" & Slider1.Value, cWinsock.GETSOCK(cWinsock.lvwDB.SelectedItem.Index)
Else
TCP_IP = cWinsock.sock1(cWinsock.GETSOCK(cWinsock.lvwDB.SelectedItem.Index)).RemoteHostIP
Winsock1.Close
With Winsock1
        .RemoteHost = TCP_IP
        .RemotePort = TCP_PORT
        .Connect
    End With

End If

End Function

Public Function sGet(cIP As String, sName As String)

End Function

Private Sub Command1_Click()
weRR = 0
wRRef
Label3.Caption = "Selected : " & cWinsock.lblsEl.Caption
End Sub



Private Sub Form_Resize()
On Error Resume Next
wAd.Width = Me.Width - 100
wAd.Height = Me.Height - Frame1.Height - Frame2.Height * 2
Frame1.Top = wAd.Top + wAd.Height
Frame1.Width = wAd.Width
Frame2.Width = wAd.Width
PgB1.Width = Frame1.Width - PgB1.Left - 100
End Sub

Private Sub Slider2_Click()
Timer1.Interval = Slider2.Value * 1000
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Winsock1.State <> 7 Then wRRef
Label3.Caption = "Selected : " & cWinsock.lblsEl.Caption
End Sub


'// CLOSE WINSOCK
Private Sub winsock1_Close()
    Dim strTempFile As String
    Dim intFile As Integer
    If Winsock1.State = sckClosing Then
        Winsock1.Close
    End If
    strTempFile = App.Path & "\tmp.jpg"
    intFile = FreeFile
    Open strTempFile For Binary As #intFile
    Put #intFile, , ReceiveData
    Close #intFile
    ReceiveData = ""
Me.Caption = "Done"
FileBytes = 0
wAd.Picture = LoadPicture(App.Path & "\tmp.jpg")
Winsock1.Close
End Sub

'// DATA ARRIVES
Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Winsock1.GetData strData
    ReceiveData = ReceiveData & strData
    FileBytes = FileBytes + bytesTotal

 Me.Caption = "Downloading: " & FileBytes & " bytes"
End Sub

'// ERROR
Private Sub winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock1.Close
    MsgBox "Error Connecting!"
End Sub

'// FORM LOAD
Private Sub Form_Load()
    'Timer1.Enabled = True
  '  Timer1.Interval = 100
weRR = 0
wRRef
Label3.Caption = "Selected : " & cWinsock.lblsEl.Caption
End Sub

'// UNLOAD FORM
Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
End Sub


