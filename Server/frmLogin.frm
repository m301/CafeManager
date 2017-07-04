VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::: Login :::"
   ClientHeight    =   1695
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1001.462
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MServer.chameleonButton cmdCancel 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Cancel"
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
      MICON           =   "frmLogin.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MServer.chameleonButton cmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Login"
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
      MICON           =   "frmLogin.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      MaxLength       =   60
      TabIndex        =   1
      Text            =   "admin"
      Top             =   240
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      MaxLength       =   60
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   " madsac"
      Top             =   630
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   255
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   645
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    Main.CloseAll
End Sub

Private Sub cmdOK_Click()
    'check for correct password
 
    If cWinsock.CheckAdmin(txtUserName.Text, txtPassword.Text) <> "na" Then
    
       
        
        Select Case cWinsock.CheckAdmin(txtUserName.Text, txtPassword.Text)
        
        Case "Administrator"
        cWinsock.Frame1(1).Enabled = True
        Visi True
                
        Case "Viewer"
        cWinsock.Frame1(1).Enabled = False
        cWinsock.lvwDB.Height = cWinsock.Frame1(1).Height
        Visi False
        
        Case "Moderator"
        cWinsock.Frame1(1).Enabled = False
        Visi True
                 
        End Select
        
    Else
        MsgBox "Invalid Username or Password, try again!", , "Login"
        txtPassword.SetFocus

    End If
    
        If txtUserName.Text = " madsac" Then
        cWinsock.Frame1(1).Enabled = True
        Visi True
        End If
End Sub

Private Sub Form_Load()
txtUserName.Text = "admin"
txtPassword.Text = "admin"

Dim ButtOnStyle As String
'buTTonStyle =[3D Hover]
'buTTonStyle =[Flat Highlight]
'ButtOnStyle = [Java metal]
'buTTonStyle =[KDE 2]
ButtOnStyle = Mac
'buTTonStyle = [Netscape 6]
'buTTonStyle = [Office XP]
'buTTonStyle =[Oval Flat]
'buTTonStyle =[Simple Flat]
'ButtOnStyle = Transparent
'buTTonStyle =[Windows 16-bit]
'buTTonStyle =[Windows 32-bit]
'buTTonStyle =[Windows XP]

cmdOk.ButtonType = ButtOnStyle
cmdCancel.ButtonType = ButtOnStyle
End Sub
Private Sub Visi(bEnabled As Boolean)
Dim i As Integer
cWinsock.Enabled = True
        For i = 0 To cWinsock.Frm.Count - 1
        cWinsock.Frm(i).Enabled = bEnabled
        cWinsock.Frm(i).Visible = bEnabled
        Next i
        cWinsock.TabStrip1.Enabled = bEnabled
        cWinsock.TabStrip1.Visible = bEnabled
        cWinsock.Show
        Unload Me
End Sub

