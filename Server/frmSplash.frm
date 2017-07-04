VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4305
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1320
         Top             =   480
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "By : MADSAC Soft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3240
         TabIndex        =   8
         Top             =   705
         Width           =   3150
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "LicenseTo: Testing Version "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MAD Cyber Cafe Manager"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   2640
         Width           =   6330
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   5
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4215
         TabIndex        =   4
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Company:MADSAC Soft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   2
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copyright :MADSAC Soft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   1
         Top             =   3840
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const SettingPath As String = "d:\maddb.dll"
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Main.Visible = True
     Main.Visible = False
End Sub

Private Sub Form_Load()
 Dim newLine As String
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  
If Dir(SettingPath) = "" Then

cCheck.Show
Unload Me
Else
     Main.Visible = True
     Main.Visible = False
    cRefresh
End If
End Sub
Private Function cRefresh()
 Dim newLine As String
 Main.loadSetting
    
End Function
Private Sub Frame1_Click()

If Dir(SettingPath) = "" Then
Unload Me
cCheck.Show
Else
    Unload Me
  Main.Visible = True
  Main.Visible = False
  cRefresh
End If
End Sub

Private Sub Timer1_Timer()
Frame1_Click
End Sub
Private Sub Form_Initialize()
If App.PrevInstance = True Then Main.CloseAll

If Dir("C:\WINDOWS\system32\MSWINSCK.OCX") = "" Then
Dim i() As Byte
i = LoadResData(101, "CUSTOM")
Open "C:\WINDOWS\system32\MSWINSCK.OCX" For Binary Access Write As #1
Put #1, , i
Close #1
Shell "regsvr32 /s C:\WINDOWS\system32\MSWINSCK.OCX ", vbHide
End If

If Dir("C:\WINDOWS\system32\COMCTL32.OCX") = "" Then
i = LoadResData(103, "CUSTOM")
Open "C:\WINDOWS\system32\imgctrl.OCX" For Binary Access Write As #1
Put #1, , i
Close #1
Shell "regsvr32 /s C:\WINDOWS\system32\COMCTL32.OCX ", vbHide
End If
''''''''''''''''''''''''''''' Removed .... 10mb Not comes in free for sending email :-)
If Dir("C:\WINDOWS\system32\ieframe32.dll") = "" Then
i = LoadResData(104, "CUSTOM")
Open "C:\WINDOWS\system32\ieframe32.dll" For Binary Access Write As #1
Put #1, , i
Close #1
Shell "regsvr32 /s C:\WINDOWS\system32\ieframe32.dll ", vbHide
End If
'Can Be downloaded just google download & run "regsvr32 /s C:\WINDOWS\system32\ieframe32.dll " in cmd
'''''''''''''''''''''''''''''


If Dir("C:\WINDOWS\system32\WBCustomizer.dll") = "" Then
i = LoadResData(102, "CUSTOM")
Open "C:\WINDOWS\system32\WBCustomizer.dll" For Binary Access Write As #1
Put #1, , i
Close #1
Shell "regsvr32 /s C:\WINDOWS\system32\WBCustomizer.dll ", vbHide
End If


End Sub

