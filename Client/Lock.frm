VERSION 5.00
Object = "*\A..\..\VB6\zip\AlphaImgControl\AlphaImgControl\LaVolpeAlphaImg.vbp"
Begin VB.Form sLoCk 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer clock 
      Interval        =   60000
      Left            =   1800
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Image           =   "Lock.frx":0000
      Attr            =   513
      Effects         =   "Lock.frx":1FFC2
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "sLoCk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public vIs As Boolean

Private Sub Command1_Click()
Call xLoCk
Main.clAll

End Sub

Private Sub Form_Load()
vIs = True
AntiTaskManagerController False
sLoCk.Width = Screen.Width
sLoCk.Height = Screen.Height
sLoCk.Left = 0
sLoCk.Top = 0
 Dim lR As Long
 lR = SetTopMostWindow(Me.hWnd, True)
Shell "taskkill /f /im explorer.exe "
End Sub



Private Function xLoCk()
AntiTaskManagerController True
sLoCk.Width = 1000
sLoCk.Height = 1000
 Dim lR As Long
 lR = SetTopMostWindow(Me.hWnd, False)
Shell "explorer.exe"
vIs = False
End Function

