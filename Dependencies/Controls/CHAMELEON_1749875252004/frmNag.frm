VERSION 5.00
Begin VB.Form frmNag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agreement"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "I DONT AGREE"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   7260
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "I AGREE TO THE TERMS"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7260
      Width           =   2415
   End
   Begin prjChameleon.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   7200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Continue"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmNag.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Left            =   6360
      Top             =   360
   End
   Begin VB.Label Label1 
      Caption         =   "PLEASE TAKE A MOMENT AND REVIEW THE TERMS BEFORE USING"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   6705
      Left            =   120
      Picture         =   "frmNag.frx":001C
      Top             =   480
      Width           =   6600
   End
End
Attribute VB_Name = "frmNag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chameleonButton1_Click()
     If Check1.Value Then
        SaveSetting App.Title, "NAG", "SHOWN", 1
        Unload Me
     ElseIf Check2.Value Then
        MsgBox "PLEASE DELETE THE CHAMELEON BUTTON SOURCE FROM YOUR COMPUTER." & vbCrLf & vbCrLf & "ONCE YOU UNDERSTAND THE OPEN SOURCE PHILOSOPHY YOU CAN USE IT AGAIN, THANKS", vbInformation
        End
     End If
End Sub

Private Sub Check1_Click()
    chameleonButton1.Enabled = Check1.Value Or Check2.Value
    Check2.Value = Abs(Not CBool(Check1.Value))
End Sub

Private Sub Check2_Click()
    chameleonButton1.Enabled = Check1.Value Or Check2.Value
    Check1.Value = Abs(Not CBool(Check2.Value))
End Sub
