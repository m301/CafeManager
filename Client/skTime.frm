VERSION 5.00
Begin VB.Form skTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: More Time :::"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox tMM 
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Min."
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter Minutes to Increase/Renew ?"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "skTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
If tMM.Text <> "" Then
Winsock.wSend "tmsk" & tMM.Text
End If
Unload Me
End Sub

Private Sub tMM_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select
End Sub
