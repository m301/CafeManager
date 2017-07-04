VERSION 5.00
Begin VB.Form cNoti 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ":: Time Over ::"
   ClientHeight    =   1845
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Yes"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   240
      Picture         =   "cNoti.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "cNoti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mType As String

Private Sub Command1_Click()
Bmet.uClick = True
If Me.Caption = "Logout " Then
sLoCk.Show
Winsock.wSend "usrl"
End If
Select Case mType
Case "ask"
skTime.Show
End Select
Unload Me
End Sub

Private Sub Command2_Click()
Bmet.uClick = False
Unload Me
End Sub

Private Sub Command3_Click()
Bmet.uClick = False
Unload Me
End Sub

Private Sub Form_Load()
Dim lR As Long
lR = SetTopMostWindow(Me.hwnd, True)

Bmet.uClick = False
End Sub

Private Sub OKButton_Click()
Bmet.uClick = False
Unload Me
End Sub

Public Sub cMessage(message As String, Title As String, button As Integer, Optional Message_Type As String)
Me.Label1.Caption = message
Me.Caption = Title
mType = Message_Type
Select Case button
Case "0"
 Command1.Visible = False
 Command3.Visible = False
 Command2.Caption = "OK"
End Select
Me.Show
Bmet.uClick = False
End Sub

