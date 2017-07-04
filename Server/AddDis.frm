VERSION 5.00
Begin VB.Form AddDis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::: Add Discount :::"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox shH 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "10"
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Discount To :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label cCurrency 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "AddDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cWinsock.lvwDB.SelectedItem.SubItems(16) = shH.Text
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.cCurrency.Caption = cWinsock.uCurrency
Label1.Caption = cWinsock.lvwDB.SelectedItem.SubItems(2)
End Sub

Private Sub shH_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select
End Sub


