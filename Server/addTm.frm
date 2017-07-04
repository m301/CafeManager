VERSION 5.00
Begin VB.Form addTm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ":: Add Time ::"
   ClientHeight    =   1830
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox sMm 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox shH 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "1"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Selected"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Hr."
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Min."
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "addTm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label3.Caption = cWinsock.lblsEl.Caption
End Sub

Private Sub OKButton_Click()
Dim mIn, hR, th As Integer
Dim data As String
data = Now & " | " & cWinsock.lblsEl.Caption & " | " & sMm.Text & " min. | " & shH.Text & " hr. | "
cWinsock.cLog (data)
mIn = sMm.Text
hR = shH.Text
cWinsock.lvwDB.ListItems.Item(cWinsock.SEL).SubItems(4) = Val(cWinsock.lvwDB.ListItems.Item(cWinsock.SEL).SubItems(4)) + mIn + hR * 60
'''''''''''''''''''''''''''''
''Convert String To Integer''
'''''''''''''''''''''''''''''

cWinsock.wSend "tadd" & th + mIn + hR * 60, cWinsock.SEL
Unload Me
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

Private Sub sMm_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select
End Sub
