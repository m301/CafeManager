VERSION 5.00
Begin VB.Form cToken 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Token :::"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frM1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox rePin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   15
         Text            =   "pin"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox cPin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   14
         Text            =   "pin"
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox cName 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Text            =   "User Name"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ComboBox idType 
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Text            =   "Voter ID"
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox idNum 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Text            =   "XYZ --ID Number"
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox Mob 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000-000-0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   360
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "1211043210"
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox sMm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "0"
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox shH 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "1"
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Re-Pin :"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Pin :"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   2895
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Token Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Tkn 
         Caption         =   "User Name"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "ADD Token"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Hr."
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Min."
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   3960
         Width           =   375
      End
      Begin VB.Shape Shape1 
         Height          =   2535
         Left            =   240
         Top             =   960
         Width           =   4215
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   240
         Top             =   3720
         Width           =   4215
      End
   End
End
Attribute VB_Name = "cToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cName_KeyUp(KeyCode As Integer, Shift As Integer)
Tkn.Caption = cName.Text
End Sub

Private Sub Form_Load()
idType.AddItem "Voter ID"
idType.AddItem "Telephone Bill"
idType.AddItem "Electric Bill"
idType.AddItem "Adhar Card/UID"
idType.AddItem "Identity Card"
idType.AddItem "Passport"
idType.AddItem "Credit/Debit Card"
idType.AddItem "PAN Card"


End Sub

Private Sub OKButton_Click()
If idNum.Text = "ID Number" Then
MsgBox "Please enter correct " & idType.Text, vbCritical + vbOKOnly, "ID Number"

ElseIf sMm.Text = "" Or shH.Text = "" Or (sMm.Text = "0" And shH.Text = "0") Then
MsgBox "Please enter correct time", vbCritical + vbOKOnly, "Incorrect Time"

ElseIf cPin.Text <> rePin.Text Then
MsgBox "Please enter correct confirmation pin !", vbInformation + vbOKOnly, "Re-Enter Pin "
Else
With cWinsock.lvwToken
.ListItems.Add , , .ListItems.Count + 1
.ListItems(.ListItems.Count).SubItems(1) = cName.Text
.ListItems(.ListItems.Count).SubItems(2) = Val(shH.Text) * 60 + Val(sMm.Text)
.ListItems(.ListItems.Count).SubItems(3) = cName.Text
.ListItems(.ListItems.Count).SubItems(4) = cPin.Text
.ListItems(.ListItems.Count).SubItems(5) = idType.Text
.ListItems(.ListItems.Count).SubItems(6) = idNum.Text
.ListItems(.ListItems.Count).SubItems(7) = Mob.Text
End With
cWinsock.lstSave
Unload Me


End If

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
Private Sub Mob_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select
End Sub


