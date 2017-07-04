VERSION 5.00
Begin VB.Form tmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ":: Set Time ::"
   ClientHeight    =   5475
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox shH 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "1"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox sMm 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "0"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Address 
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Text            =   "Address"
      Top             =   2160
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
      Left            =   240
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "1211043210"
      Top             =   3480
      Width           =   3975
   End
   Begin VB.TextBox idNum 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "XYZ --ID Number"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.ComboBox idType 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "Voter ID"
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox cName 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "NAME ABC"
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   120
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      Height          =   1815
      Left            =   120
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   120
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Min."
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Hr."
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   4200
      Width           =   255
   End
End
Attribute VB_Name = "tmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Unload Me
End Sub



Private Sub Form_Load()
Label3.Caption = cWinsock.lblsEl.Caption
idType.AddItem "Voter ID"
idType.AddItem "Telephone Bill"
idType.AddItem "Electric Bill"
idType.AddItem "Adhar Card/UID"
idType.AddItem "Identity Card"
idType.AddItem "Passport"
idType.AddItem "Credit/Debit Card"
idType.AddItem "PAN Card"

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



Private Sub OKButton_Click()
Dim mIn, hR As Integer
Dim data As String
Dim dSince As String
Dim xyS As Integer
If idNum.Text = "ID Number" Then
MsgBox "Please enter correct " & idType.Text, vbCritical + vbOKOnly, "ID Number"

ElseIf sMm.Text = "" Or shH.Text = "" Or (sMm.Text = "0" And shH.Text = "0") Then
MsgBox "Please enter correct time", vbCritical + vbOKOnly, "Incorrect Time"

Else
If cWinsock.sock1(cWinsock.GETSOCK(cWinsock.SEL)).State = 7 Then
mIn = sMm.Text
hR = shH.Text
data = Now & " | " & cWinsock.lblsEl.Caption & " | " & cName.Text & " | " & idType.Text & " | " & idNum.Text & " | " & Address.Text & " | " & Mob.Text & " | " & sMm.Text & "min. | " & shH.Text & " hr. | "
cWinsock.cLog (data)

xyS = (mIn + hR * 60)
cWinsock.wSend "tset" & Format$(Now, "hh:mm:ss AM/PM") & " | " & cWinsock.lblsEl.Caption & " | " & cName.Text & " | " & idType.Text & " | " & idNum.Text & " | " & Mob.Text & " | " & mIn + hR * 60, cWinsock.SEL


dSince = Format$(Now, "hh:mm:ss AM/PM")

With cWinsock.lvwDB.ListItems.Item(cWinsock.SEL)
    .SubItems(11) = 3
    .SubItems(12) = 0
    .SubItems(2) = cName.Text
    .SubItems(3) = dSince
    .SubItems(4) = xyS
    .SubItems(5) = dTTime - .SubItems(12)
    .SubItems(6) = " 0"
    .SubItems(7) = " 0"
    .SubItems(8) = idType.Text
    .SubItems(9) = idNum.Text
    .SubItems(10) = Mob.Text
End With

Else
 Dim Mres As Integer
 Mres = MsgBox(" Selected terminal is not connected. Do you want to continue ? ", vbInformation + vbYesNo, "Not connected")
 Select Case Mres
Case 6
        mIn = sMm.Text
hR = shH.Text
data = Now & " | Offline -" & cWinsock.lblsEl.Caption & " | " & cName.Text & " | " & idType.Text & " | " & idNum.Text & " | " & Address.Text & " | " & Mob.Text & " | " & sMm.Text & "min. | " & shH.Text & " hr. | "
cWinsock.cLog (data)

xyS = (mIn + hR * 60)
dSince = Format$(Now, "hh:mm:ss AM/PM")

With cWinsock.lvwDB.ListItems.Item(cWinsock.SEL)
    .SubItems(11) = 1
    .SubItems(12) = 0
    .SubItems(2) = cName.Text
    .SubItems(3) = dSince
    .SubItems(4) = xyS
    .SubItems(5) = dTTime - .SubItems(12)
    .SubItems(6) = " 0 Kb"
    .SubItems(7) = " 0 Kb"
    .SubItems(8) = idType.Text
    .SubItems(9) = idNum.Text
    .SubItems(10) = Mob.Text
End With


    Case 7
 End Select
End If


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
