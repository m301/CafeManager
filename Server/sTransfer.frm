VERSION 5.00
Begin VB.Form sTransfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::: Transfer Session :::"
   ClientHeight    =   1665
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Text            =   "Select a Terminal"
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "sTransfer.frx":0000
      Left            =   120
      List            =   "sTransfer.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Transfer Current Session Of :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   " To"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "sTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sFrom, sTo, sList, cList As Integer
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
sList = Combo1.ListIndex + 1
End Sub



Private Sub Combo2_Click()
cList = Combo2.ListIndex + 1
End Sub

Private Sub Form_Load()
Dim i As Integer
Combo2.Text = " Select a Terminal "

For i = 1 To cWinsock.SocketCounter
Combo1.AddItem cWinsock.lvwDB.ListItems(i).Text
If i = cWinsock.SEL Then
Combo1.Text = cWinsock.lvwDB.ListItems(i).Text
sList = i
End If
Next i

For i = 1 To cWinsock.SocketCounter
Combo2.AddItem cWinsock.lvwDB.ListItems(i).Text
Next i
End Sub

Private Sub OKButton_Click()
Dim i As Integer
If Combo2.Text = " Select a Terminal " Then
MsgBox "Please Select A Terminal to Transfer Session ! ", vbOKOnly + vbInformation, "Select A Terminal"
Else


With cWinsock.lvwDB.ListItems
    
    .Item(cList).SubItems(11) = 3
    For i = 2 To 12
    .Item(cList).SubItems(i) = .Item(sList).SubItems(i)
    Next i
    
    cWinsock.wSend "odat" & .Item(cList).SubItems(11) & " | " & Format$(Now, "hh:mm:ss AM/PM") & " | " & .Item(cList).SubItems(2) & " | " & _
    .Item(cList).SubItems(8) & " | " & cWinsock.lvwDB.ListItems.Item(cList).SubItems(9) & " | " & cWinsock.lvwDB.ListItems.Item(cList).SubItems(10) & " | " & cWinsock.lvwDB.ListItems.Item(cList).SubItems(4) & " | " & .Item(cList).SubItems(3) & " | " & .Item(cList).SubItems(12), cWinsock.GETSOCK(CInt(sList))

End With
        cWinsock.wSend "cpay", cWinsock.GETSOCK(CInt(sList))
         cWinsock.lsub14 Val(sList), 0
         cWinsock.lClear Val(sList)

Unload Me

End If
End Sub
