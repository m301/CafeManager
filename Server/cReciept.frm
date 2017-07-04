VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cReciept 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Reciept :::"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Default         =   -1  'True
      DownPicture     =   "cReciept.frx":0000
      Height          =   375
      Left            =   2040
      Picture         =   "cReciept.frx":08CA
      TabIndex        =   8
      Top             =   0
      Width           =   2655
   End
   Begin MSComctlLib.ListView lvwR 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   450
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label rNum 
      BackStyle       =   0  'Transparent
      Caption         =   "00001"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label tName 
      BackStyle       =   0  'Transparent
      Caption         =   "MADSAC Soft"
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name  :"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Terminal name :"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Note : It is an automatically generate reciept by MAD Cafe Manager"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   5655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Reciept Number :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Date And Time :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   120
      Top             =   1200
      Width           =   9495
   End
   Begin VB.Label CafeDetail 
      BackStyle       =   0  'Transparent
      Caption         =   "Cafe Mob: Number Addrees"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   8895
   End
   Begin VB.Label cName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label CafeName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Cyber Cafe "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
   Begin VB.Shape TitleBox 
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "cReciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Command1.Visible = False
Command2.Visible = False
Me.Height = Command1.Top + Command1.Height
PrintForm
Command1.Visible = True
Command2.Visible = True
Me.Height = Command1.Top + Command1.Height + 500
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
CafeName.Caption = cWinsock.cfeName.Caption
CafeDetail.Caption = "Address : " & cWinsock.cfeAddress.Caption & " Ph: " & cWinsock.cfePhone.Caption
Label5.Caption = Now
Dim Calc_Total As Integer
Dim Calc_Icount As Integer
Calc_Icount = 0
Calc_Total = 0
cWinsock.rCount = cWinsock.rCount + 1
rNum.Caption = cWinsock.rCount
tName.Caption = cWinsock.lblsEl.Caption
cName.Caption = cWinsock.lvwDB.ListItems(cWinsock.SEL).SubItems(2)
With lvwR.ListItems
.Add , , "Computer Usage "
End With
fResize
With lvwR.ListItems(1)
   .SubItems(1) = Val(cWinsock.lvwDB.ListItems(cWinsock.SEL).SubItems(4)) / 60 & " Hr."
    Calc_Icount = Calc_Icount + Val(Val(cWinsock.lvwDB.ListItems(cWinsock.SEL).SubItems(4)) / 60)
    .SubItems(2) = "Price List"
    Calc_Total = Calc_Total + Val(cWinsock.lvwDB.ListItems(cWinsock.SEL).SubItems(13))
    .SubItems(3) = cWinsock.uCurrency & Val(cWinsock.lvwDB.ListItems(cWinsock.SEL).SubItems(13))
End With


lvwR.ListItems.Add , , " "
lvwR.ListItems(lvwR.ListItems.Count).SubItems(1) = Calc_Icount
lvwR.ListItems(lvwR.ListItems.Count).SubItems(2) = " -- "
lvwR.ListItems(lvwR.ListItems.Count).SubItems(3) = cWinsock.uCurrency & Calc_Total

lvwR.ListItems.Add , , "       Total Amount : " & cWinsock.uCurrency & Calc_Total
lvwR.ListItems(lvwR.ListItems.Count).Bold = True
End Sub

Private Sub Form_Resize()
fResize
End Sub

Private Function fResize()
Dim lwidth As Integer
lvwR.Height = (lvwR.ListItems.Count) * 230 + 250
TitleBox.Width = Me.Width - 320
lvwR.Width = TitleBox.Width
lwidth = lvwR.Width / 10
With lvwR.ColumnHeaders
.Clear
.Add , , "Item Name", lwidth * 6.48
.Add , , "Quantity", lwidth * 1
.Add , , "Unit Price", lwidth * 1.2
.Add , , "Total", lwidth * 1.3
End With
lblNote.Top = lvwR.Top + lvwR.Height + 10
Command1.Top = lblNote.Top + lblNote.Height + 10
Command2.Top = lblNote.Top + lblNote.Height + 10
Me.Height = Command1.Top + Command1.Height + 500
Shape4.Width = TitleBox.Width
Label5.Left = TitleBox.Width + TitleBox.Left - Label5.Width
Label15.Left = Label5.Left - Label15.Width
End Function

Private Sub Label4_Click()

End Sub

