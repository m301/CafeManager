VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form sLoCk 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "::: Locked - MAD Cafe Manager Client :::"
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   11445
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   49
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   8880
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton Command6 
         Caption         =   "Login"
         Height          =   375
         Left            =   720
         TabIndex        =   37
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         MaxLength       =   2
         TabIndex        =   34
         Text            =   "1"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   33
         Text            =   "0"
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Logout"
         Height          =   375
         Left            =   720
         TabIndex        =   29
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Note : Be cafeful while logging in charges will be deducted quickly !"
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Request Status :"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Min."
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Hr."
         Height          =   255
         Left            =   1080
         TabIndex        =   35
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label usrLeft 
         Caption         =   "Label7"
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label usrUsed 
         Caption         =   "Label7"
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label usrTot 
         Caption         =   "Label11"
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Amount Left   :"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Amount Used :"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Total Amount :"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label usrName 
         Caption         =   "Username"
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "User Name    :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   5280
      TabIndex        =   40
      Top             =   2280
      Width           =   3495
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admin"
         Height          =   2415
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   3015
         Begin VB.TextBox Text6 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   43
            Text            =   "Cafe Code"
            Top             =   960
            Width           =   2535
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Login"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Status : Click Login"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Cafe Code :"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Width           =   1575
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser wAd 
      Height          =   7095
      Left            =   4200
      TabIndex        =   22
      Top             =   2640
      Width           =   7695
      ExtentX         =   13573
      ExtentY         =   12515
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   4200
      Top             =   6000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5520
      Top             =   1560
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "User Login"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Employee Login "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Token Login"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   435
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   4680
      ScaleHeight     =   1755
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "Login"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox UserPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "sadmin"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox UserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "¤"
         TabIndex        =   5
         Text            =   "AD"
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label uStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Status : Click Login"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   240
      TabIndex        =   10
      Top             =   4920
      Width           =   3495
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox cName 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Text            =   "Your Name"
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox idType 
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Text            =   "Voter ID"
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox idNum 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Text            =   "ID Number"
         Top             =   1560
         Width           =   3015
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
         TabIndex        =   13
         Text            =   "1211043210"
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox sMm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "0"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox shH 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "1"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Login :"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Hr."
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Min."
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   2400
         Width           =   375
      End
   End
   Begin VB.Label Cfe_Name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cafe Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   48
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      Height          =   5535
      Left            =   3960
      Top             =   2400
      Width           =   7455
   End
   Begin VB.Label MADCafe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAD Cafe Manager"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label TnAme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   0
      Top             =   2160
      Width           =   11655
   End
   Begin VB.Shape iMg 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   360
      Picture         =   "uk.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
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
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "sLoCk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vIs As Boolean
Private clCLount As Integer

Private Sub CancelButton_Click()
Mob.Text = "1234567890"
cName.Text = "Your Name"
idNum.Text = "ID Number"
idType.Text = "ID Type"
shH.Text = 1
sMm.Text = 0
End Sub

Private Sub Command1_Click()
Unload sLoCk
End Sub



Private Sub Command2_Click()
Winsock.Cafe_Code = "test"
End Sub

Private Sub Command3_Click()
If Winsock.sock1.State <> 7 Then
Label1.Caption = "Status : Not Connected "
Else

Select Case TabStrip1.SelectedItem.Index - 1
Case 0

Winsock.wSend "ulog" & UserName.Text & "|" & UserPass.Text & "|"

Case 1 'Employee
Winsock.wSend "empl" & UserPass.Text & "|" & UserName.Text & "|"


Case 2 'Token
Winsock.wSend "toke" & UserName.Text & "|" & UserPass.Text & "|"
End Select

End If
End Sub



Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
Frame3.Visible = False
UserName.Text = "abc"
UserPass.Text = "xyz"

End Sub



Private Sub Form_Load()
vIs = True
Main.loadSetting
Me.Show
Me.Width = Screen.Width - 10101
Me.Height = Screen.Height
Me.Left = 0
Me.Top = 0
Dim lR As Long
lR = SetTopMostWindow(Me.hwnd, True)

Shape2.Height = Frame2.Height + Picture1.Height + 500
Shape2.Width = Me.Width - Frame2.Width - 1000
Shape2.Left = Frame2.Width + 500
wAd.Height = Shape2.Height - 100
wAd.Left = Shape2.Left + 50
wAd.Width = Shape2.Width - 100
wAd.Top = Shape2.Top + 50
MADCafe.Left = Me.Width / 2 - MADCafe.Width / 2
Cfe_Name.Width = MADCafe.Width
Cfe_Name.Left = MADCafe.Left
Cfe_Name.Caption = Winsock.Cafe_Name

Frame1(1).Left = TabStrip1.Left
Frame1(1).Top = TabStrip1.Top + TabStrip1.Height - 50
Frame2.Left = TabStrip1.Left
Frame3.Top = TabStrip1.Top
Frame3.Left = TabStrip1.Left
Frame4.Top = TabStrip1.Top
Frame4.Left = TabStrip1.Left

Shape1.Width = Screen.Width
sLoCk.Label3.Caption = "Status : Idle ! "
'idType.Style = 2
idType.AddItem "Voter ID"
idType.AddItem "Telephone Bill"
idType.AddItem "Electric Bill"
idType.AddItem "Adhar Card/UID"
idType.AddItem "Identity Card"
idType.AddItem "Passport"
idType.AddItem "Credit/Debit Card"
idType.AddItem "PAN Card"
TnAme.Caption = Winsock.sock1.LocalHostName

Unload Bmet
cNoti.Hide
Bmet.Show
Bmet.Hide
cFet.Hide
cCheck.Hide
skTime.Hide
Desk.Hide
cTop.Hide
Unload cCheck

cFet.cIE
cFet.pKill
Winsock.noUsbLimit = False
Winsock.noDllimit = False
Winsock.hR = 0
Winsock.tH = 0
For i = 0 To 7
    Bmet.cInfo(i).Caption = " --- "
Next i

Clipboard.Clear

Shell "taskkill /f /im explorer.exe", vbHide
Label2.Caption = Format$(Now, "hh:mm AM/PM")

Winsock.wSend "1loc"
AntiTaskManagerController False
UserName.Text = "abc"
UserPass.Text = "xyz"

End Sub



Public Function xLoCk()
Shell "explorer.exe", vbHide
Clipboard.Clear
AntiTaskManagerController True

Unload Bmet
Bmet.Show
Bmet.Visible = False
For i = 0 To 7
Shape1.Width = Screen.Width
Bmet.cInfo(i).Caption = " --- "
Next i
sLoCk.Width = 0
sLoCk.Height = 0

Dim lR As Long
lR = SetTopMostWindow(Me.hwnd, False)

cTop.cTime.Caption = Format$(Now, "hh:mm AM/PM")
cTop.Left = Screen.Width / 2 - Me.Width / 2
cTop.Top = 0
Unload cNoti
vIs = False
End Function

Private Sub Form_Unload(Cancel As Integer)
Call xLoCk
Unload Me
Main.Timer1.Enabled = True
End Sub

Private Sub OKButton_Click()
If cName.Text <> "Your Name" And idNum.Text <> "ID Number" Then
If Winsock.sock1.State <> 7 Then
sLoCk.Label3.Caption = "Status : Failed To Send Previous Request !"
Else
Frame2.Enabled = False
Timer2.Enabled = True
sLoCk.Label3.Caption = "Status : Previous Request Sent !"
Winsock.wSend "nlog" & cName.Text & "|" & idType.Text & "|" & idNum.Text & "|" & Mob.Text & "|" & CInt(shH.Text) * 60 + CInt(sMm.Text)
End If
Else
Label3.Caption = "Status : Please enter correct information ! "
End If
End Sub

Private Sub TabStrip1_Click()
UserName.Text = "abc"
UserPass.Text = "xyz"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 48 To 57, 8
            'okay - do nothing
        Case Else
            ' 'Eat' the input
         KeyAscii = 0
End Select

End Sub

Private Sub Timer1_Timer()
Select Case clCLount
Case 0
Me.BackColor = &HC0C000
clCLount = clCLount + 1

Case 1
Me.BackColor = vbRed
clCLount = clCLount + 1
Case 2
Me.BackColor = vbBlack
clCLount = clCLount + 1
Case 3
Me.BackColor = vbGreen
clCLount = clCLount + 1
Case 4
Timer1.Enabled = False
clCLount = 0
Me.BackColor = vbWhite
End Select
End Sub





Private Sub Mob_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKey0 To vbKey9
  Case vbKeyBack, vbKeyClear, vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    KeyAscii = 0
    Beep
End Select
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

Private Sub Timer2_Timer()
Frame2.Enabled = True
Timer2.Enabled = False
End Sub




Private Sub wAd_TitleChange(ByVal Text As String)
If weRr >= 50 Then
ElseIf UCase(wAd.Document.Title) = "" Then
ElseIf UCase(wAd.Document.Title) = "MAD TIPS & TRICKS" Then
ElseIf UCase(wAd.Document.Title) = "MAD AD" Then
ElseIf UCase(wAd.Document.Title) = "AD" Then
ElseIf UCase(wAd.Document.Title) = "MADSACSOFT" Then
ElseIf UCase(wAd.Document.Title) = "MAD CAFE MANAGER" Then
Else

Open App.Path & "\ad_4.html" For Output As #19
Print #19, "<html><body><img src=ad_4.gif  width=99% height=99%></body></html>"
Close #19
wAd.Navigate (App.Path & "\ad_4.html")
weRr = weRr + 1

End If
End Sub
