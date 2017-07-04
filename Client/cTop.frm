VERSION 5.00
Begin VB.Form cTop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   345
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape iMg 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Shape           =   3  'Circle
      Top             =   50
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   345
      Left            =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label cTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time Used"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   80
      Width           =   1815
   End
End
Attribute VB_Name = "cTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlngX As Long
Private mlngY As Long
Private Sub Command1_Click()
Main.RefURl
End Sub

Private Sub cTime_Click()
sBmet
End Sub

Private Sub cTime_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
 If Button = vbLeftButton Then
        mlngX = X
        mlngY = y
    End If
End Sub

Private Sub cTime_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.left + X - mlngX
        lngTop = Me.top + y - mlngY
        If (lngLeft >= 0 And lngTop >= 0) And (lngLeft < Screen.Width - Me.Width And lngTop < Screen.Height - Me.Height) Then Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Form_Click()
sBmet
End Sub
Private Sub Form_Load()
Dim lR As Long
lR = SetTopMostWindow(Me.hwnd, True)
Me.left = Screen.Width / 2 - Me.Width / 2
Me.top = 0
If Winsock.tTimer.Enabled = False Then cTop.cTime.Caption = Format$(Now, "hh:mm AM/PM")
End Sub

Private Sub sBmet()
Bmet.Show
Bmet.Timer3.Enabled = True
End Sub

Private Sub Text1_Change()
Main.RefURl
End Sub
