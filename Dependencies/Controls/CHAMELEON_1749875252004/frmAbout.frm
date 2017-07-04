VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About..."
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Caption         =   "Chameleon Button™ OCX version: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.Label lblCopy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MADE IN URUGUAY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   5175
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   15
         X2              =   5400
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   30
         X2              =   5400
         Y1              =   2775
         Y2              =   2775
      End
      Begin VB.Label lblCopy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2001-2003 by gonchuki"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   5175
      End
      Begin VB.Label lblCopy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "gonchuki@yahoo.es"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   1725
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label lblFeatures 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Label lblFeatures 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks for Using the Chameleon Button!!!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00006000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Image imgLogo 
         Height          =   1125
         Left            =   975
         Picture         =   "frmAbout.frx":0123
         Top             =   240
         Width           =   3450
      End
      Begin VB.Label lblCopyB 
         Alignment       =   2  'Center
         Caption         =   "lblCopyB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000016&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label lblHigh 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblHigh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000016&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
   Begin prjChameleon.chameleonButton cbOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmAbout.frx":365F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub cbOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos hwnd, -1, 0, 0, 0, 0, &H3

Dim i As Long

fraMain.Caption = fraMain.Caption & cbOK.Version

lblHigh(0).ForeColor = ShiftColor(GetSysColor(15), &H20)
lblCopyB(0).ForeColor = lblHigh(0).ForeColor

For i = 0 To lblFeatures.Count - 1
    If i Then Load lblHigh(i)
    lblHigh(i).Visible = True
    lblHigh(i).Move lblFeatures(i).Left + 15, lblFeatures(i).Top + 15, lblFeatures(i).Width, lblFeatures(i).Height
    lblHigh(i).Caption = lblFeatures(i).Caption
    Set lblHigh(i).Font = lblFeatures(i).Font
Next


For i = 0 To lblCopy.Count - 1
    If i Then Load lblCopyB(i)
    lblCopyB(i).Visible = True
    lblCopyB(i).Move lblCopy(i).Left + 15, lblCopy(i).Top + 15, lblCopy(i).Width, lblCopy(i).Height
    lblCopyB(i).Caption = lblCopy(i).Caption
Next

Set lblCopy(1).MouseIcon = LoadResPicture(101, 2)
Set cbOK.MouseIcon = lblCopy(1).MouseIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub

Private Sub lblCopy_Click(Index As Integer)
Select Case Index
    Case 1
        Call ShellExecute(hwnd, "Open", "mailto:gonchuki@yahoo.es?subject=Chameleon_ButtonOCX", 0, 0, 0)
End Select
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long
Dim Red As Long, Blue As Long, Green As Long

Blue = ((Color \ &H10000) Mod &H100) + Value
Green = ((Color \ &H100) Mod &H100) + Value
Red = (Color And &HFF) + Value

If Value > 0 Then
    If Red > 255 Then Red = 255
    If Green > 255 Then Green = 255
    If Blue > 255 Then Blue = 255
ElseIf Value < 0 Then
    If Red < 0 Then Red = 0
    If Green < 0 Then Green = 0
    If Blue < 0 Then Blue = 0
End If

ShiftColor = Red + 256& * Green + 65536 * Blue
End Function
