VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form rPt 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   16035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   2400
      TabIndex        =   5
      Top             =   3000
      Width           =   8775
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkColWidth 
         Caption         =   "Check Column Width"
         Height          =   195
         Left            =   6720
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkIcons 
         Caption         =   "Check Icons"
         Height          =   255
         Left            =   5280
         TabIndex        =   7
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.PictureBox picScroll 
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   13035
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2040
      Width           =   13095
      Begin VB.PictureBox picTarget 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2625
         ScaleWidth      =   3825
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   3855
      End
      Begin VB.HScrollBar hscScroll 
         Height          =   255
         LargeChange     =   15
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3000
         Width           =   4575
      End
      Begin VB.VScrollBar vscScroll 
         Height          =   2535
         LargeChange     =   15
         Left            =   4320
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   10800
      Top             =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin ComctlLib.ImageList imlImages 
      Left            =   4680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   32896
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "rPt.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "rPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'download by http://www.codefans.net
Option Explicit

'The dimensions of the DIN A4 paper size in Twips:
Const A4Height = 16840, A4Width = 11907

'To get the scroll width:
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYHSCROLL = 3
Private Const SM_CXVSCROLL = 2

'Declared Private WithEvents to get NewPage event:
Private WithEvents cTP As clsTablePrint
Attribute cTP.VB_VarHelpID = -1
Public lvData As ListView

Private Sub InitializePictureBox()
    Dim sngVSCWidth As Single, sngHSCHeight As Single
    'Set the size to the DIN A4 width:
    picTarget.Width = A4Width
    picTarget.Height = A4Height
    'Resize the scrollbars:
    sngVSCWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    sngHSCHeight = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
    hscScroll.Move 0, picScroll.ScaleHeight - sngHSCHeight, picScroll.ScaleWidth - sngVSCWidth, sngHSCHeight
    vscScroll.Move picScroll.ScaleWidth - sngVSCWidth, 0, sngVSCWidth, picScroll.ScaleHeight
    
    SetScrollBars
End Sub

Private Sub SetScrollBars()
    hscScroll.Max = (picTarget.Width - picScroll.ScaleWidth + vscScroll.Width) / 120 + 1
    vscScroll.Max = (picTarget.Height - picScroll.ScaleHeight + hscScroll.Height) / 120 + 1
End Sub


Private Sub chkColWidth_Click()
    cmdRefresh_Click
End Sub

Private Sub chkIcons_Click()
    cmdRefresh_Click
End Sub



Private Sub cmdPrint_Click()
    
    If MsgBox("The application will now print the grid on the default printer (Show a print dialog here later !).", vbInformation + vbOKCancel, "Print") = vbCancel Then Exit Sub
    
    'Simply initialize the printer:
    Printer.Print
    
    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportListView cTP, lvData, IIf((chkColWidth.Value = vbChecked), Printer.ScaleWidth - 2 * 567, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567
    
    'Class begins drawing at CurrentY !
    Printer.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable Printer
    'Done with drawing !
    
    'Say VB it should finally send it:
    Printer.EndDoc
End Sub

Private Sub cmdRefresh_Click()
    
    'Read the ListView:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportListView cTP, lvData, IIf((chkColWidth.Value = vbChecked), picTarget.ScaleWidth - 2 * 567, -1), chkIcons.Value
    
    'Here you can set RowHeightMin and HeaderRowMinHeight if the rows are too small:
'    cTP.RowHeightMin = 180
'    cTP.HeaderRowHeightMin = cTP.RowHeightMin
    
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567
    
    'Clear the box:
    picTarget.Cls
    
    'Class begins drawing at CurrentY !
    picTarget.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable picTarget
    'Done with drawing !
End Sub

Private Sub cTP_NewPage(objOutput As Object, TopMarginAlreadySet As Boolean, bCancel As Boolean, ByVal lLastPrintedRow As Long)
    
    'The class wants a new page, look what to do
    If TypeOf objOutput Is Printer Then
        Printer.NewPage
    Else 'We are printing on the PictureBox !
        objOutput.CurrentY = objOutput.ScaleHeight
        'Simply increase the height of the PicBox here
        ' (very simple, but looks bad in "real" applications)
        objOutput.Height = objOutput.Height + A4Height
        'Draw a line to show the new page:
        objOutput.Line (0, objOutput.CurrentY)-(objOutput.ScaleWidth, objOutput.CurrentY), &H808080
        
        'Set the CurrentY to the position the class should continie with drawing and...
        objOutput.CurrentY = objOutput.CurrentY + cTP.MarginTop
        '... tell it to do so:
        TopMarginAlreadySet = True
        
        'Set the ScrollBar Max properties:
        SetScrollBars
    End If
End Sub

Private Sub Form_Load()
    InitializePictureBox
    Set cTP = New clsTablePrint
    
End Sub


Private Sub Form_Resize()
picScroll.Width = Me.Width - 500
picScroll.Height = Me.Height - picScroll.Top - 700
InitializePictureBox
 cmdRefresh_Click
End Sub

Private Sub hscScroll_Change()
    picTarget.Left = -hscScroll.Value * 120
End Sub

Private Sub hscScroll_Scroll()
    hscScroll_Change
End Sub


Private Sub Timer1_Timer()
cmdRefresh_Click
End Sub

Private Sub vscScroll_Change()
    picTarget.Top = -vscScroll.Value * 120
End Sub


Private Sub vscScroll_Scroll()
    vscScroll_Change
End Sub

Public Sub Prvw(ListViewName As ListView)

Me.Show
cmdRefresh_Click
End Sub



