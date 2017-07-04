VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Desk 
   BorderStyle     =   0  'None
   Caption         =   "::: Desktop - MAD Cafe Manager Client :::"
   ClientHeight    =   5790
   ClientLeft      =   -660
   ClientTop       =   3840
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Title"
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   840
      Top             =   3720
   End
   Begin SHDocVwCtl.WebBrowser wAd 
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   7335
      ExtentX         =   12938
      ExtentY         =   4683
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
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   5775
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Desk.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Desk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long
                
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
                ByVal hWnd1 As Long, _
                ByVal hWnd2 As Long, _
                ByVal lpsz1 As String, _
                ByVal lpsz2 As String) As Long
                
Private Declare Function SetParent Lib "user32" ( _
                ByVal hWndChild As Long, _
                ByVal hWndNewParent As Long) As Long
Private Type LVBKIMAGE
   uFlags As Long
   hBmp As Long
   pszImage As String
   cchImageMax As Long
   xOffsetPercent As Long
   yOffsetPercent  As Long
End Type
Private lhWnd As Long, lhWndLV As Long
Public vIs As Boolean
Dim weRr As Integer
Private Sub Command1_Click()

Me.Visible = False
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()

Desk.wAd.Navigate Main.vuRl2
Bmet.wAd.Navigate Main.vuRl3

End Sub

Private Sub Command3_Click()
cFet.pKill

End Sub

Private Sub Form_Load()
vIs = True
weRr = 0
  
    
weRr = 0
wAd.MenuBar = False
wAd.RegisterAsDropTarget = False
wAd.Silent = True
wAd.StatusBar = False
wAd.ToolBar = 0
wAd.TheaterMode = True
wAd.AddressBar = False
wAd.Navigate (Main.vuRl2)
Call drLoad


End Sub
'==================================================================
Public Function drLoad()
On Error Resume Next
    Me.Width = (Screen.Width / 4) * 3
    Me.Height = Screen.Height
    Command1.Height = Screen.Height
    wAd.Height = Screen.Height
    wAd.Width = Me.Width - Command1.Width
    Me.Left = Screen.Width / 4
    Me.Top = 0
    
wAd.Navigate (Main.vuRl2)
Timer1.Enabled = True
End Function
Public Function sBelow()
Dim ProgMan&, shellDllDefView&, sysListView&
    
    ProgMan = FindWindow("progman", vbNullString)
    shellDllDefView = FindWindowEx(ProgMan&, 0&, "shelldll_defview", vbNullString)
    sysListView = FindWindowEx(shellDllDefView&, 0&, "syslistview32", vbNullString)
    
    SetParent Me.hwnd, sysListView
End Function

Private Sub Text1_Click()
Text1.Text = wAd.Document.Title
End Sub

Private Sub Timer1_Timer()
Me.Visible = True
sBelow
Timer1.Enabled = False
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


If Dir(App.Path & "\ad_2.html") = "" Then
Open App.Path & "\ad_2.html" For Output As #19
Print #19, "<html><body alink=#fffff vlink=#ffffff><marquee><h1 color=red>Welcome To " & Winsock.Cafe_Name & _
"</h1></marquee><center> :: MAD Cafe Manager :: </center><hr width=100%><img src=ad_2.gif  width=100%>" & _
"</body></html>"
Close #19
End If

wAd.Navigate (App.Path & "\ad_2.html")
weRr = weRr + 1
End If


End Sub


