VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "::: MAD Cafe Manager Client :::"
   ClientHeight    =   2715
   ClientLeft      =   9690
   ClientTop       =   2655
   ClientWidth     =   2880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer clock 
      Interval        =   60000
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   840
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":E859
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1
Public vuRl2, vuRl3, vuRl4, vUR As String
Private vCount As Integer

Private Sub Form_Load()

Me.Width = 1
Me.Height = 1

chkSysTray_Click
cFet.Show
cFet.Hide
sLoCk.Show ''' HERE I AM
RefURl

Winsock.Visible = True
Winsock.Visible = False
Desk.Show

Bmet.Visible = True
Bmet.Visible = False
bm.Show
bm.Visible = False
For i = 0 To 7
Bmet.cInfo(i).Caption = " --- "
Next i

cFet.pFirstLoad
cFet.cInternet

cFet.vAr.ListItems.Add , , "TaskKill By Title"


End Sub
Public Sub clAll()
Winsock.tmR.Enabled = False
Unload Winsock
Unload Main
Unload sLoCk
Unload Desk
Unload Bmet
Unload frmSysTray
Unload bm
Unload frmSplash
Unload frmAbout
Unload cNoti
Unload skTime
Unload cFet
Unload cCheck
Unload cTop
Unload bWork
End Sub


Public Sub chkSysTray_Click()
    
        Set m_frmSysTray = New frmSysTray
        With m_frmSysTray
            .AddMenuItem "&Lock", "lock", True
            .AddMenuItem "-"
            .AddMenuItem "&About...", "About"
            .AddMenuItem "-"
            .AddMenuItem "&Exit", "Exit"
            .ToolTip = "MAD Cafe Manager"
        End With
         m_frmSysTray.IconHandle = Me.Icon
   
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload m_frmSysTray
    Set m_frmSysTray = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
clAll
End Sub

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
   Select Case sKey
   Case "open"
      Me.Show
      Me.ZOrder
   Case "lock"
    Dim yon As Integer
    yon = MsgBox("Are you sure ? ", vbYesNo + 32, "Logout ")
    If yon = 6 Then sLoCk.Show
   Case "Exit"
      Unload Me
Case "About"
frmAbout.Show
   Case Else
      
   End Select
    
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then
        m_frmSysTray.ShowMenu
    End If
End Sub

Public Function RefURl()
vUR = "http://madsacsoft.com/cafe/"

vuRl2 = vUR & "ad.php?id=" & cFet.Crypt(Winsock.Cafe_Code & "|^MAD^|full|^MAD^|" & Bmet.usrType.Text & "|^MAD^|" & Bmet.cInfo(0) & "|^MAD^|")
cTop.Text1.Text = vuRl2
vuRl3 = vUR & "ad.php?id=" & cFet.Crypt(Winsock.Cafe_Code & "|^MAD^|noti|^MAD^|" & Bmet.usrType.Text & "|^MAD^|" & Bmet.cInfo(0) & "|^MAD^|")
vuRl4 = vUR & "ad.php?id=" & cFet.Crypt(Winsock.Cafe_Code & "|^MAD^|lock|^MAD^|")


If Desk.vIs = True Then
Desk.wAd.Stop
Desk.wAd.Navigate Main.vuRl2
ElseIf sLoCk.vIs = True Then
sLoCk.wAd.Stop
sLoCk.wAd.Navigate Main.vuRl4
End If
Bmet.wAd.Stop
Bmet.wAd.Navigate Main.vuRl3


End Function
Private Sub clock_Timer()
If vCount >= 5 Then
RefURl
vCount = 0
Else
vCount = vCount + 1
End If

If sLoCk.vIs = True Then
Dim lR As Long
lR = SetTopMostWindow(cTop.hwnd, True)
sLoCk.Label2.Caption = Format$(Now, "hh:mm AM/PM")
End If

If Winsock.tTimer.Enabled = False Then cTop.cTime.Caption = Format$(Now, "hh:mm AM/PM")

Winsock.wSend "sett"
Call Winsock.cRefSett
Bmet.tNow.Caption = Format$(Now, "hh:mm AM/PM")

End Sub


Private Sub Timer1_Timer()
Unload Desk
Desk.Show
cTop.Show
Dim lR As Long
lR = SetTopMostWindow(cTop.hwnd, True)
Desk.Timer1.Enabled = True
Timer1.Enabled = False
End Sub

Public Function loadSetting()
Dim lstType, newLine, aRR() As String
On Error GoTo cErr
If Dir("C:\WINDOWS\system32\MADWIN.dll") <> "" Then
Open "C:\WINDOWS\system32\MADWIN.dll" For Input As #1
   
   Do Until EOF(1)
   
     Line Input #1, newLine
     
        If left(newLine, 3) = "***" Then
                lstType = Right(newLine, 5)
                
                If lstType = "ENDOF" Then
                Close #1
                Exit Function
                End If
                 uCount = 0 'Line Count as 0
        ElseIf newLine = "" Or newLine = " " Then
        
        Else
        
            uCount = uCount + 1 'Increase Line Count
        
            Select Case lstType
                                   
                Case "cafec"
                    Winsock.Cafe_Code = newLine
                Case "cafen"
                    Winsock.Cafe_Name = newLine
                End Select
        End If
    
   Loop
Close #1
End If
cErr:
End Function
Public Function saveSetting()

Open "C:\WINDOWS\system32\MADWIN.dll" For Output As #1
Print #1, "***cafec"
Print #1, Winsock.Cafe_Code
Print #1, "***cafen"
Print #1, Winsock.Cafe_Name
Print #1, "***ENDOF"
Close #1

End Function


