VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Winsock 
   BorderStyle     =   0  'None
   Caption         =   ":::  MAD Cafe Manager Client :::"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   720
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   -360
      Width           =   4815
   End
   Begin VB.Timer tmR 
      Interval        =   1000
      Left            =   2760
      Top             =   840
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   -120
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CommandButton bntSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   2160
      Width           =   735
   End
   Begin MSWinsockLib.Winsock sock1 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SckUdp 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtLog 
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   -120
      Width           =   9255
   End
   Begin VB.Label dLLimit 
      Caption         =   "Download Limit"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4695
   End
End
Attribute VB_Name = "Winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Function MessageBox _
        Lib "user32" Alias "MessageBoxA" _
           (ByVal hwnd As Long, _
            ByVal lpText As String, _
            ByVal lpCaption As String, _
            ByVal wType As Long) _
        As Long
Private Const txtport As String = 1996
Public USBEnabled, wDL, sDlUnlock, sDLLock, dlExceed, noDllimit, noUsbLimit As Boolean
Dim abc As Boolean
Dim wD(10) As String
Dim Tcoun As Integer
Public hR, tH, tBef, cdLLimit As Integer
Dim IP As Integer
Dim txtIp As String
Private ServerIP As String
Private ServerPort As Long
Public Cafe_Name, Cafe_Code As String
Private TimerID As Long
Public tRunning As Boolean

Private Sub bntSend_Click()
On Error GoTo t

sock1.SendData txtSend
txtLog = txtLog & "Client : " & txtSend & vbCrLf
txtSend = ""

Exit Sub
t:
MsgBox "Error : " & eRr.Description
sock1_Close
End Sub

Public Sub ConnectToServer()
On Error GoTo t
If sock1.State <> 7 Then
sock1.Close
sock1.RemoteHost = txtIp
sock1.RemotePort = txtport
sock1.Connect
End If
Exit Sub
t:
MsgBox "Error : " & eRr.Description, vbCritical
End Sub




Private Sub Form_Unload(Cancel As Integer)
If TimerID <> 0 Then KillTimer 0, TimerID
End Sub

Private Sub sock1_Close()
sock1.Close
txtLog = txtLog & "*** Disconnected" & vbCrLf

End Sub

Private Sub sock1_Connect()
txtLog = "Connected to " & sock1.RemoteHostIP & vbCrLf
wSend "cfen" & sock1.LocalHostName
End Sub
''================================================================================''
Private Sub sock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo seRr
Dim dat As String
Dim aRR() As String
Dim cHrr() As String
txtLog.Text = txtLog.Text & dat
sock1.GetData dat, vbString
cHrr = Split(dat, "|½|")
dat = cHrr(0)
wD(0) = Left$(dat, 4)
wD(0) = LCase(wD(0))
wD(1) = Right(dat, Len(dat) - 4)
''----------------------------------------------------------------------------------''
Select Case wD(0)
Case "tadd"
tH = tH + CInt(wD(1))
tTimer.Enabled = True
TDisplay
''----------------------------------------------------------------------------------''
Case "your"
sock1.SendData "hnam" & sock1.LocalHostName
''----------------------------------------------------------------------------------''
Case "lock"
sLoCk.Show

Case "uock"
Unload sLoCk
tH = 0
hR = 0
For i = 0 To 7
Bmet.cInfo(i).Caption = " --- "
Next i
Me.tTimer.Enabled = False
cTop.cTime.Caption = Format$(Now, "hh:mm AM/PM")
''----------------------------------------------------------------------------------''
Case "tset"
Unload sLoCk
aRR = Split(wD(1), " | ")
Bmet.cInfo(3) = aRR(0)
Time = aRR(0)
Bmet.cInfo(8) = aRR(1)
Bmet.cInfo(0) = aRR(2)
Bmet.cInfo(1) = aRR(3)
Bmet.cInfo(2) = aRR(4)
Bmet.cInfo(7) = aRR(5)
Bmet.cInfo(4) = aRR(6) & " Min."
hR = 0

tH = aRR(6)
tTimer.Enabled = True
TDisplay
cFet.cInternet
If CInt(aRR(6)) = 0 Then sLoCk.Show
''----------------------------------------------------------------------------------''
Case "stat"
wSend "stat" & Bmet.lblRecv2 & "|" & Bmet.lblSent2 & "|" & Bmet.cInfo(6).Caption

Case "cfen"
Cafe_Name = wD(1)
Bmet.cFen.Caption = Cafe_Name

Case "atat"
If hR > 0 Then wSend "atat" & Bmet.lblRecv2 & "|" & Bmet.lblSent2 & "|" & Bmet.cInfo(0).Caption & "|" & Bmet.cInfo(1).Caption & "|" & Bmet.cInfo(2).Caption & "|" & Bmet.cInfo(3).Caption & "|" & tH & "|" & hR & "|" & Bmet.cInfo(6).Caption & "|" & Bmet.cInfo(7).Caption & "|" & tTimer.Enabled
Case "cpay"
tH = 0
hR = 0
For i = 0 To 7
Bmet.cInfo(i).Caption = " --- "
Next i

sLoCk.Show
Case "odat"
aRR = Split(wD(1), " | ")
Select Case aRR(0)
Case "1"
Unload sLoCk
Time = aRR(1)
Bmet.cInfo(0) = aRR(2)
Bmet.cInfo(1) = aRR(3)
Bmet.cInfo(2) = aRR(4)
Bmet.cInfo(7) = aRR(5)
Bmet.cInfo(3) = aRR(7)
Bmet.cInfo(4) = aRR(6) & " Min."
hR = 0
tH = aRR(6)
tTimer.Enabled = True
TDisplay
cFet.cInternet
Case "3"
Unload sLoCk
Time = aRR(1)
Bmet.cInfo(0) = aRR(2)
Bmet.cInfo(1) = aRR(3)
Bmet.cInfo(2) = aRR(4)
Bmet.cInfo(7) = aRR(5)
Bmet.cInfo(4) = aRR(6) & " Min."
Bmet.cInfo(3) = aRR(7)
hR = aRR(8)
tH = aRR(6)
tTimer.Enabled = True
TDisplay
cFet.cInternet

Case "2"
sLoCk.Show
Time = 0
Bmet.cInfo(0) = 0
Bmet.cInfo(1) = 0
Bmet.cInfo(2) = 0
Bmet.cInfo(7) = 0
Bmet.cInfo(4) = 0 & " Min."
hR = aRR(8)
Case "0"
If hR > 0 Then wSend "atat" & Bmet.lblRecv2 & "|" & Bmet.lblSent2 & "|" & Bmet.cInfo(0).Caption & "|" & Bmet.cInfo(1).Caption & "|" & Bmet.cInfo(2).Caption & "|" & Bmet.cInfo(3).Caption & "|" & tH & "|" & hR & "|" & Bmet.cInfo(6).Caption & "|" & Bmet.cInfo(7).Caption & "|" & tTimer.Enabled
Case "4"
Unload sLoCk
End Select

Case "rrdp"
bm.iMWidth = wD(1)
bm.rRefresh
Case "iden"
If sLoCk.vIs = True Then sLoCk.Timer1.Enabled = True
Case "cmsg"
aRR = Split(wD(1), "|")
cNoti.cMessage aRR(0), aRR(2), CInt(aRR(1))
cNoti.Show
Case "lmsg"
sLoCk.Label3.Caption = "Status : Request Disapproved ! "
sLoCk.Timer2.Enabled = True

Case "eusb"
cFet.uEnable
noUsbLimit = True
Case "dusb"
cFet.uDisable
noUsbLimit = False
Case "scli"
Clipboard.SetText wD(1)
Case "gcli"
wSend "gcli" & Clipboard.GetText
Case "atit"
cFet.CaTitle
wSend "atit" & cFet.aTitle
Case "sett"

aRR = Split(wD(1), "|")
Me.USBEnabled = CInt(aRR(0))
Me.wDL = CInt(aRR(1))
Me.sDlUnlock = CInt(aRR(2))
Me.sDLLock = CInt(aRR(3))
Me.cdLLimit = CInt(aRR(4))
Call cRefSett

Case "ulda"
aRR = Split(wD(1), "|")
If aRR(0) <> "0" Then
sLoCk.usrName.Caption = aRR(0)
sLoCk.usrTot.Caption = aRR(1)
sLoCk.usrUsed.Caption = aRR(2)
sLoCk.usrLeft.Caption = CInt(aRR(1)) - CInt(aRR(2))
sLoCk.Frame3.Visible = True
sLoCk.Frame3.ZOrder 0
Command6.Enabled = True
Command5.Enabled = True
Else
sLoCk.uStatus.Caption = "Last Status : Incorrect Information !"
End If

Case "ulti"
sLoCk.Label12.Caption = "Request Status : Not enough amount in your account"
sLoCk.Command6.Enabled = True
sLoCk.Command5.Enabled = True

Case "sedl"
cdLLimit = CInt(wD(1))
Call cRefSett
Case "nodl"
noDllimit = True
Case "wopn"
cFet.uOpen wD(1)
Case "wblo"
cFet.wBlock wD(1)
Case "cblo"
cFet.cBlock
Case "pkil"
cFet.pKill
Case "dude"
cFet.unHideDrive
Case "dhde" 'hide selected drive
cFet.hDrive CLng(wD(1))
Case "oshu" 'shutdown
Call Shell("Shutdown /s")
Case "ores" 'restart
Call Shell("Shutdown /r")
Case "olog" 'loggof
Call Shell("Shutdown /l")
Case "ocsh" 'cancel shutdown
Call Shell("Shutdown /a")
Case "oten" 'enable tskmanager
Shell "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 0 /f", vbHide
Case "otdi" 'dis taskmanger
Shell "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 1 /f", vbHide
Case "oren" 'enable run
Shell "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoRun /t REG_DWORD /d 0 /f", vbHide
Case "ordi" 'disable run
Shell "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoRun /t REG_DWORD /d 1 /f", vbHide
Case "ocen" 'control enable
Shell "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoControlPanel /t REG_DWORD /d 0 /f", vbHide
Case "ocdi" 'control dis
Shell "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoControlPanel /t REG_DWORD /d 1 /f", vbHide
Case "osen" 'enable cmd
Shell "REG add HKCU\Software\Policies\Microsoft\Windows\System /v DisableCMD /t REG_DWORD /d 0 /f", vbHide
Case "osdi" 'dis cmd
Shell "REG add HKCU\Software\Policies\Microsoft\Windows\System /v DisableCMD /t REG_DWORD /d 1 /f", vbHide
Case "oshe" 'enable shutdown
Shell "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoClose /t REG_DWORD /d 0 /f", vbHide
Case "oshd" 'dis shutdown
Shell "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoClose /t REG_DWORD /d 1 /f", vbHide
Case "oldi" 'dis log
Shell "REG add HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer /v StartMenuLogOff /t REG_DWORD /d 1 /f", vbHide
Case "olen" 'en log
Shell "REG add HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer /v StartMenuLogOff /t REG_DWORD /d 0 /f", vbHide
Case "rdis" 'disable reg
cFet.EnableRegistryTools False
Case "rena" 'enable reg
cFet.EnableRegistryTools True
Case "ocre"
cFet.Create_Restore "Manual : MAD Cafe Manager", 12
Case "exit"
Main.clAll
Case "tnot"
sLoCk.uStatus.Caption = "Status : Invalid Details !"
Case "empl"
 If wD(1) = "Administrator" Then
    Unload sLoCk
    cFet.Show
 ElseIf wD(1) = "Moderator" Then
    Unload sLoCk
 ElseIf wD(1) = "Viewer" Then
    sLoCk.uStatus.Caption = "Status : Not enough rights !"
 Else
    sLoCk.uStatus.Caption = "Status : Invalid Details !"
 End If

End Select
txtLog = txtLog & wD(0) & vbCrLf
seRr:
End Sub
''================================================================================''
Public Function cRefSett()

If (Val(Bmet.lblRecv) / 1024) - 10 > Me.cdLLimit Then
cNoti.cMessage " You Have 10Mb Download Remaining Please Minimize Your Download Or Ask Administrator To increase download limit", "Download Limite", 0
End If

If Val(Bmet.lblRecv) / 1024 > Me.cdLLimit Then
dlExceed = True
If Bmet.cInfo(0).Caption = " --- " And sDlUnlock <> 0 Then dlExceed = False
If noDllimit = True Then dlExceed = False
End If

If Me.USBEnabled = 0 Then
If noUsbLimit = False Then cFet.uDisable
Else
cFet.uEnable
End If

If Me.wDL <> 0 And dlExceed = True Then wSend "dlex"  'Send Notification
If dlExceed = True And sDLLock <> 0 Then
sLoCk.Show
End If

End Function

Private Sub sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtLog = txtLog & "*** Error : " & Description & vbCrLf
sock1_Close

End Sub
Public Function wSend(data As String)
   If sock1.State = 7 Then sock1.SendData data & "|½|"
End Function

Private Sub Form_Load()
Main.loadSetting
Me.Width = 1
Me.Height = 1
noUsbLimit = True
Me.cdLLimit = 100
tBef = 5
IP = 4420
On Error GoTo errIP
Me.Visible = False
SckUdp.Bind IP
Exit Sub
errIP:
IPIncrease
Exit Sub
End Sub

Private Sub SckUdp_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errIP
    Dim strData As String
    Dim aRR() As String
    SckUdp.GetData strData
    
    
    aRR = Split(strData, "|")

    txtIp = aRR(0)
    If aRR(1) = Winsock.Cafe_Code Then ConnectToServer
    Exit Sub
errIP:
IPIncrease
End Sub

Sub IPIncrease()
    On Error GoTo errIP
    IP = IP + 1
    SckUdp.Close
    SckUdp.Bind IP
    Exit Sub
errIP:
    IPIncrease

End Sub
Private Sub TDisplay()
On Error Resume Next
Bmet.cInfo(5).Caption = hR & " Min."
Bmet.cInfo(6).Caption = tH - hR & " Min."
cTop.cTime.Caption = "Time Used: " & hR & " Min."
End Sub


Private Sub tmR_Timer()

If sLoCk.vIs = True Then
    Desk.Visible = False
    cTop.Visible = False
Else
    cTop.Visible = True
End If

    If sock1.State = 7 Then
            If sLoCk.vIs = True Then
            sLoCk.iMg.BorderColor = vbGreen
            sLoCk.iMg.FillColor = vbGreen
            sLoCk.TabStrip1.Enabled = True
            sLoCk.Frame1(1).Enabled = True
            sLoCk.Frame2.Enabled = True
            sLoCk.Frame4.Visible = False
            
            Else
            cTop.iMg.BorderColor = vbGreen
            cTop.iMg.FillColor = vbGreen
            End If
            
    Else
        
            If sLoCk.vIs = True Then
            sLoCk.iMg.BorderColor = vbRed
            sLoCk.iMg.FillColor = vbRed
            sLoCk.TabStrip1.Enabled = False
            sLoCk.Frame1(1).Enabled = False
            sLoCk.Frame2.Enabled = False
            sLoCk.Frame4.Visible = True
            Else
            cTop.iMg.BorderColor = vbRed
            cTop.iMg.FillColor = vbRed
            End If
        End If

End Sub

Private Sub tTimer_Timer()
If tRunning = False Then tStart
TDisplay
End Sub

Public Sub tIncrease()
If tTimer.Enabled = True Then
If Val(tH) - Val(hR) <= 0 Then
sLoCk.Show
wSend "sfin"
hR = 0
tH = 0
tStop
tTimer.Enabled = False
Else

hR = hR + 1
End If

If tH - hR = tBef Then
cNoti.cMessage tBef & " Min. remaining, To increase time click Yes else ignore the message ?", tBef & " Min. Remaining !", 3, "ask"
ElseIf tH - hR = 1 Then
cNoti.cMessage "60 Sec. remaining, To increase time click Yes else ignore the message ?", "60 Sec. Remaining !", 3, "ask"

End If

End If


End Sub

Public Sub tStart()
TimerID = SetTimer(0, 0, 1000, _
            AddressOf MyTimer)
tRunning = True
End Sub
Public Sub tStop()
KillTimer 0, TimerID
tRunning = False
End Sub
