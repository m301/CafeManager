VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "SMain"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1
Const SettingPath As String = "d:\maddb.dll"

Private Sub Command1_Click()

Shell "explorer.exe"
End Sub
Public Sub EDecrypt(Input_File As String, Output_File As String, Password As String)
On Error Resume Next
 Dim FileName As String
  Dim FileNum As Integer
  Dim FileBytes() As Byte

  'Open File Here
  'Veriable = "FileName"
  
  FileName = Input_File
  
  'Reading File
  ReDim FileBytes(FileLen(FileName) - 1)
  FileNum = FreeFile
  Open FileName For Binary Access Read As FileNum
    Get #FileNum, , FileBytes
  Close FileNum

  'Encrypting/Decrypting Data
  RC4 FileBytes, Password 'Password

  'Save File Here
  'Veriable = "FileName"
  FileName = Output_File

  'Writing File
  FileNum = FreeFile
  Open FileName For Binary Access Write As FileNum
    Put #FileNum, , FileBytes
  Close FileNum
End Sub
Private Sub Form_Load()
frmLogin.Visible = True
chkSysTray_Click

'buTTonStyle =[3D Hover]
'buTTonStyle =[Flat Highlight]
'buTTonStyle =[Java metal]
'buTTonStyle =[KDE 2]
'buTTonStyle = Mac
'buTTonStyle = [Netscape 6]
'ButtOnStyle = [Office XP]
'buTTonStyle =[Oval Flat]
'buTTonStyle =[Simple Flat]
'ButtOnStyle = Transparent
'buTTonStyle =[Windows 16-bit]
'buTTonStyle =[Windows 32-bit]
'ButtOnStyle = [Windows XP]


cWinsock.uCurrency = "Rs. "
'Call loadSetting

End Sub
Public Sub fClear(File_Path As String)
Open File_Path For Output As #15
Print #15, "Temp File for temperory data"
Close #15
End Sub
Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKEY As String)
   Select Case sKEY
   Case "Show"
      cWinsock.Show
      cWinsock.ZOrder
    Case "Hide"
      cWinsock.Hide
   Case "Exit"
      CloseAll
Case "About"
frmAbout.Show
   Case Else
      
   End Select
    
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    cWinsock.Show
    cWinsock.ZOrder
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then
        m_frmSysTray.ShowMenu
    End If
End Sub
Public Sub chkSysTray_Click()
    
        Set m_frmSysTray = New frmSysTray
        With m_frmSysTray
            .AddMenuItem "&Hide", "Hide", True
            .AddMenuItem "&Show", "Show", True
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

Main.CloseAll

End Sub

Public Function CloseAll()

Dim Frm As Form

cWinsock.SaveAll
For Each Frm In Forms
  Unload Frm
Next Frm
Unload addTm
Unload cCheck
Unload cMessage
Unload cOpt
Unload cRDP
Unload cReciept
Unload cToken
Unload FAbout
Unload FPrinter
Unload FPrinters
Unload frmSplash
Unload frmSysTray
Unload Main
Unload sTransfer
Unload tmAdd
Unload frmLogin
Unload Main
Unload cWinsock
End Function

Public Function loadSetting()
On Error GoTo cerr
Dim lstType, newLine, aRR() As String

If Dir(SettingPath) <> "" Then

EDecrypt SettingPath, SettingPath & "x", "m1a9d9s6@MADHAXDATABASE"

Open SettingPath & "x" For Input As #1
   'MsgBox SettingPath & "x"
   
   Do Until EOF(1)
   
     Line Input #1, newLine
     
        If Left(newLine, 3) = "***" Then
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
                Case "setti"
                aRR = Split(newLine, "|@|")
                    With cWinsock
                    .aLogin.Value = aRR(0)
                    .aRenew.Value = aRR(1)
                    .Check1.Value = aRR(2)
                    .Check3.Value = aRR(3)
                    .tPassword.Value = aRR(4)
                    .cUSB.Value = aRR(5)
                    .wDL.Value = aRR(6)
                    .sDlLock.Value = aRR(7)
                    .sDlUnlock.Value = aRR(8)
                    .uDllimit.Text = aRR(9)
                    End With
                Case "cafec"
                    cWinsock.cfeCode.Caption = newLine
                Case "cafen"
                    cWinsock.cfeName.Caption = newLine
                Case "addre"
                    cWinsock.cfeAddress.Caption = newLine
                Case "phone"
                    cWinsock.cfePhone.Caption = newLine
                Case "owner"
                    cWinsock.CfeOwner.Caption = newLine
                End Select
        End If
    
   Loop
Close #1
fClear SettingPath & "x"
If Dir(SettingPath & "x") <> "" Then Kill SettingPath & "x"
End If

Exit Function
cerr:
Close #1
fClear SettingPath & "x"

If Dir(SettingPath & "x") <> "" Then Kill SettingPath & "x"
End Function
Public Function saveSetting()
Dim cDat As String

Open SettingPath & "i" For Output As #6
Print #6, "***setti"
                    With cWinsock
                    cDat = .aLogin.Value & "|@|" & .aRenew.Value & "|@|" & .Check1.Value & "|@|" & _
                    .Check3.Value & "|@|" & .tPassword.Value & "|@|" & .cUSB.Value & "|@|" & _
                    .wDL.Value & "|@|" & .sDlLock.Value & "|@|" & .sDlUnlock.Value & "|@|" & _
                    .uDllimit.Text
                    End With
Print #6, cDat
Print #6, "***cafec"
Print #6, cWinsock.cfeCode.Caption
Print #6, "***cafen"
Print #6, cWinsock.cfeName.Caption
Print #6, "***addre"
Print #6, cWinsock.cfeAddress.Caption
Print #6, "***phone"
Print #6, cWinsock.cfePhone.Caption
Print #6, "***owner"
Print #6, cWinsock.CfeOwner.Caption
Print #6, "***ENDOF"

Close #6
EDecrypt SettingPath & "i", SettingPath, "m1a9d9s6@MADHAXDATABASE"

Kill SettingPath & "i"
End Function



