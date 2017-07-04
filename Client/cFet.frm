VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form cFet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::: Admin Panel :::"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView vAr 
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer kIllByTitle 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Details"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "cFet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iAS, pTot As Integer '' for process
Dim pName(500), pName2(500) As String 'for process
'----------------Start Block app by title
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Const WM_CLOSE = &H10

'-------------------------------------------------------------

'**************************************************************************'
'                          URL Execute Declarations                                     '
'**************************************************************************'
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

'**************************************************************************'
'                         get active Window Title                           '
'**************************************************************************'
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public aTitle As String
'***********************************************************************
'CRYPT DECLARATION
Dim x1a0(9) As Long
Dim cle(17) As Long
Dim x1a2 As Long

Dim inter As Long, res As Long, ax As Long, bx As Long
Dim cx As Long, dx As Long, si As Long, tmp As Long
Dim i As Long, c As Byte
'**************************************************************************'
'                          USB Declarations                 '
'
'*************************************************************************'
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function PaintDesktop Lib "user32" (ByVal hDC As Long) As Long

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
'**************************************************************************'
'                           END USB Declaration                                       '
'**************************************************************************'


Public Function uEnable() ' Disable USB
Dim retvalue As Long, result As Long
Dim KeyID As Long, keyvalue As Long
Dim subKey As String
Dim bufSize As Long
Dim regkey As String
Dim abc As Long
Dim a1 As Long
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String

    regkey = "SYSTEM\ControlSet001\services\USBSTOR"
    retvalue = RegCreateKey(HKEY_LOCAL_MACHINE, regkey, KeyID)
    subKey = "Type"
    keyvalue = "1"
    retvalue = RegSetValueEx(KeyID, subKey, 0&, 4, keyvalue, 4)
End Function
Public Function uDisable() 'Disable USB
Dim retvalue As Long, result As Long
Dim KeyID As Long, keyvalue As Long
Dim subKey As String
Dim bufSize As Long
Dim regkey As String
Dim abc As Long
Dim a1 As Long
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String
 s = "SYSTEM\ControlSet001\services\USBSTOR"
 a = "Type"
 lRegResult = RegOpenKey(HKEY_LOCAL_MACHINE, s, hCurKey)
 lRegResult = RegDeleteValue(hCurKey, a)
 lRegResult = RegCloseKey(hCurKey)
End Function


Public Function uOpen(ByVal URL As String) As Boolean 'Opens URL
Dim res As Long
' it is mandatory that the URL is prefixed with http:// or https://
If InStr(1, URL, "http", vbTextCompare) <> 1 Then
URL = "http://" & URL
End If
res = ShellExecute(0&, "open", URL, vbNullString, vbNullString, _
vbNormalFocus)
OpenBrowser = (res > 32)
End Function

Private Sub CloseApplication(ByVal strCaption As String)
On Error GoTo Hell

'Find the Handle of The window based
' on its caption
Dim lngHandle As Long
lngHandle = FindWindow(vbNullString, strCaption)

'Verify a Window was found
If lngHandle > 0 Then
'Tell the Window to close itself
PostMessage lngHandle, WM_CLOSE, 0&, 0&
End If

Exit_For:
Exit Sub

Hell:
GoTo Exit_For
End Sub

Public Function Pack(str As String) As String ''removes spaces
Dim words As Variant
Dim X As Long
Dim temp As String

words = Split(str, " ")
For X = LBound(words) To UBound(words)
    If words(X) <> "" Then
        temp = temp & "" & words(X)
    End If
Next X
Pack = temp
End Function

Public Function cInternet() ' Creates Shortcut of FACEBOOK & Google on desktop
sCreate "http:\\madsacsoft.com\to\facebook.com", "Facebook", App.Path & "\facebook.ico"
sCreate "http:\\madsacsoft.com\to\google.com", "Google", App.Path & "\google.ico"
sCreate "http:\\madsacsoft.com\mynote\" & cFet.Pack(Bmet.cInfo(0).Caption), "Online Notepad", "%windir%\system32\notepad.exe,0 "
'Chrome
End Function

Public Function cShort(cTarget As String, sName As String) ' Can Be used To create shortcut on desktop

Set cd = CreateObject("WScript.Shell")
DesktopPath = cd.SpecialFolders("Desktop")
Set link = cd.CreateShortcut(DesktopPath & "\" & sName & " .lnk")
link.Description = "Shortcut"
link.TargetPath = cTarget
link.WindowStyle = 3
link.IconLocation = App.Path & "\" & App.EXEName & ".exe, 0"
link.Save
End Function
Public Function sCreate(cTarget As String, sName As String, iPath As String) ' Can Be used To create shortcut on desktop

Set cd = CreateObject("WScript.Shell")
DesktopPath = cd.SpecialFolders("Desktop")
Set link = cd.CreateShortcut(DesktopPath & "\" & sName & " .lnk")
link.Description = "Shortcut"
link.TargetPath = cTarget
link.WindowStyle = 3
link.IconLocation = iPath
link.Save
End Function
Public Function CaTitle() ' active window title
    Dim ActiveWindowHandle As Long
    ActiveWindowHandle = GetForegroundWindow()
    Dim Title As String * 255
    GetWindowText ActiveWindowHandle, Title, Len(Title)
    aTitle = Trim(Title)
End Function

Public Function wBlock(weBsite As String)

Open ("C:\Windows\System32\drivers\etc\hosts") For Append As #1
Print #1, "127.0.0.1 " & weBsite
Close #1
End Function

Public Function cBlock()
Open ("C:\Windows\System32\drivers\etc\hosts") For Output As #1
Close #1
End Function
Public Function pFirstLoad() ''call it on first load
    iAS = 0
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        pName2(iAS) = Process.Name
    iAS = iAS + 1
    Next
    pTot = iAS
End Function


Public Function pKill() 'kill all app except  first load
'Dim pExist As Boolean
'Dim i, X As Integer
   ' iAS = 0
    '''list process
  '  Dim Process As Object
    'For Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
      '  pName(iAS) = Process.Name
  '  iAS = iAS + 1
   ' Next
    ''end list
''start kill
'For i = 0 To iAS
   ' pExist = False
    
   ' For X = 0 To pTot
      '  If pName2(X) = pName(i) Then pExist = True
   ' Next X
    
    'If pExist = False Then Shell "taskkill /F /IM " & pName(i), vbHide

'Next i
''end kill
End Function
Public Function hDrive(a1 As Long)
Dim retvalue As Long
Dim KeyID As Long, keyvalue As Long
Dim subKey As String
Dim regkey As String


If a1 <> 0 Then
    regkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    retvalue = RegCreateKey(HKEY_CURRENT_USER, regkey, KeyID)
    subKey = "NoDrives"
    keyvalue = a1
    retvalue = RegSetValueEx(KeyID, subKey, 0&, 4, keyvalue, 4)
End If
End Function
Public Function unHideDrive()
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String
 s = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
 a = "NoDrives"
 lRegResult = RegOpenKey(HKEY_CURRENT_USER, s, hCurKey)
 lRegResult = RegDeleteValue(hCurKey, a)
 lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub EnableRegistryTools(Optional ByVal bEnable As Boolean = True)
'Dim sKey As String
'Const HKEY_CURRENT_USER = &H80000001
'sKey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
'If bEnable Then
' delete the value whose setting disables the registry
'DeleteRegistryValue HKEY_CURRENT_USER, sKey, "DisableRegistryTools"
'Else
' if the Key doesn not exist
'If CheckRegistryKey(HKEY_CURRENT_USER, sKey) = False Then
' create the key
'CreateRegistryKey HKEY_CURRENT_USER, sKey
'End If
' create and set the value to disable the Registry
'SetRegistryValue HKEY_CURRENT_USER, sKey, "DisableRegistryTools", 1
'End If
End Sub
Public Sub Create_Restore(RPName As String, RPType As Integer) 'create restore
Dim obj As Object
Set obj = GetObject("winmgmts:{impersonationLevel=impersonate}!root/default:SystemRestore")
Call obj.CreateRestorePoint(RPName, RPType, 100)
End Sub

Private Sub Command1_Click()
cCheck.Show
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Public Function cIE() ''Clear ie history cookie

Shell ("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1"), vbHide
Shell ("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2"), vbHide
Shell ("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8"), vbHide


End Function
Public Sub FileED(InputFile As String, OutputFile As String, PasswordKey As String)
    
    Dim temp As Single
    Dim Char As String * 1
    Dim XORMask As Single
    Dim temp1 As Integer
    
    Open InputFile For Binary As #1
    Open OutputFile For Binary As #2
    
    For X = 1 To Len(PasswordKey)
        temp = Asc(Mid$(PasswordKey, X, 1))
        For y = 1 To temp
            temp1 = Rnd
        Next y
        
        ' Re-seed to throw off prying eyes
        Randomize temp1
    Next X
        
    
    Counter = 0
    For z = 1 To FileLen(InputFile)
        
        'Generate random mask
        XORMask = Int(Rnd * 256)
        
        'Get the char & change it
        Get 1, , Char
        Char = Chr$((Asc(Char) Xor XORMask))
        Put 2, , Char
        
        Counter = Counter + 1
        If Counter > Len(PasswordKey) Then Counter = 1
        
        ' Pull random numbers from the hat
        For X = 1 To (Asc(Mid$(PasswordKey, Counter, 1)) * 2)
            temp = Rnd
        Next X
    Next z


    Close #1
    Close #2
    
End Sub

Public Sub Chrome()
Dim appData As String
appData = left(Environ$("AppData"), Len(Environ$("AppData")) - 8)
Dim chFile As String
chFile = appData & "\Local\Google\Chrome\User Data\Default\Preferences"
i = LoadResData(105, "CUSTOM")
Open chFile For Binary Access Write As #1
Put #1, , i
Close #1

End Sub

'**************************Cryptography


Private Sub Assemble()

x1a0(0) = ((cle(1) * 256) + cle(2)) Mod 65536
code
inter = res

x1a0(1) = x1a0(0) Xor ((cle(3) * 256) + cle(4))
code
inter = inter Xor res


x1a0(2) = x1a0(1) Xor ((cle(5) * 256) + cle(6))
code
inter = inter Xor res

x1a0(3) = x1a0(2) Xor ((cle(7) * 256) + cle(8))
code
inter = inter Xor res

x1a0(4) = x1a0(3) Xor ((cle(9) * 256) + cle(10))
code
inter = inter Xor res

x1a0(5) = x1a0(4) Xor ((cle(11) * 256) + cle(12))
code
inter = inter Xor res

x1a0(6) = x1a0(5) Xor ((cle(13) * 256) + cle(14))
code
inter = inter Xor res

x1a0(7) = x1a0(6) Xor ((cle(15) * 256) + cle(16))
code
inter = inter Xor res

i = 0

End Sub

Private Sub code()
dx = (x1a2 + i) Mod 65536
ax = x1a0(i)
cx = &H15A
bx = &H4E35

tmp = ax
ax = si
si = tmp

tmp = ax
ax = dx
dx = tmp

If (ax <> 0) Then
ax = (ax * bx) Mod 65536
End If

tmp = ax
ax = cx
cx = tmp

If (ax <> 0) Then
ax = (ax * si) Mod 65536
cx = (ax + cx) Mod 65536
End If

tmp = ax
ax = si
si = tmp
ax = (ax * bx) Mod 65536
dx = (cx + dx) Mod 65536

ax = ax + 1

x1a2 = dx
x1a0(i) = ax

res = ax Xor dx
i = i + 1

End Sub


Public Function Crypt(TextToCrypt As String) As String
Crypt = ""
si = 0
x1a2 = 0
i = 0

For fois = 1 To 16
cle(fois) = 0
Next fois

champ1 = "m1a9d9s6@MADHAX"
lngchamp1 = Len(champ1)

For fois = 1 To lngchamp1
cle(fois) = Asc(Mid(champ1, fois, 1))
Next fois

champ1 = TextToCrypt
lngchamp1 = Len(champ1)
For fois = 1 To lngchamp1
c = Asc(Mid(champ1, fois, 1))

Assemble

If inter > 65535 Then
inter = inter - 65536
End If

cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
cfd = inter Mod 256

For compte = 1 To 16

cle(compte) = cle(compte) Xor c

Next compte

c = c Xor (cfc Xor cfd)

d = (((c / 16) * 16) - (c Mod 16)) / 16
e = c Mod 16

Crypt = Crypt + Chr$(&H61 + d) ' d+&h61 give one letter range from a to p for the 4 high bits of c
Crypt = Crypt + Chr$(&H61 + e) ' e+&h61 give one letter range from a to p for the 4 low bits of c


Next fois


End Function

Public Function Decrypt(TextToDecrypt As String) As String
Decrypt = ""
si = 0
x1a2 = 0
i = 0

For fois = 1 To 16
cle(fois) = 0
Next fois

champ1 = "m1a9d9s6@MADHAX"
lngchamp1 = Len(champ1)

For fois = 1 To lngchamp1
cle(fois) = Asc(Mid(champ1, fois, 1))
Next fois

champ1 = TextToDecrypt
lngchamp1 = Len(champ1)

For fois = 1 To lngchamp1

d = Asc(Mid(champ1, fois, 1))
If (d - &H61) >= 0 Then
d = d - &H61  ' to transform the letter to the 4 high bits of c
If (d >= 0) And (d <= 15) Then
d = d * 16
End If
End If
If (fois <> lngchamp1) Then
fois = fois + 1
End If
e = Asc(Mid(champ1, fois, 1))
If (e - &H61) >= 0 Then
e = e - &H61 ' to transform the letter to the 4 low bits of c
If (e >= 0) And (e <= 15) Then
c = d + e
End If
End If

Assemble

If inter > 65535 Then
inter = inter - 65536
End If

cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
cfd = inter Mod 256

c = c Xor (cfc Xor cfd)

For compte = 1 To 16

cle(compte) = cle(compte) Xor c

Next compte

Decrypt = Decrypt + Chr$(c)

Next fois
End Function
'*********************END CRYPTOGRAPHY
Private Sub Form_Load()

End Sub
