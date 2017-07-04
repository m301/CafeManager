Attribute VB_Name = "atBottom"
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long
                
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
                ByVal hWnd1 As Long, _
                ByVal hWnd2 As Long, _
                ByVal lpsz1 As String, _
                ByVal lpsz2 As String) As Long
                
Public Declare Function SetParent Lib "user32" ( _
                ByVal hWndChild As Long, _
                ByVal hWndNewParent As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                          (ByVal lpPrevWndFunc As Long, _
                           ByVal hwnd As Long, _
                           ByVal msg As Long, _
                           ByVal wParam As Long, _
                            lParam As WINDOWPOS) As Long
                            
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                          (ByVal hwnd As Long, _
                           ByVal nIndex As Long, _
                           ByVal dwNewLong As Long) As Long

Public Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

Public pOldWindPoc As Long ' A pointer to the old window procedure

Public Const GWL_WNDPROC& = (-4)

Public Const WM_WINDOWPOSCHANGING As Long = &H46
Public Const SWP_NOMOVE As Long = &H2

' Our new window procedure
Public Function WndProc(ByVal hwnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       lParam As WINDOWPOS) As Long
   
    If uMsg = WM_WINDOWPOSCHANGING Then
        lParam.flags = lParam.flags Or SWP_NOMOVE 'restrict form movement
    End If
    
    ' Pass the message to original WinProc
    WndProc = CallWindowProc(pOldWindPoc, hwnd, uMsg, wParam, lParam)
     
End Function

Public Sub FormAlwaysAtBottom(hwnd As Long)
    ' Original code by Bushmobile
    ' [url]http://www.vbforums.com/showpost.php?p=2420562&postcount=10[/url]
    Dim ProgMan&, shellDllDefView&, sysListView&
    
    ProgMan = FindWindow("progman", vbNullString)
    shellDllDefView = FindWindowEx(ProgMan&, 0&, "shelldll_defview", vbNullString)
    sysListView = FindWindowEx(shellDllDefView&, 0&, "syslistview32", vbNullString)
    
    SetParent hwnd, sysListView
End Sub

