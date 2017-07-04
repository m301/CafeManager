Attribute VB_Name = "modAppBar"
Option Explicit
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'For intellisense
Public Enum AppBarPos
    abpLeft = 0&
    abpTop = 1&
    abpRight = 2&
    abpBottom = 3&
End Enum

'A rect(angle)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'AppBarData struct
Private Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long '  message specific
End Type

'This function makes it happen. Nothing can be done without it
Private Declare Function SHAppBarMessage Lib "shell32" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
'Used for adding the WS_EX_TOOLWINDOW style
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'We dont *have* to subclass, but we do want to do things right, dont we?
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Used to forward window messages to the next window proc in the queue
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Move the window
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Get the window dimensions
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Get desktop window
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Const ABM_NEW = &H0
Const ABM_REMOVE = &H1
Const ABM_QUERYPOS = &H2
Const ABM_SETPOS = &H3
Const ABM_GETSTATE = &H4
Const ABM_GETTASKBARPOS = &H5
Const ABM_ACTIVATE = &H6               '  lParam == TRUE/FALSE means activate/deactivate
Const ABM_GETAUTOHIDEBAR = &H7
Const ABM_SETAUTOHIDEBAR = &H8          '  this can fail at any time.  MUST check the result
Const ABM_WINDOWPOSCHANGED = &H9

Const ABN_STATECHANGE = &H0
Const ABN_POSCHANGED = &H1
Const ABN_FULLSCREENAPP = &H2
Const ABN_WINDOWARRANGE = &H3 '  lParam == TRUE means hide

Const ABS_AUTOHIDE = &H1
Const ABS_ALWAYSONTOP = &H2

Const WM_USER = &H400
Const WM_ACTIVATE = &H6
Const WM_SIZE = &H5
Const WM_MOVE = &H3

Const GWL_WNDPROC = (-4)
Const GWL_EXSTYLE = (-20)

Const WS_EX_TOOLWINDOW = &H80&

Const HWND_TOP = 0&
Const HWND_BOTTOM = 1&

Const SWP_NOSIZE = &H1&
Const SWP_NOMOVE = &H2&
Const SWP_NOZORDER = &H4

'The old windowproc
Dim lOldProc As Long
'The hWnd
Dim lhWnd As Long
'Since we need this so much, just keep a copy permanently
Dim abdAppBar As APPBARDATA

Public Sub StartAppBar(Frm As Form, Position As AppBarPos)
    'Dont want to subclass twice
    If lOldProc = 0 Then
        Dim rScreen As RECT
        Dim rFrm As RECT
        
        lhWnd = Frm.hWnd
        SetWindowLong lhWnd, GWL_EXSTYLE, GetWindowLong(lhWnd, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW
        
        GetWindowRect GetDesktopWindow, rScreen
        GetWindowRect lhWnd, rFrm
        
        rFrm.Bottom = rFrm.Bottom - rFrm.Top
        rFrm.Right = rFrm.Right - rFrm.Left
        rFrm.Top = 0
        rFrm.Left = 0
    
        'Subclass!
        lOldProc = SetWindowLong(lhWnd, GWL_WNDPROC, AddressOf AppBarProc)
        
        abdAppBar.cbSize = Len(abdAppBar)
        abdAppBar.hWnd = lhWnd
        abdAppBar.uCallbackMessage = WM_USER
        
        If SHAppBarMessage(ABM_NEW, abdAppBar) = 0 Then
            'Uh-oh, something went wrong!
            StopAppBar
            Exit Sub
        End If
        
        'Where is the taskbar?
        SHAppBarMessage ABM_GETTASKBARPOS, abdAppBar
        
        'Size our window so its in the right place
        With abdAppBar.rc
        
            If .Top > rScreen.Top Then
                'Taskbar is at the bottom
                rScreen.Bottom = .Top
            ElseIf .Bottom < rScreen.Bottom Then
                'Taskbar is at the top
                rScreen.Top = .Bottom
            ElseIf .Right < rScreen.Right Then
                'Taskbar is at the left
                rScreen.Left = .Right
            Else
                'Taskbar is at the right
                rScreen.Right = .Left
            End If
                
            abdAppBar.rc = rScreen
    
            Select Case Position
            Case AppBarPos.abpLeft
                .Right = rFrm.Right
                
            Case AppBarPos.abpTop
                .Bottom = rFrm.Bottom
                
            Case AppBarPos.abpRight
                .Left = .Right - rFrm.Right
                
            Case AppBarPos.abpBottom
                .Top = .Bottom - rFrm.Bottom
                
            End Select
        End With
        
        'Which edge are we using?
        abdAppBar.uEdge = Position
                
        'Ask the OS to find us a space to put the AppBar
        SHAppBarMessage ABM_QUERYPOS, abdAppBar
        'Tell the OS we're putting our AppBar there (OS reduces desktop space to fit)
        SHAppBarMessage ABM_SETPOS, abdAppBar
        'Move our window
        SetWindowPos lhWnd, 0, abdAppBar.rc.Left, abdAppBar.rc.Top, abdAppBar.rc.Right - abdAppBar.rc.Left, abdAppBar.rc.Bottom - abdAppBar.rc.Top, SWP_NOZORDER
        
    End If
End Sub
Public Sub StopAppBar()
    'Dont want to unsubclass a non-subclassed window
    If lOldProc Then
        'Tell the OS we're done with the AppBar
        SHAppBarMessage ABM_REMOVE, abdAppBar
        'Unsubclass
        SetWindowLong lhWnd, GWL_WNDPROC, lOldProc
        'Reset so we can do it all again
        lOldProc = 0
    End If
End Sub

Public Function AppBarProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case uMsg
    Case WM_ACTIVATE
        'Window got activated
        SHAppBarMessage ABM_ACTIVATE, abdAppBar
        
    Case WM_USER
        'Special AppBar message
        
        Select Case wParam
        Case ABN_STATECHANGE
            'Notifies an appbar that the taskbar's autohide or always-on-top state has changed—that is,
            'the user has selected or cleared the "Always on top" or "Auto hide" check box on the taskbar's property sheet.
        
        Case ABN_POSCHANGED
            'Notifies an appbar when an event has occurred that may affect the appbar's size and position.
            'Events include changes in the taskbar's size, position, and visibility state, as well as the
            'addition, removal, or resizing of another appbar on the same side of the screen.
        
            GetWindowRect lhWnd, abdAppBar.rc
            SHAppBarMessage ABM_QUERYPOS, abdAppBar
            SHAppBarMessage ABM_SETPOS, abdAppBar
            SetWindowPos lhWnd, 0, abdAppBar.rc.Left, abdAppBar.rc.Top, abdAppBar.rc.Right, abdAppBar.rc.Bottom, SWP_NOZORDER
        
        Case ABN_FULLSCREENAPP
            'Notifies an appbar when a full-screen application is opening or closing.
            'This notification is sent in the form of an application-defined message that is set by the ABM_NEW message.
            
            If CBool(lParam) Then
                'Fullscreen app is loading!
                'Pop AppBar to the back
                SetWindowPos lhWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            Else
                'Fullscreen app finished
                'Pop AppBar to the front
                SetWindowPos lhWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            End If
            
        Case ABN_WINDOWARRANGE
            'Notifies an appbar that the user has selected the Cascade,
            'Tile Horizontally, or Tile Vertically command from the taskbar's context menu.
            
        End Select
    End Select
    
    'Forward message to next windowproc
    AppBarProc = CallWindowProc(lOldProc, hWnd, uMsg, wParam, lParam)
End Function


