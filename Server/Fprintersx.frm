VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FPrinters 
   Caption         =   "Printers"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "Fprintersx.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   3360
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   5000
      Left            =   2880
      Top             =   120
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2940
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList16 
      Left            =   4320
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8388736
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":061C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":07F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":09D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":0BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":0D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":0F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":1312
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":14EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":16C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":18A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fprintersx.frx":1A7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList16"
      SmallIcons      =   "ImageList16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImageList32 
      Left            =   1380
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   8388736
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":1C54
            Key             =   "local"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":1E2E
            Key             =   "net"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":1E8C
            Key             =   "file"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":1EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":1F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":1FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":2004
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":2062
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":20C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":211E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":217C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":21DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Fprintersx.frx":2238
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mFile 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mFile 
         Caption         =   "P&ause Printing"
         Index           =   2
      End
      Begin VB.Menu mFile 
         Caption         =   "Set as De&fault"
         Index           =   3
      End
      Begin VB.Menu mFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mFile 
         Caption         =   "P&roperties"
         Index           =   5
      End
      Begin VB.Menu mFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mFile 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu mMain 
      Caption         =   "&View"
      Index           =   1
      Begin VB.Menu mView 
         Caption         =   "&Refresh"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mView 
         Caption         =   "&AutoUpdate"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   ^A
      End
      Begin VB.Menu mView 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mView 
         Caption         =   "&Columns"
         Index           =   3
         Begin VB.Menu mColumns 
            Caption         =   "Name"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mColumns 
            Caption         =   "Documents"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mColumns 
            Caption         =   "Status"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mColumns 
            Caption         =   "Comment"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu mColumns 
            Caption         =   "Location"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu mColumns 
            Caption         =   "Model"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu mColumns 
            Caption         =   "Port"
            Index           =   6
         End
         Begin VB.Menu mColumns 
            Caption         =   "DataType"
            Index           =   7
         End
         Begin VB.Menu mColumns 
            Caption         =   "Parameters"
            Index           =   8
         End
         Begin VB.Menu mColumns 
            Caption         =   "Attributes"
            Index           =   9
         End
      End
      Begin VB.Menu mView 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mIcons 
         Caption         =   "Lar&ge Icons"
         Index           =   0
      End
      Begin VB.Menu mIcons 
         Caption         =   "S&mall Icons"
         Index           =   1
      End
      Begin VB.Menu mIcons 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mIcons 
         Caption         =   "&Details"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mMain 
      Caption         =   "&Help"
      Index           =   2
      Begin VB.Menu mHelp 
         Caption         =   "&About this demo..."
         Index           =   0
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2001 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Win32 API declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_HWNDPARENT = (-8)

' Menu constants
Private Const mmFile = 0
Private Const mmView = 1
Private Const mfOpen = 0
Private Const mfPause = 2
Private Const mfSetDefault = 3
Private Const mfProps = 5
Private Const mfExit = 7
Private Const mvRefresh = 0
Private Const mvAutoUpdate = 1
Private Const mcName = 0
Private Const mcDocuments = 1
Private Const mcStatus = 2
Private Const mcComment = 3
Private Const mcLocation = 4
Private Const mcModel = 5
Private Const mcPort = 6
Private Const mcDatatype = 7
Private Const mcParameters = 8
Private Const mcAttributes = 9

' Imagelist constants
Private Const icoLocal = 1
Private Const icoNet = 2
Private Const icoFile = 3
Private Const icoShared = 3
Private Const icoDefault = 6
Private Const icoAddNew = 13

' Default duration between updates
Private Const defNormalInterval = 5000
Private Const defForcedUpdate = 10
Private Const tmrNormalTick = 0
Private Const tmrForcedUpdate = 1

' Member variables
Private m_X As Single
Private m_Y As Single
Private m_prns As Collection

' Hooked messages
Private Const WM_SPOOLERSTATUS = &H2A
Private Const WM_SETTINGCHANGE = &H1A

' Common constants
Private Const m_NewPrn As String = "Add Printer"

' Notification interface
Implements IUpdateNotification

' ****************************************************
'  Implemented Methods
' ****************************************************
Private Sub IUpdateNotification_Rebuild()
   ' Completely rebuild data set/display.
   Call RebuildList(True)
   Call LVSetAllColWidths(ListView1, LVSCW_AUTOSIZE_USEHEADER)
End Sub

Private Sub IUpdateNotification_Update()
   ' Set timer for immediate update
   ' upon return from this call.
   Timer1(tmrForcedUpdate).Enabled = True
End Sub

' ****************************************************
'  Custom Form Methods
' ****************************************************
Public Sub Rebuild(Optional ByVal Propogate As Boolean = False)
   Dim Frm As Form
   Dim obj As IUpdateNotification
   ' Propogate across application
   For Each Frm In Forms
      Set obj = Frm
      If Frm Is Me Then
         obj.Rebuild
      ElseIf Propogate Then
         obj.Rebuild
      End If
   Next Frm
End Sub

Public Sub Update(Optional ByVal Propogate As Boolean = False)
   Dim Frm As Form
   Dim obj As IUpdateNotification
   ' Propogate across application
   For Each Frm In Forms
      Set obj = Frm
      If Frm Is Me Then
         obj.Update
      ElseIf Propogate Then
         obj.Update
      End If
   Next Frm
End Sub

' ****************************************************
'  Form Events
' ****************************************************
Private Sub Form_Load()
   ' Add a menu shortcut
   mFile(mfExit).Caption = mFile(mfExit).Caption & vbTab & "Alt-F4"
   ' Set some default properties for listview
   With ListView1
      .Arrange = lvwAutoTop
      .LabelEdit = lvwManual
      .View = lvwReport
      Set .Icons = ImageList32
      Set .SmallIcons = ImageList16
   End With
   Call LVSetStyleEx(ListView1, FullRowSelect, True)
   Call LVSetStyleHeader(ListView1, HeaderFlat)
   ' Setup a collection for CPrinterInfo classes
   Set m_prns = New Collection
   ' Build listview headers and go visible
   Call RebuildList(False)
   Me.Width = Screen.Width \ 2
   Me.Height = Screen.Height \ 3
   Me.Show
   DoEvents
   ' Fill list then adjust column widths
   Call FillList
   Call LVSetAllColWidths(ListView1, LVSCW_AUTOSIZE_USEHEADER)
   ' Setup timers
   Timer1(tmrNormalTick).Interval = defNormalInterval
   Timer1(tmrNormalTick).Enabled = True
   Timer1(tmrForcedUpdate).Interval = defForcedUpdate
   Timer1(tmrForcedUpdate).Enabled = False
   ' Hook this form's messages to watch for wm_spoolerstatus
   If Compiled Then
      Call HookWindow(GetWindowLong(Me.hWnd, GWL_HWNDPARENT), Me)
   End If
End Sub

Private Sub Form_Resize()
   ' Reposition controls
   On Error Resume Next
   ListView1.Move 0, 0, Me.ScaleWidth, _
      Me.ScaleHeight - StatusBar1.Height
   StatusBar1.Panels(1).Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim Frm As Form
   ' Make sure all spawned forms are unloaded.
   For Each Frm In Forms
      If TypeOf Frm Is FPrinter Then
         Unload Frm
      End If
   Next Frm
   ' Gotta remember to unhook!
   Call UnhookWindow(GetWindowLong(Me.hWnd, GWL_HWNDPARENT))
End Sub

Private Sub ListView1_DblClick()
   ' Nothing should happen unless the dblclick
   ' was actually *on* an item.
   Call OpenItem(ListView1.HitTest(m_X, m_Y))
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Cache X/Y so we can use them in DblClick
   m_X = X
   m_Y = Y
   ' Offer context menu, with many choices disabled
   ' if New Printer (first) is the selected item.
   If Button = vbRightButton Then
      mFile(mfExit - 1).Visible = False
      mFile(mfExit).Visible = False

      mFile(mfProps).Enabled = (ListView1.SelectedItem.Index > 1)
      PopupMenu mMain(mmFile), , , , mFile(mfOpen)
   End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
   Dim Frm As Form
   ' This may have been triggered by a WM_SPOOLERSTATUS
   ' message, in which case we need to restore timer
   ' to default operation mode.
   If Index = tmrForcedUpdate Then
      Timer1(tmrForcedUpdate).Enabled = False
   End If
   Call FillList    ' Update printer data
End Sub

' ****************************************************
'  Menu Events
' ****************************************************
Private Sub mColumns_Click(Index As Integer)
   ' Update which columns are displayed
   mColumns(Index).Checked = Not mColumns(Index).Checked
   Call Me.Rebuild(False)
End Sub

Private Sub mFile_Click(Index As Integer)
   Dim inf As CPrinterInfo
   
   ' Reset visibility of hidden items.
   mFile(mfExit - 1).Visible = True
   mFile(mfExit).Visible = True
   ' Get reference to selected printer.
   Set inf = GetSelectedPrinter()
   ' Act on selection.
   Select Case Index
      Case mfOpen
         ' Same as a double-click
         Call OpenItem(ListView1.SelectedItem)
         
      Case mfPause
         ' Toggle current state
         If Not inf Is Nothing Then
            If inf.IsPaused Then
               inf.ControlResume
            Else
               inf.ControlPause
            End If
            inf.Refresh
         End If
         Me.Update
         
      Case mfSetDefault
         If Not inf Is Nothing Then
            inf.IsDefault = True
            If inf.IsDefault Then
               ' Gotta force new icons, and
               ' alert all child forms.
               Call Me.Rebuild(True)
            End If
         End If
         
      Case mfProps
         If Not inf Is Nothing Then
            Call inf.ShowPropertiesDialog
         End If
         
      Case mfExit
         Unload Me
   End Select
End Sub

Private Sub mHelp_Click(Index As Integer)
   Dim Frm As FAbout
   Set Frm = New FAbout
   Frm.Show vbModal
End Sub

Private Sub mIcons_Click(Index As Integer)
   Dim i As Long
   ' Switch to users desired view
   For i = lvwIcon To lvwReport
      mIcons(i).Checked = (i = Index)
   Next i
   ListView1.View = Index
End Sub

Private Sub mMain_Click(Index As Integer)
   Dim inf As CPrinterInfo
   Dim itm As ListItem
   ' Make sure only relevent options are enabled.
   Select Case Index
      Case mmFile
         ' Open needs to work with Add New as well as
         ' with any installed printer.  Assume false.
         mFile(mfOpen).Enabled = False
         If (ListView1.SelectedItem Is Nothing) = False Then
            If ListView1.SelectedItem.Selected Then
               mFile(mfOpen).Enabled = True
            End If
         End If
         ' Other menu option's enabled states are based
         ' on whether there's a selected printer.
         Set inf = GetSelectedPrinter(True)
         mFile(mfProps).Enabled = (Not (inf Is Nothing))
         mFile(mfPause).Enabled = (Not (inf Is Nothing))
         mFile(mfSetDefault).Enabled = (Not (inf Is Nothing))
         ' Toggles need to be checked appropriately.
         If Not inf Is Nothing Then
            mFile(mfPause).Checked = inf.IsPaused
            mFile(mfSetDefault).Checked = inf.IsDefault
         End If
      Case mmView
   End Select
End Sub

Private Sub mView_Click(Index As Integer)
   Select Case Index
      Case mvRefresh
         Call Me.Rebuild(False)
      Case mvAutoUpdate
         mView(mvAutoUpdate).Checked = Not mView(mvAutoUpdate).Checked
         Timer1(tmrNormalTick).Enabled = mView(mvAutoUpdate).Checked
   End Select
End Sub

' ****************************************************
'  Private Methods
' ****************************************************
Private Sub CheckUninstalled()
   Dim prn As Printer
   Dim dev As String
   Dim found As Boolean
   Dim i As Long
   
   ' Remove printers no longer installed
   With ListView1.ListItems
      For i = .Count To 1 Step -1
         dev = .Item(i).Tag
         If dev <> m_NewPrn Then
            found = False
            For Each prn In Printers
               If prn.DeviceName = dev Then
                  found = True
                  Exit For
               End If
            Next prn
            If Not found Then
               On Error Resume Next
               .Remove i
               m_prns.Remove dev
            End If
         End If
      Next i
   End With
End Sub

Private Function Compiled() As Boolean
   On Error Resume Next
   Debug.Print 1 / 0
   Compiled = (Err.Number = 0)
End Function

Private Sub FillList()
   Dim prn As Printer
   Dim inf As CPrinterInfo
   Dim Status As String
   Dim itm As ListItem
   
   Me.MousePointer = vbHourglass
   StatusBar1.Panels(1).Text = "Retrieving printer information..."
   DoEvents
   
   With ListView1
      ' Make sure nothing's been uninstalled
      Call CheckUninstalled
      
      For Each prn In Printers
         ' Get reference to corresponding CPrinterInfo
         ' object, and refresh its properties.
         Set inf = GetPrinter(prn.DeviceName, True)
         ' Get reference to corresponding listitem
         Set itm = GetItem(inf)
         ' Update datafields
         Call UpdateSubitems(itm, inf)
         ' Update statusbar text
         If .SelectedItem.Tag = inf.DeviceName Then
            Status = "Status: " & inf.StatusText & _
                     ", Documents: " & inf.Jobs.Count & _
                     ", Location: " & inf.Location
         End If
         
         ' Give UI chance to breath
         DoEvents
      Next prn
   End With
   StatusBar1.Panels(1).Text = Status
   Me.MousePointer = vbDefault
End Sub

Private Function GetItem(ByVal inf As CPrinterInfo) As ListItem
   Dim foo As Long
   Dim itm As ListItem
   Const errElementNotFound As Long = 35601
   
   ' Try to reference existing item, add if not there
   On Error Resume Next
   Set itm = ListView1.ListItems(inf.DeviceName)
   If Err.Number = errElementNotFound Then
      On Error GoTo 0
      ' Determine which icon to use, assume local
      foo = icoLocal
      If inf.IsToFile Then
         foo = icoFile
      ElseIf inf.IsNetwork Then
         foo = icoNet
      End If
      If inf.IsDefault Then foo = foo + icoDefault
      If inf.IsShared Then foo = foo + icoShared
      ' Add item to listview
      Set itm = ListView1.ListItems.Add(, inf.DeviceName, inf.DisplayName, foo, foo)
      itm.Tag = inf.DeviceName
   End If
   On Error GoTo 0
   Set GetItem = itm
End Function

Private Function GetPrinter(ByVal DevName As String, Optional ByVal Refresh As Boolean = False) As CPrinterInfo
   Dim inf As CPrinterInfo
   Dim NewObj As Boolean
   
   ' Check collection for existing reference
   On Error Resume Next
      Set inf = m_prns(DevName)
   On Error GoTo 0
   
   ' Create new object, if none found
   If inf Is Nothing Then
      Set inf = New CPrinterInfo
      inf.DeviceName = DevName
      m_prns.Add inf, DevName
      NewObj = True
   End If
   
   ' Return requested object
   If Refresh = True And NewObj = False Then
      m_prns(DevName).Refresh
   End If
   Set GetPrinter = m_prns(DevName)
End Function

Private Function GetSelectedPrinter(Optional ByVal Refresh As Boolean = False) As CPrinterInfo
   Dim inf As CPrinterInfo
   With ListView1
      ' Make sure something's selected...
      If (.SelectedItem Is Nothing) = False Then
         ' ... and highlighted!
         If .SelectedItem.Selected Then
            ' Make sure we don't have Add New
            If .SelectedItem.Tag <> m_NewPrn Then
               ' Return corresponding reference.
               Set inf = GetPrinter(.SelectedItem.Tag)
               If Refresh Then inf.Refresh
               Set GetSelectedPrinter = inf
            End If
         End If
      End If
   End With
End Function

Private Sub OpenItem(ByVal itm As ListItem)
   Dim inf As CPrinterInfo
   Dim Frm As Form
   
   ' Bail if nothing passed
   If itm Is Nothing Then Exit Sub
   
   ' Either open job list for printer or
   ' start New Printer wizard.
   If itm.Tag = m_NewPrn Then
      Call Shell("rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter")
   Else
      Set inf = GetPrinter(itm.Tag)
      If inf.IsToFile Then
         MsgBox "Files are created for documents printed to this printer. " & _
            "To find your print jobs, click the Find command on the Start menu.", _
            vbInformation, inf.DeviceName
      Else
         ' Check to see if form is already loaded.
         For Each Frm In Forms
            If TypeOf Frm Is FPrinter Then
               If Frm.DeviceName = inf.DeviceName Then
                  Frm.SetFocus
                  Exit For
               End If
            End If
         Next Frm
         ' Create new form is one wasn't found.
         If Screen.ActiveForm Is Me Then
            Set Frm = New FPrinter
            Frm.DeviceName = inf.DeviceName
            Frm.Show , Me
         End If
      End If
   End If
End Sub

Private Sub RebuildList(Optional Refill As Boolean = True)
   Dim i As Long
   Dim itm As ListItem
   
   With ListView1
      ' Clear data, then make sure that
      ' "New Printer" is first
      With .ListItems
         .Clear
         Set itm = .Add(1, m_NewPrn, m_NewPrn, icoAddNew, icoAddNew)
         itm.Tag = m_NewPrn
      End With
      With .ColumnHeaders
         .Clear
         ' Item 0, Name, is always included.
         .Add , , "Name"
         For i = 1 To mColumns.UBound
            If mColumns(i).Checked Then
               .Add , , mColumns(i).Caption
            End If
         Next i
      End With
   End With
   
   ' Filler-up!
   If Refill Then Call FillList
End Sub

Private Sub ShowProperties(ByVal itm As ListItem)
   Dim inf As CPrinterInfo
   ' Bail if nothing passed
   If itm Is Nothing Then Exit Sub
   ' Nothing to do if New Printer is selected
   If itm.Tag <> m_NewPrn Then
      Set inf = GetPrinter(itm.Tag)
      inf.ShowPropertiesDialog Me.hWnd
   End If
End Sub

Private Sub UpdateSubitems(ByVal itm As ListItem, ByVal inf As CPrinterInfo)
   Dim nSubItem As Long
   Dim OldData As String
   Dim NewData As String
   Dim i As Long
   
   ' Iterate through menu, skipping first
   ' item (Printer name), which is
   nSubItem = 0
   For i = mColumns.LBound To mColumns.UBound
      If mColumns(i).Checked And mColumns(i).Enabled Then
         With inf
            Select Case i
               Case mcDocuments
                  NewData = CStr(.Jobs.Count)
               Case mcStatus
                  NewData = .StatusText
               Case mcComment
                  NewData = .Comment
               Case mcLocation
                  NewData = .Location
               Case mcModel
                  NewData = .DriverName
               Case mcPort
                  NewData = .PortName
               Case mcDatatype
                  NewData = .Datatype
               Case mcParameters
                  NewData = .Parameters
               Case mcAttributes
                  NewData = Hex$(.Attributes)
            End Select
         End With
         ' Retrieve existing data, and update if
         ' new data is different.
         nSubItem = nSubItem + 1
         OldData = itm.SubItems(nSubItem)
         If OldData <> NewData Then
            itm.SubItems(nSubItem) = NewData
         End If
      End If
   Next i
End Sub

' ****************************************************
'  Subclassing Methods
' ****************************************************
Friend Function WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
   Dim Result As Long
   ' IMPORTANT: This routine isn't hooking messages for *this* window,
   ' but rather for this window's hidden owner window!
   Select Case msg
      ' Add handlers here for each message you're interested in.
      Case WM_SPOOLERSTATUS
         ' Set timer to shortest possible duration
         ' so updates occur almost immediately after
         ' return from this message.
         ' Unfortunately, this message isn't offered
         ' by Windows 2000 or XP.  :-(
         Result = CallWindowProc(GetProp(hWnd, MHookMe.keyWndProc), hWnd, msg, wp, lp)
         Me.Update
      
      Case WM_SETTINGCHANGE
         ' MSDN: In general, when you receive this message, you
         ' should check and reload any system parameter settings
         ' that are used by your application.
         Result = CallWindowProc(GetProp(hWnd, MHookMe.keyWndProc), hWnd, msg, wp, lp)
         Me.Rebuild Propogate:=True
      
      Case Else
         ' Pass along to default window procedure.
         Result = CallWindowProc(GetProp(hWnd, MHookMe.keyWndProc), hWnd, msg, wp, lp)
   End Select
   
   ' Return desired result code to Windows.
   WindowProc = Result
End Function

