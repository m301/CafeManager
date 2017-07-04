VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FPrinters 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   90
   ControlBox      =   0   'False
   Icon            =   "FPrinters.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList16 
      Left            =   5400
      Top             =   1680
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
            Picture         =   "FPrinters.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":061C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":07F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":09D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":0BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":0D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":0F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":1312
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":14EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":16C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":18A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":1A7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList32 
      Left            =   4560
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   8388736
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":1C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":1F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":2288
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":25A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":28BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":2BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":2EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":320A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":3524
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":383E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":3B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":3E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinters.frx":418C
            Key             =   "local"
         EndProperty
      EndProperty
   End
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
      Top             =   -165
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mMain 
      Caption         =   "&File"
      Index           =   0
      Visible         =   0   'False
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
   End
   Begin VB.Menu mMain 
      Caption         =   "&View"
      Index           =   1
      Visible         =   0   'False
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
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
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

Private m_prns As Collection

' Hooked messages
Private Const WM_SPOOLERSTATUS = &H2A
Private Const WM_SETTINGCHANGE = &H1A

' Common constants
Private Const m_NewPrn As String = "Add Printer"

' Notification interface
Implements IUpdateNotification



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Hide
Cancel = 1
End Sub

' ****************************************************
'  Implemented Methods
' ****************************************************
Private Sub IUpdateNotification_Rebuild()
   ' Completely rebuild data set/display.
   Call RebuildList(True)
   Call LVSetAllColWidths(cWinsock.lvwPrinters, LVSCW_AUTOSIZE_USEHEADER)
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

Public Function Update(Optional ByVal Propogate As Boolean = False)
  
   Call FillList    ' Update printer data

End Function

' ****************************************************
'  Form Events
' ****************************************************
Private Sub Form_Load()
Me.Width = 0
Me.Height = 0
   ' Add a menu shortcut
   ' Set some default properties for listview
   With cWinsock.lvwPrinters
      .Arrange = lvwAutoTop
      .LabelEdit = lvwManual
      .View = lvwReport


   End With
   Call LVSetStyleEx(cWinsock.lvwPrinters, FullRowSelect, True)
   Call LVSetStyleHeader(cWinsock.lvwPrinters, HeaderFlat)
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
   Call LVSetAllColWidths(cWinsock.lvwPrinters, LVSCW_AUTOSIZE_USEHEADER)
   ' Setup timers
   Timer1(tmrNormalTick).Interval = defNormalInterval
   Timer1(tmrNormalTick).Enabled = True
   Timer1(tmrForcedUpdate).Interval = defForcedUpdate
   Timer1(tmrForcedUpdate).Enabled = False
   ' Hook this form's messages to watch for wm_spoolerstatus
   If Compiled Then
      Call HookWindow(GetWindowLong(Me.hwnd, GWL_HWNDPARENT), Me)
   End If
   
End Sub

Private Sub Form_Resize()
   ' Reposition controls
   On Error Resume Next
   cWinsock.lvwPrinters.Move 0, 0, Me.ScaleWidth, _
      Me.ScaleHeight - StatusBar1.Height
   StatusBar1.Panels(1).Width = Me.ScaleWidth
   FPrinters.Left = cWinsock.lvwDB.Left + 100
FPrinters.Top = cWinsock.lvwDB.Top + 100
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
   Call UnhookWindow(GetWindowLong(Me.hwnd, GWL_HWNDPARENT))
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
   ' Get reference to selected printer.
   Set inf = GetSelectedPrinter()

   ' Act on selection.
   Select Case Index
      Case mfOpen
         ' Same as a double-click
         Call OpenItem(cWinsock.lvwPrinters.SelectedItem)
         
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
        Me.Hide
   End Select
End Sub



Private Sub mIcons_Click(Index As Integer)
   Dim i As Long
   ' Switch to users desired view
   For i = lvwIcon To lvwReport
      mIcons(i).Checked = (i = Index)
   Next i
   cWinsock.lvwPrinters.View = Index
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
         If (cWinsock.lvwPrinters.SelectedItem Is Nothing) = False Then
            If cWinsock.lvwPrinters.SelectedItem.Selected Then
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
         Call Me.pAll
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
   With cWinsock.lvwPrinters.ListItems
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
   Compiled = (err.Number = 0)
End Function

Private Sub FillList()
   Dim prn As Printer
   Dim inf As CPrinterInfo
   Dim Status As String
   Dim itm As ListItem
   
   Me.MousePointer = vbHourglass
   StatusBar1.Panels(1).Text = "Retrieving printer information..."
   DoEvents
   
   With cWinsock.lvwPrinters
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
   Set itm = cWinsock.lvwPrinters.ListItems(inf.DeviceName)
   If err.Number = errElementNotFound Then
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
      Set itm = cWinsock.lvwPrinters.ListItems.Add(, inf.DeviceName, inf.DisplayName, foo, foo)
      itm.Tag = inf.DeviceName
   End If
   On Error GoTo 0
   Set GetItem = itm
End Function

Public Function GetPrinter(ByVal DevName As String, Optional ByVal Refresh As Boolean = False) As CPrinterInfo
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
   With cWinsock.lvwPrinters
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

Public Sub OpenItem(ByVal itm As ListItem)
   Dim inf As CPrinterInfo
   Dim Frm As Form
   
   ' Bail if nothing passed
   If itm Is Nothing Then
   Exit Sub

   End If
   
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
   
   With cWinsock.lvwPrinters
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
      inf.ShowPropertiesDialog Me.hwnd
   End If
End Sub

Private Sub UpdateSubitems(ByVal itm As ListItem, ByVal inf As CPrinterInfo)
   On Error GoTo err
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
err:
End Sub

' ****************************************************
'  Subclassing Methods
' ****************************************************
Friend Function WindowProc(hwnd As Long, msg As Long, wp As Long, lp As Long) As Long
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
         Result = CallWindowProc(GetProp(hwnd, MHookMe.keyWndProc), hwnd, msg, wp, lp)
         Me.Update
      
      Case WM_SETTINGCHANGE
         ' MSDN: In general, when you receive this message, you
         ' should check and reload any system parameter settings
         ' that are used by your application.
         Result = CallWindowProc(GetProp(hwnd, MHookMe.keyWndProc), hwnd, msg, wp, lp)
         Me.Rebuild Propogate:=True
      
      Case Else
         ' Pass along to default window procedure.
         Result = CallWindowProc(GetProp(hwnd, MHookMe.keyWndProc), hwnd, msg, wp, lp)
   End Select
   
   ' Return desired result code to Windows.
   WindowProc = Result
End Function

Public Function pAll()
Dim inf As CPrinterInfo
Dim i As Integer
   For i = 1 To cWinsock.lvwPrinters.ListItems.Count
        Set inf = GetPrinter(cWinsock.lvwPrinters.ListItems.Item(i).Tag)
   If cWinsock.Check3.Value = 1 Then
   inf.ControlPause
   Else
   inf.ControlResume
    End If
    Next i
End Function
