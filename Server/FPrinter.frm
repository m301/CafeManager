VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FPrinter 
   Caption         =   "Printer"
   ClientHeight    =   0
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   1560
   Icon            =   "FPrinter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   1560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   3420
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
      Top             =   -255
      Width           =   1560
      _ExtentX        =   2752
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
            Picture         =   "FPrinter.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":061C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":07F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":09D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":0BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":0D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":0F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":1312
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":14EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":16C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":18A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrinter.frx":1A7A
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
      Icons           =   "ImageList32"
      SmallIcons      =   "ImageList16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mMain 
      Caption         =   "&Printer"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mPrinter 
         Caption         =   "P&ause Printing"
         Index           =   0
      End
      Begin VB.Menu mPrinter 
         Caption         =   "Set as De&fault"
         Index           =   1
      End
      Begin VB.Menu mPrinter 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mPrinter 
         Caption         =   "Purge Print Documents"
         Index           =   3
      End
      Begin VB.Menu mPrinter 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mPrinter 
         Caption         =   "P&roperties"
         Index           =   5
      End
      Begin VB.Menu mPrinter 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mPrinter 
         Caption         =   "&Close"
         Index           =   7
      End
   End
   Begin VB.Menu mMain 
      Caption         =   "&Document"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mDocument 
         Caption         =   "P&ause"
         Index           =   0
      End
      Begin VB.Menu mDocument 
         Caption         =   "R&esume"
         Index           =   1
      End
      Begin VB.Menu mDocument 
         Caption         =   "Re&start"
         Index           =   2
      End
      Begin VB.Menu mDocument 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mDocument 
         Caption         =   "&Cancel"
         Index           =   4
      End
      Begin VB.Menu mDocument 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mDocument 
         Caption         =   "Move &Up"
         Index           =   6
      End
      Begin VB.Menu mDocument 
         Caption         =   "Move &Down"
         Index           =   7
      End
   End
   Begin VB.Menu mMain 
      Caption         =   "&View"
      Index           =   2
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
            Caption         =   "Document Name"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mColumns 
            Caption         =   "Status"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mColumns 
            Caption         =   "Owner"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mColumns 
            Caption         =   "Progress"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu mColumns 
            Caption         =   "Size"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu mColumns 
            Caption         =   "Submitted"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu mColumns 
            Caption         =   "Port"
            Index           =   6
         End
         Begin VB.Menu mColumns 
            Caption         =   "Position"
            Index           =   7
         End
         Begin VB.Menu mColumns 
            Caption         =   "Job ID"
            Index           =   8
         End
         Begin VB.Menu mColumns 
            Caption         =   "Priority"
            Index           =   9
         End
         Begin VB.Menu mColumns 
            Caption         =   "Time"
            Index           =   10
         End
      End
   End
   Begin VB.Menu mMain 
      Caption         =   "&Help"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mHelp 
         Caption         =   "&About "
         Index           =   0
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FPrinter"
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

' Menu constants
Private Const mmPrinter = 0
Private Const mmDocument = 1
Private Const mmView = 2
Private Const mpPausePrinting = 0
Private Const mpSetAsDefault = 1
Private Const mpPurgeDocs = 3
Private Const mpProperties = 5
Private Const mpClose = 7
Private Const mdPause = 0
Private Const mdResume = 1
Private Const mdRestart = 2
Private Const mdCancel = 4
Private Const mdMoveUp = 6
Private Const mdMoveDown = 7
Private Const mvRefresh = 0
Private Const mvAutoUpdate = 1
Private Const mcDocumentName = 0
Private Const mcStatus = 1
Private Const mcOwner = 2
Private Const mcProgress = 3
Private Const mcSize = 4
Private Const mcSubmitted = 5
Private Const mcPort = 6
Private Const mcPosition = 7
Private Const mcJobId = 8
Private Const mcPriority = 9
Private Const mcTime = 10

' Print job priorities
Private Const NO_PRIORITY = 0
Private Const MAX_PRIORITY = 99
Private Const MIN_PRIORITY = 1
Private Const DEF_PRIORITY = 1

' Imagelist constants
Private Const icoFlag = 1

' Default duration between updates
Private Const defNormalInterval = 5000
Private Const defForcedUpdate = 10
Private Const tmrNormalTick = 0
Private Const tmrForcedUpdate = 1

' Member variables
Private m_DevName As String
Private m_Info As CPrinterInfo
Private m_Loaded As Boolean

' Notification interface
Implements IUpdateNotification

' ****************************************************
'  Custom Form Properties
' ****************************************************
Public Property Let DeviceName(ByVal newVal As String)
   ' Setup class that drives form
   m_DevName = newVal
   Set m_Info = New CPrinterInfo
   m_Info.DeviceName = m_DevName
   If m_Loaded Then
      Call UpdateCaption
   End If
End Property

Public Property Get DeviceName() As String
   DeviceName = m_DevName
End Property

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

Public Function Update(Optional ByVal Propogate As Boolean = False)
 
End Function

' ****************************************************
'  Form Events
' ****************************************************
Private Sub Form_Load()
   ' Set some default properties for listview
   With ListView1
      .Arrange = lvwAutoTop
      .LabelEdit = lvwManual
      .View = lvwReport
      Set .SmallIcons = ImageList16
   End With
   Call LVSetStyleEx(ListView1, FullRowSelect, True)
   Call LVSetStyleHeader(ListView1, HeaderFlat)
   ' Build listview headers and go visible
   Call RebuildList(False)
   Me.Width = Screen.Width \ 2
   Me.Height = Screen.Height \ 4
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
   ' Flag lets other properties know they can reference form/controls
   m_Loaded = True
   If cWinsock.Check3.Value = 1 Then m_Info.ControlPause
End Sub

Private Sub Form_Resize()
   ' Reposition controls
   On Error Resume Next
   ListView1.Move 0, 0, Me.ScaleWidth, _
      Me.ScaleHeight - StatusBar1.Height
   StatusBar1.Panels(1).Width = Me.ScaleWidth
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' If a print job is selected, then pop "Documents"
   ' menu. If no print job is selected, then pop the
   ' "Printer" menu with last two options invisible.
   If Button = vbRightButton Then
      If ListView1.SelectedItem Is Nothing Then
         mPrinter(mpClose - 1).Visible = False
         mPrinter(mpClose).Visible = False
         PopupMenu mMain(mmPrinter)
      Else
         PopupMenu mMain(mmDocument)
      End If
   End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
   ' This may have been internally triggered to provide
   ' semi-asyncronous updates, in which case we need to
   ' restore timer to default operation mode.
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
  
End Sub

Private Sub mDocument_Click(Index As Integer)
   Dim job As CPrinterJobInfo
   ' Get reference to higlighted job.
   Set job = GetSelectedJob()
   ' React to menu selection.
   Select Case Index
      Case mdPause
         m_Info.Jobs.ControlPause job.JobId
      Case mdResume
         m_Info.Jobs.ControlResume job.JobId
      Case mdRestart
         m_Info.Jobs.ControlRestart job.JobId
      Case mdCancel
         m_Info.Jobs.ControlCancel job.JobId
      Case mdMoveUp
         m_Info.Jobs.PositionMoveUp job.JobId
      Case mdMoveDown
         m_Info.Jobs.PositionMoveDown job.JobId
   End Select
   ' Update or rebuild, as needed.
   If Index < mdMoveUp Then
    Me.Update
   Else
      Me.Update
   End If
End Sub

Private Sub mHelp_Click(Index As Integer)
   Dim Frm As FAbout
   Set Frm = New FAbout
   Frm.Show vbModal
End Sub

Private Sub mPrinter_Click(Index As Integer)
   ' Reset visibility of hidden items.
   mPrinter(mpClose - 1).Visible = True
   mPrinter(mpClose).Visible = True
   ' Act on selection.
   Select Case Index
      Case mpPausePrinting
         ' Pause or resume depending on what
         ' state printer is already in.
         If m_Info.IsPaused Then
            m_Info.ControlResume
         Else
            m_Info.ControlPause
         End If
         m_Info.Refresh
   '     Call Me.Update
         
      Case mpSetAsDefault
         ' Set this printer to be the default.
         m_Info.IsDefault = True
         Me.Update Propogate:=True
         
      Case mpPurgeDocs
         ' Purge all documents for this printer.
         m_Info.ControlPurge
         Me.Update
         
      Case mpProperties
         If Not m_Info Is Nothing Then
            m_Info.ShowPropertiesDialog
         End If
         
      Case mpClose
         Unload Me
   End Select
End Sub

Private Sub mMain_Click(Index As Integer)
   Dim job As CPrinterJobInfo
   Dim Switch As Boolean
   ' Make sure only relevent options are enabled
   Select Case Index
      Case mmPrinter
         ' Check if default/paused
         If Not m_Info Is Nothing Then
            mPrinter(mpSetAsDefault).Checked = m_Info.IsDefault
            mPrinter(mpPausePrinting).Checked = m_Info.IsPaused
         End If
         ' These settings are dependent on admin privs
         Switch = m_Info.CanAdminister
         mPrinter(mpPausePrinting).Enabled = Switch
         mPrinter(mpPurgeDocs).Enabled = Switch
         
      Case mmDocument
         Set job = GetSelectedJob(True)
         Switch = Not (job Is Nothing)
         mDocument(mdPause).Enabled = Switch
         mDocument(mdResume).Enabled = Switch
         mDocument(mdRestart).Enabled = Switch
         mDocument(mdCancel).Enabled = Switch
         ' Job position can only be adjusted by an admin.
         mDocument(mdMoveUp).Enabled = False
         mDocument(mdMoveDown).Enabled = False
         If Switch Then
            If m_Info.Jobs.Count > 1 Then
               If m_Info.CanAdminister Then
                  mDocument(mdMoveUp).Enabled = True
                  mDocument(mdMoveDown).Enabled = True
               End If
            End If
         End If
   End Select
End Sub

Private Sub mView_Click(Index As Integer)
   Select Case Index
      Case mvRefresh
         Call Me.Update(False)
      Case mvAutoUpdate
         mView(mvAutoUpdate).Checked = Not mView(mvAutoUpdate).Checked
         Timer1(tmrNormalTick).Enabled = mView(mvAutoUpdate).Checked
   End Select
End Sub

' ****************************************************
'  Private Methods
' ****************************************************
Private Sub FillList()
   Dim Selected As String
   Dim Status As String
   Dim itm As ListItem
   Dim inf As CPrinterJobInfo
   
   Me.MousePointer = vbHourglass
   StatusBar1.Panels(1).Text = "Retrieving printer information..."
   
   ' Update PrinterInfo, caption, and UI
   m_Info.Refresh
   Call UpdateCaption
   DoEvents
      
   ' Check for completed jobs, remove
   Call RemoveDeadJobs
   ListView1.Refresh
   
   ' Add/update each queued job.
   For Each inf In m_Info.Jobs
      Set itm = GetItem(inf)
      Call UpdateSubitems(itm, inf)
      ListView1.Refresh
   Next inf
   
   StatusBar1.Panels(1).Text = m_Info.Jobs.Count & " document(s) in queue"
   Me.MousePointer = vbDefault
End Sub

Public Function FormatBytes(ByVal Size As Long) As String
   Dim sRet As String
   Const KB& = 1024
   Const MB& = KB * KB
   ' Return size of file in kilobytes.
   If Size < KB Then
      sRet = Format$(Size, "#,##0") & " bytes"
   Else
      Select Case Size \ KB
         Case Is < 10
            sRet = Format$(Size / KB, "0.00") & " KB"
         Case Is < 100
            sRet = Format$(Size / KB, "0.0") & " KB"
         Case Is < 1000
            sRet = Format$(Size / KB, "0") & " KB"
         Case Is < 10000
            sRet = Format$(Size / MB, "0.00") & " MB"
         Case Is < 100000
            sRet = Format$(Size / MB, "0.0") & " MB"
         Case Is < 1000000
            sRet = Format$(Size / MB, "0") & " MB"
         Case Is < 10000000
            sRet = Format$(Size / MB / KB, "0.00") & " GB"
      End Select
      'sRet = sRet & " (" & Format$(Size, "#,##0") & " bytes)"
   End If
   FormatBytes = sRet
End Function

Private Function GetItem(ByVal PJI As CPrinterJobInfo) As ListItem
   Dim ndx As Long
   Dim itm As ListItem
   Dim PJIx As CPrinterJobInfo
   Const errElementNotFound As Long = 35601
   
   With ListView1
      ' Try to reference existing item, add if not there
      On Error Resume Next
      Set itm = .ListItems("x" & Hex$(PJI.JobId))
      If err.Number = errElementNotFound Then
         On Error GoTo 0
         ' Base insertion index on relative Position.
         For ndx = 1 To .ListItems.Count
            Set itm = .ListItems(ndx)
            Set PJIx = m_Info.Jobs(itm.Tag)
            If PJIx.Position > PJI.Position Then Exit For
         Next ndx
         ' Add item to listview
         Set itm = .ListItems.Add(ndx, "x" & Hex$(PJI.JobId), PJI.Document, , icoFlag)
         itm.Tag = Hex$(PJI.JobId)
      End If
      Set GetItem = itm
   End With
End Function

Private Function GetSelectedJob(Optional ByVal Refresh As Boolean = False) As CPrinterJobInfo
   ' Return item that's selected and highlighted.
   With ListView1
      If Not .SelectedItem Is Nothing Then
         If .SelectedItem.Selected Then
            If Refresh Then m_Info.Jobs.Refresh
            Set GetSelectedJob = m_Info.Jobs(.SelectedItem.Tag)
         End If
      End If
   End With
End Function

Private Sub RebuildList(Optional Refill As Boolean = True)
   Dim i As Long
   Dim itm As ListItem
   
   With ListView1
      .ListItems.Clear
      With .ColumnHeaders
         .Clear
         ' Item 0, Document Name is always included.
         .Add , , "Document Name"
         For i = 1 To mColumns.UBound
            If mColumns(i).Checked Then
               .Add , , mColumns(i).Caption
            End If
         Next i
      End With
   End With
   
   ' Filler-up!
   If Not m_Info Is Nothing Then
      Call UpdateCaption
      If Refill Then Call FillList
   End If
End Sub

Private Sub RemoveDeadJobs()
   Dim itm As ListItem
   Dim inf As CPrinterJobInfo
   Const errInvalidProcedure = 5
   ' Check each job, to insure it's still running.
   On Error Resume Next
   For Each itm In ListView1.ListItems
      Set inf = m_Info.Jobs(itm.Tag)
      ' Remove if not present in jobs collections.
      If err.Number = errInvalidProcedure Then
         err.Clear
         ListView1.ListItems.Remove itm.Index
      End If
   Next itm
End Sub

Private Sub UpdateCaption()
   Dim cap As String
   Dim svr As String
   ' Check to see if caption matches device
   ' name used by CPrinterInfo class, and
   ' add note if printer is currently paused.
   If m_Info Is Nothing Then
      cap = "Printer"
   Else
      cap = m_Info.DisplayName
      If m_Info.IsPaused Then
         cap = cap & " - Paused"
      End If
   End If
   ' Only assign if different to avoid flicker.
   If Me.Caption <> cap Then
      Me.Caption = cap
   End If
End Sub

Private Sub UpdateSubitems(ByVal itm As ListItem, ByVal PJI As CPrinterJobInfo)
   Dim nSubItem As Long
   Dim OldData As String
   Dim NewData As String
   Dim i As Long
   
   ' Iterate through menu, skipping first
   ' item (Printer name), which is
   nSubItem = 0
   For i = mColumns.LBound To mColumns.UBound
      If mColumns(i).Checked And mColumns(i).Enabled Then
         Select Case i
            Case mcStatus
               NewData = PJI.StatusText
            Case mcOwner
               NewData = PJI.NotifyName
            Case mcProgress
               If PJI.PagesPrinted > 0 Then
                  NewData = "Page " & PJI.PagesPrinted & " of " & PJI.TotalPagesMax
               Else
                  NewData = PJI.TotalPages & " page(s)"
               End If
            Case mcSize
               If PJI.SizeMax > PJI.Size Then
                  NewData = FormatBytes(PJI.SizeMax - PJI.Size) & "/" & FormatBytes(PJI.SizeMax)
               Else
                  NewData = FormatBytes(PJI.Size)
               End If
            Case mcSubmitted
               NewData = CStr(PJI.Submitted)
            Case mcPort
               NewData = m_Info.PortName
            Case mcPosition
               NewData = CStr(PJI.Position)
            Case mcJobId
               NewData = CStr(PJI.JobId)
            Case mcPriority
               Select Case PJI.Priority
                  Case NO_PRIORITY
                     NewData = "NO_PRIORITY"
                  Case MAX_PRIORITY
                     NewData = "MAX_PRIORITY"
                  Case MIN_PRIORITY
                     NewData = "MIN_PRIORITY"
                  Case Else
                     NewData = CStr(PJI.Priority)
               End Select
            Case mcTime
               NewData = CStr(PJI.Time)
         End Select
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

Public Function fUnload()
Unload Me
End Function


