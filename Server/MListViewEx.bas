Attribute VB_Name = "MListView"
' ****************************************************************
'  Copyright ©1997-2001, Karl E. Peterson
'  http://www.mvps.org/vb
' ****************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' ****************************************************************
'  Some routines originally written by Todd Wolff <twolff@source.com>
'  Sent to Karl Peterson on 9/4/97, and heavily modified since.
'  Full demo available as LVStyles.zip at http://www.mvps.org/vb/
' *******************************************************************
Option Explicit

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type POINT
   X As Long
   Y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETITEM As Long = LVM_FIRST + 5
Private Const LVM_FINDITEM As Long = LVM_FIRST + 13
Private Const LVM_ENSUREVISIBLE = LVM_FIRST + 19
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Private Const LVM_GETTOPINDEX = LVM_FIRST + 39
Private Const LVM_SETITEMSTATE As Long = LVM_FIRST + 43
Private Const LVM_GETITEMSTATE As Long = LVM_FIRST + 44
Private Const LVM_GETITEMTEXT As Long = LVM_FIRST + 45
Private Const LVM_SORTITEMS As Long = LVM_FIRST + 48
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 55
Private Const LVM_SETCOLUMNORDERARRAY = LVM_FIRST + 58
Private Const LVM_GETCOLUMNORDERARRAY = LVM_FIRST + 59

Private Const LVS_EX_GRIDLINES As Long = &H1
Private Const LVS_EX_SUBITEMIMAGES As Long = &H2
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_TRACKSELECT As Long = &H8
Private Const LVS_EX_HEADERDRAGDROP As Long = &H10
Private Const LVS_EX_FULLROWSELECT As Long = &H20

Private Const LVFI_PARAM As Long = 1

Private Const LVIF_TEXT As Long = 1
Private Const LVIF_IMAGE As Long = 2
Private Const LVIF_PARAM As Long = 4
Private Const LVIF_STATE As Long = 8
Private Const LVIF_INDENT As Long = &H10
Private Const LVIF_NORECOMPUTE As Long = &H800
Private Const LVIS_STATEIMAGEMASK As Long = &HF000&

Private Const GWL_STYLE = (-16)
Private Const LVM_GETHEADER = (LVM_FIRST + 31)
Private Const HDS_BUTTONS = &H2

Private Type LV_ITEM
   Mask As Long
   Index As Long
   SubItem As Long
   State As Long
   StateMask As Long
   Text As String
   TextMax As Long
   Icon As Long
   Param As Long
   Indent As Long
End Type

Private Type LV_FINDINFO
   Flags As Long
   pSz As String
   lParam As Long
   pt As POINT
   vkDirection As Long
End Type


'--- Array used to speed custom sorts ---'
Private m_lvSortData() As LV_ITEM
Private m_lvSortColl As Collection
Private m_lvSortColumn As Long
Private m_lvHWnd As Long
Private m_lvSortType As LVItemTypes

'--- ListView Set Column Width Messages ---'
Public Enum LVSCW_Styles
   LVSCW_AUTOSIZE = -1
   LVSCW_AUTOSIZE_USEHEADER = -2
End Enum

'LVS_EX_CHECKBOXES: Enables items in a list view control to be displayed
'   as check boxes. This style uses item state images to produce the check
'   box effect.
'LVS_EX_FULLROWSELECT: Specifies that when an item is selected, the item
'   and all its subitems are highlighted. This style is available only in
'   conjunction with the LVS_REPORT style.
'LVS_EX_GRIDLINES: Displays gridlines around items and subitems. This style
'   is available only in conjunction with the LVS_REPORT style.
'LVS_EX_HEADERDRAGDROP: Enables drag-and-drop reordering of columns in a
'   list view control. This style is only available to list view controls
'   that use the LVS_REPORT style.
'LVS_EX_SUBITEMIMAGES: Allows images to be displayed for subitems. This
'   style is available only in conjunction with the LVS_REPORT style.
Public Enum LVStylesEx
   CheckBoxes = LVS_EX_CHECKBOXES
   FullRowSelect = LVS_EX_FULLROWSELECT
   GridLines = LVS_EX_GRIDLINES
   HeaderDragDrop = LVS_EX_HEADERDRAGDROP
   SubItemImages = LVS_EX_SUBITEMIMAGES
   TrackSelect = LVS_EX_TRACKSELECT
End Enum

Public Enum LVHeaderStyles
   HeaderFlat = 0
   Header3D = 1
End Enum

'--- Sorting Variables ---'
Public Enum LVItemTypes
   lvDate = 0
   lvNumber = 1
   lvBinary = 2
   lvAlphabetic = 3
End Enum
Public Enum LVSortTypes
   lvAscending = 0
   lvDescending = 1
End Enum

Public BuildLookup As Long
Public PerformSort As Long

Public Function LVSetStyleHeader(lv As ListView, ByVal NewStyle As LVHeaderStyles)
   Dim hHeader As Long
   Dim nStyle As Long
   
   ' get the header handle and current style
   hHeader = SendMessage(lv.hWnd, LVM_GETHEADER, 0, ByVal 0&)
   nStyle = GetWindowLong(hHeader, GWL_STYLE)
   
   ' set as requested
   Select Case NewStyle
      Case HeaderFlat  ' prevents clicks on header
         nStyle = nStyle And (Not HDS_BUTTONS)
         Call SetWindowLong(hHeader, GWL_STYLE, nStyle)
      Case Header3D    ' allows clicks to sort (or whatever)
         nStyle = nStyle Or HDS_BUTTONS
         Call SetWindowLong(hHeader, GWL_STYLE, nStyle)
   End Select
End Function

Public Function LVSetStyleEx(lv As ListView, ByVal NewStyle As LVStylesEx, ByVal NewVal As Boolean) As Boolean
   Dim nStyle As Long
   
   ' get the current ListView style
   nStyle = SendMessage(lv.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0&)
   
   If NewVal Then
      ' set the extended style bit
      nStyle = nStyle Or NewStyle
   Else
      ' remove the extended style bit
      nStyle = nStyle Xor NewStyle
   End If
   
   ' set the new ListView style
   LVSetStyleEx = CBool(SendMessage(lv.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal nStyle))
End Function

Public Function LVGetColOrder(lv As ListView) As Variant
   Dim cols() As Long
   Dim nRet As Long
   '#define ListView_GetColumnOrderArray(hWnd, iCount, pi) \
   '        (BOOL)SNDMSG((hWnd), LVM_GETCOLUMNORDERARRAY, (WPARAM)iCount, (LPARAM)(LPINT)pi)
   With lv
      ReDim cols(0 To .ColumnHeaders.Count - 1) As Long
      nRet = SendMessage(.hWnd, LVM_GETCOLUMNORDERARRAY, .ColumnHeaders.Count, cols(0))
      If nRet Then
         LVGetColOrder = cols
      End If
   End With
End Function

Public Function LVSetColOrder(lv As ListView, cols() As Long) As Boolean
   Dim nRet As Long
   Dim rClient As RECT
   '#define ListView_SetColumnOrderArray(hWnd, iCount, pi) \
   '        (BOOL)SNDMSG((hWnd), LVM_SETCOLUMNORDERARRAY, (WPARAM)iCount, (LPARAM)(LPINT)pi)
   With lv
      If (UBound(cols) + 1) = .ColumnHeaders.Count Then
         nRet = SendMessage(.hWnd, LVM_SETCOLUMNORDERARRAY, .ColumnHeaders.Count, cols(0))
         LVSetColOrder = CBool(nRet)
         Call GetClientRect(.hWnd, rClient)
         Call InvalidateRect(.hWnd, rClient, True)
      End If
   End With
End Function

Public Sub LVSetColWidth(lv As ListView, ByVal ColumnIndex As Long, ByVal Style As LVSCW_Styles)
   '------------------------------------------------------------------------------
   '--- If you include the header in the sizing then the last column will
   '--- automatically size to fill the remaining listview width.
   '------------------------------------------------------------------------------
   With lv
      ' verify that the listview is in report view and that the column exists
      If .View = lvwReport Then
         If ColumnIndex >= 1 And ColumnIndex <= .ColumnHeaders.Count Then
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal Style)
         End If
      End If
   End With
End Sub

Public Sub LVSetAllColWidths(lv As ListView, ByVal Style As LVSCW_Styles)
   Dim ColumnIndex As Long
   '--- loop through all of the columns in the listview and size each
   With lv
      For ColumnIndex = 1 To .ColumnHeaders.Count
         LVSetColWidth lv, ColumnIndex, Style
      Next ColumnIndex
   End With
End Sub

Public Function LVItemChecked(lv As ListView, ByVal Index As Long) As Boolean
   Dim nRet As Long
   Const MaskBit As Long = &H1000   '(2 ^ 12)
   
   ' get current statemask bits
   nRet = SendMessage(lv.hWnd, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_STATEIMAGEMASK)
   
   ' return what the Checked bit is set to
   LVItemChecked = (((nRet \ MaskBit) - 1) <> 0)
End Function

Public Function LVSetItemCheck(lv As ListView, ByVal Index As Long, ByVal Value As Boolean) As Boolean
   Dim lvi As LV_ITEM

   ' index must be adjusted to zero-based
   Index = Index - 1
   
   ' fill ListView Info structure to determine values to return
   lvi.Index = Index
   lvi.Mask = LVIF_STATE
   lvi.StateMask = LVIS_STATEIMAGEMASK

   ' retrieve current settings
   Call SendMessage(lv.hWnd, LVM_GETITEM, 0&, lvi)
   
   ' set appropriate mask bit on or off
   If Value Then
      lvi.State = (lvi.State And (Not LVIS_STATEIMAGEMASK)) Or &H2000
   Else
      lvi.State = (lvi.State And (Not LVIS_STATEIMAGEMASK)) Or &H1000
   End If
   
   ' send message to apply the new value
   LVSetItemCheck = SendMessage(lv.hWnd, LVM_SETITEMSTATE, Index, lvi)
End Function

Public Function LVGetFirstVisible(lv As ListView) As Long
   LVGetFirstVisible = SendMessage(lv.hWnd, LVM_GETTOPINDEX, 0&, ByVal 0&)
End Function

Public Function LVEnsureVisible(lv As ListView, ByVal Index As Long) As Boolean
   LVEnsureVisible = SendMessage(lv.hWnd, LVM_ENSUREVISIBLE, Index, ByVal 0&)
End Function

' *********************************************************
'  Knowledge Base-based Sorting Routines
' *********************************************************
Public Function LVSortK(lv As ListView, ByVal Index As Long, ByVal ItemType As LVItemTypes, ByVal SortOrder As LVSortTypes) As Boolean
   Dim tmr As New CStopWatch
   
   ' turn off the default sorting of the control
   With lv
      .Sorted = False
      .SortKey = Index
      .SortOrder = SortOrder
   End With

   ' store some values used during the sort
   m_lvSortColumn = Index
   m_lvSortType = ItemType
   m_lvHWnd = lv.hWnd
   BuildLookup = 0
   
   ' start sorting to type-specific callback routines
   tmr.Reset
   Select Case ItemType
      Case lvDate
         Call SendMessageLong(lv.hWnd, LVM_SORTITEMS, SortOrder, AddressOf LVCompareDates)
      Case lvNumber
         Call SendMessageLong(lv.hWnd, LVM_SORTITEMS, SortOrder, AddressOf LVCompareNumbers)
   End Select
   PerformSort = tmr.Elapsed
End Function

Private Function LVCompareDates(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As Long) As Long
   Static dat1 As Date
   Static dat2 As Date
   
   ' lookup text in listview based on index, and convert to date
   On Error Resume Next
   dat1 = CDate(LVGetItemText(lParam1, m_lvHWnd))
   dat2 = CDate(LVGetItemText(lParam2, m_lvHWnd))
   On Error GoTo 0

   '--- this sorts ascending
   LVCompareDates = Sgn(dat1 - dat2)
   
   '--- this sorts descending
   If SortOrder = lvDescending Then
      LVCompareDates = -LVCompareDates
   End If
End Function

Private Function LVCompareNumbers(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As Long) As Long
   Static dat1 As Double
   Static dat2 As Double
   
   ' lookup text in listview based on index, and convert to double
   On Error Resume Next
   dat1 = CDbl(LVGetItemText(lParam1, m_lvHWnd))
   dat2 = CDbl(LVGetItemText(lParam2, m_lvHWnd))
   On Error GoTo 0
   
   '--- this sorts ascending
   LVCompareNumbers = Sgn(dat1 - dat2)
   
   '--- this sorts descending
   If SortOrder = lvDescending Then
      LVCompareNumbers = -LVCompareNumbers
   End If
End Function

Public Function LVGetItemText(lParam As Long, hWnd As Long) As String
   Dim objFind As LV_FINDINFO
   Dim Index As Long
   Dim objItem As LV_ITEM
   Dim nRet As Long
   
   ' Convert the input parameter to an index in the list view
   With objFind
      .Flags = LVFI_PARAM
      .lParam = lParam
   End With
   Index = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
   
   ' Obtain the name of the specified list view item
   With objItem
      .Mask = LVIF_TEXT
      .SubItem = m_lvSortColumn
      .Text = Space(32)
      .TextMax = Len(.Text)
   End With
   
   ' Grab the text
   nRet = SendMessage(hWnd, LVM_GETITEMTEXT, Index, objItem)
   If nRet Then
      LVGetItemText = Left$(objItem.Text, nRet)
   End If
End Function

' *********************************************************
'  Collection-based Sorting Routines
' *********************************************************
Public Function LVSortC(lv As ListView, ByVal Index As Long, ByVal ItemType As LVItemTypes, ByVal SortOrder As LVSortTypes) As Boolean
   Dim tmr As New CStopWatch
   
   ' turn off the default sorting of the control
   With lv
      .Sorted = False
      .SortKey = Index
      .SortOrder = SortOrder
   End With

   ' prepare collection of data for quicker lookups during callbacks
   tmr.Reset
   Call LVPrepareSortCollection(lv, Index, ItemType)
   BuildLookup = tmr.Elapsed
   
   ' initiate sort then delete collection
   tmr.Reset
   Call SendMessageLong(lv.hWnd, LVM_SORTITEMS, SortOrder, AddressOf LVCompare)
   PerformSort = tmr.Elapsed
   
   ' delete collection of sort data
   Set m_lvSortColl = Nothing
End Function

Private Function LVCompare(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As Long) As Long
   '--- this sorts ascending
   With m_lvSortColl
      LVCompare = Sgn(.Item("k" & lParam1) - .Item("k" & lParam2))
   End With
      
   '--- this sorts descending
   If SortOrder = lvDescending Then
      LVCompare = -LVCompare
   End If
End Function

Private Function LVPrepareSortCollection(lv As ListView, ByVal SubItemIndex As Long, ByVal ItemType As LVItemTypes) As Boolean
   Dim i As Long, n As Long
   Dim lvi As LV_ITEM
   Dim dat As Date
   
   ' initialize collection
   Set m_lvSortColl = New Collection
   
   ' obtain the ItemData value and string for each item in the list
   With lvi
      .Mask = LVIF_TEXT Or LVIF_PARAM
      .SubItem = SubItemIndex
      .TextMax = 256
      .Text = Space(256)
      If ItemType = lvDate Then
         For i = 1 To lv.ListItems.Count
            .Index = i - 1
            Call SendMessage(lv.hWnd, LVM_GETITEM, 0&, lvi)
            n = InStr(.Text, vbNullChar)
            If n > 1 Then
               On Error Resume Next
                  dat = CDate(Left$(.Text, n - 1))
               On Error GoTo 0
               m_lvSortColl.add dat, "k" & .Param
            Else
               m_lvSortColl.add 0, "k" & .Param
            End If
         Next i
      ElseIf ItemType = lvNumber Then
         For i = 1 To lv.ListItems.Count
            .Index = i - 1
            Call SendMessage(lv.hWnd, LVM_GETITEM, 0&, lvi)
            n = InStr(.Text, vbNullChar)
            If n > 1 Then
               m_lvSortColl.add CDbl(Left$(.Text, n - 1)), "k" & .Param
            Else
               m_lvSortColl.add 0, "k" & .Param
            End If
         Next i
      End If
   End With
End Function

' *********************************************************
'  IListItem-based Sorting Routines
' *********************************************************
Public Function LVSortI(lv As ListView, ByVal Index As Long, ByVal ItemType As LVItemTypes, ByVal SortOrder As LVSortTypes) As Boolean
   Dim tmr As New CStopWatch
   
   ' turn off the default sorting of the control
   With lv
      .Sorted = False
      .SortKey = Index
      .SortOrder = SortOrder
   End With

   ' no lookups used by this method
   BuildLookup = 0
   
   ' need to use module variables to let compare routine know
   ' which column and method to use
   m_lvSortColumn = Index
   m_lvSortType = ItemType
   
   ' fire off sorting
   tmr.Reset
   Call SendMessageLong(lv.hWnd, LVM_SORTITEMS, SortOrder, AddressOf LVCompareI)
   PerformSort = tmr.Elapsed
   
   ' delete collection of sort data
   Set m_lvSortColl = Nothing
End Function

Private Function LVCompareI(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As Long) As Long
   Static ListItem1 As ListItem
   Static ListItem2 As ListItem
   Static sItem1 As String
   Static sItem2 As String
   
   ' WARNING: This method *will* likely break in the future!
   ' Glom references to internal ListItem class using magic number
   CopyMem ListItem1, lParam1 + 84, 4
   CopyMem ListItem2, lParam2 + 84, 4
   
   ' Grab text items of interest
   If m_lvSortColumn = 0 Then
      sItem1 = ListItem1.Text
      sItem2 = ListItem2.Text
   Else
      sItem1 = ListItem1.SubItems(m_lvSortColumn)
      sItem2 = ListItem2.SubItems(m_lvSortColumn)
   End If
   
   ' Clean up hacked reference
   CopyMem ListItem1, Nothing, 4
   CopyMem ListItem2, Nothing, 4
   
   ' Perform ascending comparison
   On Error GoTo Failure
      Select Case m_lvSortType
         Case lvDate
            LVCompareI = Sgn(CDate(sItem1) - CDate(sItem2))
         Case lvNumber
            LVCompareI = Sgn(CDbl(sItem1) - CDbl(sItem2))
         Case lvBinary
            LVCompareI = StrComp(sItem1, sItem2, vbBinaryCompare)
         Case lvAlphabetic
            LVCompareI = StrComp(sItem1, sItem2, vbTextCompare)
         Case Else ' default ascending text
            LVCompareI = StrComp(sItem1, sItem2, vbTextCompare)
      End Select
   On Error GoTo 0
   
   ' Negate if descending
   If SortOrder = lvDescending Then
      LVCompareI = -LVCompareI
   End If
   Exit Function
   
Failure:
   ' Bail with 0 for failed comparison, because it's "just a visual sort" <g>
   ' Might want to return failure code in real app by setting flag here.
   Exit Function
End Function

