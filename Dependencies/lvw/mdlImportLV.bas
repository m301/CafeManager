Attribute VB_Name = "mdlImportLV"
'download by http://www.codefans.net
Option Explicit

'ImportDBGrid:
' This Sub reads the DBGrid specified by dbGrd into clsTP.
' rstData has to be set to the recordset dbGrd gets its data from (it seems to be impossible to get DataSource at runtime !???)
' (e.g. if it's bound to Data1, rstData should be Data1.Recordset)
Sub ImportListView(clsTP As clsTablePrint, LV As ListView, Optional ByVal sngDesiredWidth As Single = -1, Optional ByVal bWithIcons As Boolean = True)
    Dim lCol As Long, lRow As Long, spcCount As Long
    Dim sngFXGGesWidth As Single, Fnt As StdFont
    
    clsTP.Rows = LV.ListItems.Count
    clsTP.Cols = LV.ColumnHeaders.Count
    If (Not (LV.SmallIcons Is Nothing)) And bWithIcons Then
        Set Fnt = LV.Parent.Font
        Set LV.Parent.Font = LV.Font
        spcCount = Int(LV.Parent.ScaleX(LV.SmallIcons.ImageWidth, vbPixels, LV.Parent.ScaleMode) / LV.Parent.TextWidth(" ")) + 2
        Set LV.Parent.Font = Fnt
    Else
        spcCount = 0
    End If
    clsTP.HeaderRows = 1
    clsTP.HasFooter = False
    clsTP.LineThickness = 1
    'Use double line width
    clsTP.HeaderLineThickness = 2 * clsTP.LineThickness

    'Use some reasonable default values:
    clsTP.CellXOffset = 60
    clsTP.CellYOffset = 30
    clsTP.CenterMergedHeader = False
    clsTP.ResizeCellsToPicHeight = True
    clsTP.PrintHeaderOnEveryPage = True
    
    With LV
        sngFXGGesWidth = 0
        Set clsTP.HeaderFont(-1, -1) = .Font
        Set clsTP.FontMatrix(-1, -1) = .Font
        For lCol = 0 To .ColumnHeaders.Count - 1
            With .ColumnHeaders(lCol + 1)
                Select Case .Alignment
                Case lvwColumnLeft
                    clsTP.ColAlignment(lCol) = eLeft
                Case lvwColumnRight
                    clsTP.ColAlignment(lCol) = eRight
                Case lvwColumnCenter
                    clsTP.ColAlignment(lCol) = eCenter
                End Select
                sngFXGGesWidth = sngFXGGesWidth + .Width
                clsTP.HeaderText(0, lCol) = .Text
            End With
        Next
        For lRow = 0 To .ListItems.Count - 1
            With .ListItems(lRow + 1)
                clsTP.TextMatrix(lRow, 0) = Space(spcCount) & .Text
                If (Not (LV.SmallIcons Is Nothing)) And bWithIcons Then
                    Set clsTP.PictureMatrix(lRow, 0) = LV.SmallIcons.ListImages(.SmallIcon).ExtractIcon
                End If
                For lCol = 1 To clsTP.Cols - 1
                    clsTP.TextMatrix(lRow, lCol) = .SubItems(lCol)
                Next
            End With
        Next
        For lCol = 0 To .ColumnHeaders.Count - 1
            If sngDesiredWidth > 0 Then
                clsTP.ColWidth(lCol) = (.ColumnHeaders(lCol + 1).Width / sngFXGGesWidth) * sngDesiredWidth
            Else
                clsTP.ColWidth(lCol) = .ColumnHeaders(lCol + 1).Width
            End If
        Next
    End With
End Sub
