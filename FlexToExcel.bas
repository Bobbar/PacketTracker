Attribute VB_Name = "ToExcel"
Public Sub FlexToExcel(TargetGrid As MSHFlexGrid)
    Dim xlObject As Excel.Application
    Dim xlWB     As Excel.Workbook
    Dim Color1
    Dim FadeColor As Long
    Dim R         As Long
    TargetGrid.Redraw = False
    Set xlObject = New Excel.Application
    'This Adds a new woorkbook, you could open the workbook from file also
    Set xlWB = xlObject.Workbooks.Add
    Clipboard.Clear 'Clear the Clipboard
    With TargetGrid
        'Select Full Contents (You could also select partial content)
        .Col = 1               'From first column
        .Row = 0               'From first Row (header)
        .ColSel = .Cols - 1    'Select all columns
        .RowSel = .Rows - 1    'Select all rows
        Clipboard.SetText .Clip 'Send to Clipboard
    End With
    With xlObject.ActiveWorkbook.ActiveSheet
        .Range("A1").Select 'Select Cell A1 (will paste from here, to different cells)
        .Paste              'Paste clipboard contents
        For R = 1 To TargetGrid.Rows
            FadeColor = GetRealColor(GetFlexGridRowColor(TargetGrid, R - 1))
            ColorCodeToRGB FadeColor, iRed, iGreen, iBlue
            .Range("A" & R & ":I" & R).Interior.Color = RGB(iRed, iGreen, iBlue)
        Next R
    End With
    ' This makes Excel visible
    TargetGrid.Redraw = True
    xlObject.Visible = True
End Sub
Private Function GetRealColor(ByVal Color As OLE_COLOR) As Long
    Dim R As Long
    R = TranslateColor(Color, 0, GetRealColor)
    If R <> 0 Then 'raise an error
    End If
End Function
