Attribute VB_Name = "Printing"
Public Type PrintData
    MSHGrid As MSHFlexGrid
    strText As String
    lngLeft As Long
    lngTop As Long
    lngRight As Long
    lngBottom As Long
    lngTextLeft As Long
    lngTextTop As Long
    intTextWidth As Integer
    intTextHeight As Integer
    intRowHeight As Integer
    lngTotalWidth As Long
    lngBackColor As Long
    intColWidth As Integer
    intPage As Integer
    intTotPage As Integer
End Type
Public GridPrint()        As PrintData
Public Const prntFontSize As Integer = 7
Public Sub PrintGridArray(ArrGrid() As PrintData)
    Dim Row         As Long, Col As Long
    Dim lngXMin     As Long, lngXMax As Long, lngYMin As Long, lngYMax As Long 'Constraints for cursor. Keeps it on da pappahs\
    Dim intCurPage  As Integer, intTotPages As Integer
    Dim lngPrevX    As Long, lngPrevY As Long
    Dim lngPageXLoc As Long, lngPageYLoc As Long
    lngXMin = 300
    lngXMax = 15000
    lngYMin = 1500
    lngYMax = 10800
    lngPageXLoc = lngXMax - 800
    lngPageYLoc = lngYMax + 400
    Printer.ScaleMode = 1
    'Printer.Orientation = vbPRORLandscape
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbSolid
    Printer.CurrentX = lngXMin
    Printer.CurrentY = lngYMin
    Printer.FontSize = prntFontSize
    intCurPage = 1
    intTotPages = ArrGrid(UBound(ArrGrid, 1), UBound(ArrGrid, 2)).intTotPage
    For Row = 0 To UBound(ArrGrid, 1)
        For Col = 1 To UBound(ArrGrid, 2)
            If ArrGrid(Row, Col).intPage > intCurPage Then
                intCurPage = ArrGrid(Row, Col).intPage
                Printer.NewPage
            End If
            Printer.Line (ArrGrid(Row, Col).lngLeft, ArrGrid(Row, Col).lngTop)-(ArrGrid(Row, Col).lngRight, ArrGrid(Row, Col).lngBottom), ArrGrid(Row, Col).lngBackColor, BF
            Printer.CurrentX = ArrGrid(Row, Col).lngTextLeft
            Printer.CurrentY = ArrGrid(Row, Col).lngTextTop
            Printer.Print ArrGrid(Row, Col).strText
            Printer.Line (ArrGrid(Row, Col).lngLeft, ArrGrid(Row, Col).lngTop)-(ArrGrid(Row, Col).lngRight, ArrGrid(Row, Col).lngBottom), vbBlack, B
            intCurPage = ArrGrid(Row, Col).intPage
            lngPrevX = Printer.CurrentX
            lngPrevY = Printer.CurrentY
            Printer.CurrentX = lngPageXLoc
            Printer.CurrentY = lngPageYLoc
            Printer.Print "Page: " & intCurPage & " of " & intTotPages
            Printer.CurrentX = lngPrevX
            Printer.CurrentY = lngPrevY
        Next Col
    Next Row
    Printer.EndDoc
    Dim blah
    blah = MsgBox(intTotPages & " pages sent to " & Printer.DeviceName, vbOKOnly + vbInformation, "Print")
End Sub
Public Sub PrintHeaders(strHeader As String, strSubHeader As String)
    Dim lngXMin           As Long, lngXMax As Long, lngYMin As Long, lngYMax As Long 'Constraints for cursor. Keeps it on da pappahs\
    Dim intCurPage        As Integer
    Dim lngPrevX          As Long, lngPrevY As Long
    Dim lngPageXLoc       As Long, lngPageYLoc As Long
    Dim lngHeaderWidth    As Long
    Dim intHeaderFontSize As Integer, intSubHeaderFontSize As Integer, intStampFontSize As Integer
    Dim lngCenter         As Long
    Dim lngHeaderGap      As Long
    lngXMin = 300
    lngXMax = 15000
    lngYMin = 100
    lngYMax = 10800
    intHeaderFontSize = 20
    intSubHeaderFontSize = 8
    intStampFontSize = 6
    lngHeaderGap = 100
    lngCenter = (lngXMax - lngXMin) / 2
    Printer.ScaleMode = 1
    Printer.Orientation = vbPRORLandscape
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbSolid
    Printer.CurrentX = lngXMin
    Printer.CurrentY = lngYMin
    Printer.FontSize = prntFontSize
    Printer.FontSize = intHeaderFontSize
    lngHeaderWidth = Printer.TextWidth(strHeader)
    Printer.CurrentX = lngCenter - (lngHeaderWidth / 2)
    Printer.CurrentY = lngYMin
    Printer.Print strHeader
    Printer.CurrentY = Printer.CurrentY + lngHeaderGap
    Printer.FontSize = intSubHeaderFontSize
    lngHeaderWidth = Printer.TextWidth(strSubHeader)
    Printer.CurrentX = lngXMin 'lngCenter - (lngHeaderWidth / 2)
    Printer.Print strSubHeader
    Printer.CurrentY = Printer.CurrentY + lngHeaderGap
    Printer.CurrentX = lngXMin
    Printer.FontSize = intStampFontSize
    Printer.Print "Print Date: " & Now
    Printer.CurrentX = lngXMin
    Printer.Print "Printed By: " & GetFullName(strLocalUser)
End Sub
Public Function BoundedText(ByVal ptr As Object, _
                            ByVal txt As String, _
                            ByVal max_wid As Single) As String
    On Error GoTo EF
    Printer.FontSize = prntFontSize
    Do While Printer.TextWidth(txt) > max_wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    BoundedText = txt
EF:
End Function
