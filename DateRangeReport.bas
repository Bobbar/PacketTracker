Attribute VB_Name = "modDateRangeReport"
Option Explicit
Public Sub DateRangeReport()
    StartTimer
    bolRunning = True
    Dim StartTime2 As Long
    StartTime2 = GetTickCount
    Dim rs As New ADODB.Recordset
    Dim Line, Row, s As Integer
    Dim Found            As Boolean
    Dim strUsedJobNums() As String
    Dim strQryRebuild()  As String
    Dim strQry           As String
    Dim dtTicketDate     As Date
    Dim TotT             As Single
    Dim Entries          As Integer
    Const ColorsRGB      As Integer = 255
    Dim CalcColor        As Integer
    On Error Resume Next
    Form1.Flexgrid1.Clear
    Form1.Flexgrid1.Visible = False
    Form1.Flexgrid1.Redraw = False
    Form1.Flexgrid1.Rows = 2
    Form1.Flexgrid1.FixedCols = 1
    Form1.Flexgrid1.FixedRows = 1
    strReportType = "Job Packets dated: " & (IIf(frmReportFilter.chkAllTickets.Value = 1, "Any Date", dtStartDate & " to " & dtEndDate))
    Form1.ShowData
    If frmReportFilter.chkHeatMap.Value = 1 Then
        QryEntryNumbers
        Form1.Flexgrid1.Cols = 11
    Else
        Form1.Flexgrid1.Cols = 10
    End If
    Line = 1
    strQry = "SELECT * FROM ticketdb.packetentrydb LEFT JOIN (ticketdb.packetlist) ON (packetlist.idJobNum=packetentrydb.idJobNum) WHERE" _
       & " ticketdb.packetentrydb.idDate=(SELECT MAX(s2.idDate) FROM ticketdb.packetentrydb s2 WHERE ticketdb.packetentrydb.idJobNum = s2.idJobNum) AND" _
       & (IIf(frmReportFilter.txtSearchJobNum.Text <> "", " packetlist.idJobNum LIKE '" & frmReportFilter.txtSearchJobNum.Text & "%' AND", "")) _
       & (IIf(frmReportFilter.txtSearchDesc.Text <> "", " packetlist.idDescription LIKE '%" _
       & frmReportFilter.txtSearchDesc.Text & "%' AND", "")) & (IIf(frmReportFilter.txtSearchPart.Text <> "", " packetlist.idPartNum LIKE '%" _
       & frmReportFilter.txtSearchPart.Text & "%' AND", "")) & (IIf(frmReportFilter.txtSearchSales.Text <> "", " packetlist.idSalesNum LIKE '%" _
       & frmReportFilter.txtSearchSales.Text & "%' AND", "")) & (IIf(frmReportFilter.txtSearchDraw.Text <> "", " packetlist.idDrawingNum LIKE '%" _
       & frmReportFilter.txtSearchDraw.Text & "%' AND", "")) & (IIf(frmReportFilter.txtSearchCust.Text <> "", " packetlist.idCustPONum LIKE '%" _
       & frmReportFilter.txtSearchCust.Text & "%' AND", "")) & " Order By packetentrydb.idDate Desc"
    strQryRebuild = Split(strQry, " AND ")
    strQry = ""
    If UBound(strQryRebuild) = 0 Then
        strQry = "SELECT * FROM ticketdb.packetentrydb LEFT JOIN (ticketdb.packetlist) ON (packetlist.idJobNum=packetentrydb.idJobNum) WHERE" & " ticketdb.packetentrydb.idDate=(SELECT MAX(s2.idDate) FROM ticketdb.packetentrydb s2 WHERE ticketdb.packetentrydb.idJobNum = s2.idJobNum) Order by packetentrydb.idDate Desc"
        GoTo SkipQryRebuild
    End If
    For s = 0 To UBound(strQryRebuild)
        If s = UBound(strQryRebuild) - 1 Then
            strQry = strQry + strQryRebuild(s) + " " + strQryRebuild(s + 1)
            Exit For
        Else
            strQry = strQry + strQryRebuild(s) + " AND "
        End If
    Next s
SkipQryRebuild:
    cn_global.CursorLocation = adUseClient
    Set rs = cn_global.Execute(strQry)
    If rs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        ShowBanner &HC0FFFF, "No packets were found that meet the specified criteria.", 300
        bolRunning = False
        Form1.HideData
        TotT = lngQryTimes(intQryIndex) * 0.001
        Form1.StatusBar1.Panels.Item(1).Text = "Custom search returned " & Line - 1 & " results in " & TotT & " seconds"
        Form1.Flexgrid1.Redraw = True
        Form1.Flexgrid1.Clear
        Form1.Flexgrid1.Visible = False
        Exit Sub
    End If
    Form1.Flexgrid1.TextMatrix(0, 1) = "Job Number"
    Form1.Flexgrid1.TextMatrix(0, 2) = "Part Number"
    Form1.Flexgrid1.TextMatrix(0, 3) = "Description"
    Form1.Flexgrid1.TextMatrix(0, 4) = "Sales Number"
    Form1.Flexgrid1.TextMatrix(0, 5) = "Customer/PO Number"
    Form1.Flexgrid1.TextMatrix(0, 6) = "Created By"
    Form1.Flexgrid1.TextMatrix(0, 7) = "Create Date"
    'form1.Flexgrid1.TextMatrix(0, 8) = "Status"
    Form1.Flexgrid1.TextMatrix(0, 8) = "Last Activity Date"
    Form1.Flexgrid1.TextMatrix(0, 9) = "Last Activity"
    If frmReportFilter.chkHeatMap.Value = 1 Then
        Form1.Flexgrid1.TextMatrix(0, 10) = "Entries"
    Else
    End If
    Form1.Flexgrid1.Rows = rs.RecordCount + 1
    ReDim strUsedJobNums(rs.RecordCount + 1)
    Row = 0
    dtStartDate = Format$(dtStartDate, "MM/DD/YYYY")
    dtEndDate = Format$(dtEndDate, "MM/DD/YYYY")
    Screen.MousePointer = vbHourglass
    Form1.pBar.Value = 0
    Form1.pBar.Max = rs.RecordCount
    Form1.frmpBar.Visible = True
    'DoEvents
    
   Debug.Print StopTimer
    Do Until rs.EOF
        With rs
            dtTicketDate = Format$(!idDate, "MM/DD/YYYY")
            If frmReportFilter.cmbPacketType.ListIndex = 0 Or frmReportFilter.cmbPacketType.ListIndex = 1 And !idUser = strSearchUser Or frmReportFilter.cmbPacketType.ListIndex = 2 And !idUserTo = strSearchUser Or frmReportFilter.cmbPacketType.ListIndex = 3 And !idUserFrom = strSearchUser Then
                If frmReportFilter.chkClosed.Value = 0 And frmReportFilter.chkFiled.Value = 0 And frmReportFilter.chkOpened.Value = 0 And frmReportFilter.chkInTransit.Value = 0 And frmReportFilter.chkReceived.Value = 0 And frmReportFilter.chkCreated.Value = 0 Then GoTo NoFilters
                'Start Ticket State filters
                If frmReportFilter.chkClosed.Value = 0 And !idStatus = "CLOSED" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkFiled.Value = 0 And !idStatus = "OPEN" And !idAction = "FILED" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkOpened.Value = 0 And !idStatus = "OPEN" Or frmReportFilter.chkOpened.Value = 0 And !idAction = "REOPENED" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkReceived.Value = 0 And !idAction = "RECEIVED" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkInTransit.Value = 0 And !idAction = "INTRANSIT" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkCreated.Value = 0 And !idAction = "CREATED" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkReOpened.Value = 0 And !idAction = "REOPENED" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                End If
NoFilters:
                If frmReportFilter.chkSF.Value = 0 And frmReportFilter.chkN.Value = 0 And frmReportFilter.chkRMT.Value = 0 And frmReportFilter.chkC.Value = 0 And frmReportFilter.chkW.Value = 0 And frmReportFilter.chkIM.Value = 0 Then GoTo NoPlantFilters             'Start Plant Filters
                If frmReportFilter.chkSF.Value = 0 And !idPlant = "STEEL FAB" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkN.Value = 0 And !idPlant = "NUCLEAR" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkRMT.Value = 0 And !idPlant = "ROCKY MT" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkC.Value = 0 And !idPlant = "CONTROLS" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkW.Value = 0 And !idPlant = "WOOSTER" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                ElseIf frmReportFilter.chkIM.Value = 0 And !idPlant = "INDUSTRIAL MACH" Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                End If
NoPlantFilters:
                If frmReportFilter.chkAllTickets.Value = 0 And dtTicketDate < dtStartDate Or dtTicketDate > dtEndDate Then 'Date range filter
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                Else
                    'let the ticket be displayed
                End If
                Found = InStr(1, vbNullChar & Join$(strUsedJobNums(), vbNullChar) & vbNullChar, vbNullChar & !idJobNum & vbNullChar) > 0
                If Found = False Then
                    strUsedJobNums(Row) = !idJobNum
                    Row = Row + 1
                    If frmReportFilter.chkHeatMap.Value = 1 Then
                        Entries = GetNumberOfEntries(!idJobNum)
                        CalcColor = ColorsRGB - (Entries * RGBMulti)
                        If CalcColor <= 0 Then CalcColor = 0
                    End If
                    Form1.Flexgrid1.TextMatrix(Line, 0) = Line
                    Form1.Flexgrid1.TextMatrix(Line, 1) = !idJobNum
                    Form1.Flexgrid1.TextMatrix(Line, 2) = !idPartNum
                    Form1.Flexgrid1.TextMatrix(Line, 3) = !idDescription
                    Form1.Flexgrid1.TextMatrix(Line, 4) = !idSalesNum
                    Form1.Flexgrid1.TextMatrix(Line, 5) = !idCustPoNum
                    Form1.Flexgrid1.TextMatrix(Line, 6) = !idCreator
                    Form1.Flexgrid1.TextMatrix(Line, 7) = !idCreateDate
                    Form1.Flexgrid1.TextMatrix(Line, 8) = !idDate
                    If frmReportFilter.chkHeatMap.Value = 1 Then
                        Form1.Flexgrid1.TextMatrix(Line, 10) = Entries
                    Else
                    End If
                    If !idAction = "CREATED" Then
                        Call Form1.FlexGridRowColor(Form1.Flexgrid1, Line, IIf(frmReportFilter.chkHeatMap.Value = 0, &H80C0FF, RGB(255, CalcColor, CalcColor)))
                        Form1.Flexgrid1.TextMatrix(Line, 9) = "Job packet was CREATED by " & !idCreator
                    ElseIf !idAction = "INTRANSIT" Then
                        Call Form1.FlexGridRowColor(Form1.Flexgrid1, Line, IIf(frmReportFilter.chkHeatMap.Value = 0, &H80FF80, RGB(255, CalcColor, CalcColor)))
                        Form1.Flexgrid1.TextMatrix(Line, 9) = !idUserFrom & " SENT the job packet to " & !idUserTo
                    ElseIf !idAction = "RECEIVED" Then
                        Call Form1.FlexGridRowColor(Form1.Flexgrid1, Line, IIf(frmReportFilter.chkHeatMap.Value = 0, &H80FFFF, RGB(255, CalcColor, CalcColor)))
                        Form1.Flexgrid1.TextMatrix(Line, 9) = !idUser & " RECEIVED the job packet from " & !idUserFrom
                    ElseIf !idStatus = "CLOSED" Then
                        Call Form1.FlexGridRowColor(Form1.Flexgrid1, Line, IIf(frmReportFilter.chkHeatMap.Value = 0, &H8080FF, RGB(255, CalcColor, CalcColor)))
                        Form1.Flexgrid1.TextMatrix(Line, 9) = !idUser & " CLOSED the job packet."
                    ElseIf !idStatus = "OPEN" And !idAction = "FILED" Then
                        Call Form1.FlexGridRowColor(Form1.Flexgrid1, Line, IIf(frmReportFilter.chkHeatMap.Value = 0, &HFF8080, RGB(255, CalcColor, CalcColor)))
                        Form1.Flexgrid1.TextMatrix(Line, 9) = !idUser & " FILED the job packet."
                    ElseIf !idAction = "REOPENED" Then
                        Call Form1.FlexGridRowColor(Form1.Flexgrid1, Line, IIf(frmReportFilter.chkHeatMap.Value = 0, &HFF80FF, RGB(255, CalcColor, CalcColor)))
                        Form1.Flexgrid1.TextMatrix(Line, 9) = !idUser & " REOPENED the job packet."
                    End If
                    Line = Line + 1
                ElseIf Found = True Then
                End If
ContNext:
                rs.MoveNext
                Form1.pBar.Value = .AbsolutePosition
                DoEvents
            ElseIf frmReportFilter.cmbPacketType.ListIndex = 1 And !idUser <> strSearchUser Or frmReportFilter.cmbPacketType.ListIndex = 2 And !idUserTo <> strSearchUser Or frmReportFilter.cmbPacketType.ListIndex = 3 And !idUserFrom <> strSearchUser Then
                strUsedJobNums(Row) = !idJobNum
                Row = Row + 1
                rs.MoveNext
                Form1.pBar.Value = .AbsolutePosition
                DoEvents
            End If
        End With
    Loop
    Form1.Flexgrid1.Rows = Line
    If Line > 1 Then
        Form1.Flexgrid1.Visible = True
        Form1.Flexgrid1.Redraw = True
        Form1.SizeTheSheet Form1.Flexgrid1
        Screen.MousePointer = vbDefault
        bolRunning = False
        Form1.HideData
        TotT = lngQryTimes(intQryIndex) * 0.001
        Form1.StatusBar1.Panels.Item(1).Text = "Custom search returned " & Line - 1 & " results in " & TotT & " seconds"
    Else
        bolRunning = False
        Form1.HideData
        Screen.MousePointer = vbDefault
        TotT = lngQryTimes(intQryIndex) * 0.001
        Form1.StatusBar1.Panels.Item(1).Text = "Custom search returned " & Line - 1 & " results in " & TotT & " seconds"
        ShowBanner &HC0FFFF, "No packets were found that meet the specified criteria.", 300
        Form1.Flexgrid1.Clear
        Form1.Flexgrid1.Visible = False
    End If
    Form1.Flexgrid1.ColSel = 0
    Erase strUsedJobNums
    Erase strEntries, intNumOfEntries
    Form1.pBar.Value = 0
    Form1.frmpBar.Visible = False
End Sub
