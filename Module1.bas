Attribute VB_Name = "MyMod"
Option Explicit
Public Type tLine
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    Left As Long
    Top As Long
    Width As Long
    Height As Integer
    Color As Long
    Visible As Boolean
    FillStyle As Long
End Type
Public Type tTxtBox
    Left As Long
    Top As Long
    Width As Long
    Height As Integer
    Color As Long
    Text As String
    Visible As Boolean
End Type
Public dLine()      As tLine
Public dGrid()      As tLine
Public dDayLine()   As tLine
Public dPointLine   As tLine
Public dAction()    As tTxtBox
Public dNote()      As tTxtBox
Public dTimer       As tTxtBox
Public strSQLDriver As String
Public bolHook      As Boolean
Public strINILoc    As String
Public strUserTo    As String, strSelectUserTo As String, strUserFrom As String, strCurUser As String
Public strTicketAction, strTicketStatus As String
Public bolHasTicket      As Boolean
Public strServerAddress  As String, strUsername As String, strPassword As String, strSearchUser As String, strPlant As String
Public strUserIndex()    As String
Public bolOpenForm       As Boolean, bolOpenConfirm    As Boolean
Public intFormHMax       As Integer, intFormHMin As Integer
Public strReportType     As String
Public dtStartDate       As Date
Public dtEndDate         As Date
Public sAddlMsg          As String
Public bolCancelPrint    As Boolean
Public strTicketComment  As String
Public strLatestComment  As String
Public strSortMode       As String
Public bolPrinting       As Boolean
Public FlexINLastSel(2)  As Integer
Public FlexOUTLastSel(2) As Integer
Public DrawDayLines      As Boolean
Public MouseX, MouseY, MouseXPrev, MouseYPrev As Long
Public strConfirmClickCase As String
Public ProgHasFocus        As Boolean
Public FlexHistLastTopRow  As Integer
Public Const intRowH       As Integer = 400
Public strLocalUser        As String
Public intMovement, intConfirmMovement, intMovementAccel          As Integer
Public bolOptionClicked     As Boolean
Public HistoryIcons()       As StdPicture
Public HelpPics()           As StdPicture
Public ButtonPics()         As StdPicture
Public TabPics(3)           As StdPicture
Public picDataPics(2)       As StdPicture
Public WhichGrid            As MSHFlexGrid
Public TicketHours(99)      As Single
Public TicketAction(99)     As String
Public TicketActionText(99) As String
Public TicketDate(99)       As String
Public TotalTime            As Single
Public LStep                As Single
Public Entry, Clicks As Integer
Public strTimelineComments() As String
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public StartTime As Long
Private Const WM_ACTIVATEAPP = &H1C
Private Const GWL_WNDPROC = -4
Public gHW As Long
Public Declare Function SendMessage _
               Lib "user32.dll" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal Msg As Long, _
                                     wParam As Any, _
                                     lParam As Any) As Long
Private Declare Function CallWindowProc _
                Lib "user32.dll" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long
Public Declare Function SetWindowLong _
               Lib "user32.dll" _
               Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long
Public Const CB_SHOWDROPDOWN As Long = &H14F
Private Const WM_MOUSEWHEEL = &H20A
Dim LocalHwnd                        As Long
Dim LocalPrevWndProc                 As Long
Dim MyForm                           As Form
Private lpPrevWndProc                As Long
Private lpWndProcTmp                 As Long
Private Const WM_POWERBROADCAST      As Long = &H218
Private Const PBT_APMRESUMEAUTOMATIC As Long = &H12
Private Const PBT_APMSUSPEND         As Long = &H4
Public HelpClosed                    As Boolean
Public intSearchWaitTicks, intSearchWait As Integer
Public bolCanEdit As Boolean
Public ActiveText As String
Public PrevPartNum, PrevDrawNoRev, PrevCustPoNo, PrevSalesNo, PrevDescription As String
Public FOCUS                 As Integer
Public Const colCreate       As Long = &H80C0FF
Public Const colInTransit    As Long = &H80FF80
Public Const colReceived     As Long = &H80FFFF
Public Const colClosed       As Long = &H8080FF
Public Const colFiled        As Long = &HFF8080
Public Const colReopened     As Long = &HFF80FF
Public bolWaitToClose        As Boolean
Public intTicksWaited        As Integer
Public bolBannerOpen         As Boolean
Public intTicksToWait        As Integer
Public intPrevInPackets      As Integer
Public sngShapeResize        As Single
Public intShpTimerStartWidth As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public EditMode              As Boolean
Public BannerColor(99)       As Long
Public BannerText(99)        As String
Public BannerTicks(99)       As Integer
Public BannerCase(99)        As String
Public BannerFontColor(99)   As Long
Public bolBannerCleared      As Boolean
Public Const strDBDateFormat As String = "YYYY-MM-DD"
Public bolMessageDelivered   As Boolean
Public bolInitialLoad        As Boolean
Public intFlexGridInLastRow, intFlexGridOutLastRow As Integer
Public lngQryStart, lngQryEnd As Long
Public intQryIndex               As Integer
Public lngQryTimes(20)           As Long
Public strLastJobNum             As String
Public dtLatestHistDate          As String
Public Const strDBDateTimeFormat As String = "YYYY-MM-DD hh:mm:ss"
Public intCachedBanners, intCurrentCache As Integer
Public bolRunning        As Boolean
Public Const RGBMulti    As Integer = 4
Public intColorFlash     As Integer
Public bolRefreshRunning As Boolean
Public iRed              As Integer, iGreen As Integer, iBlue As Integer, iStep As Integer
Public r1, r2, g1, g2, b1, b2 As Integer
Public strEntries()      As String
Public intNumOfEntries() As Integer
Public Declare Function FlashWindow _
               Lib "user32" (ByVal hwnd As Long, _
                             ByVal bInvert As Long) As Long
Public Const Invert = 1
Public bolNewHistWindow As Boolean
Public bolIsAdmin       As Boolean
Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue _
                Lib "advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpValueName As String, _
                                       lpcbValueName As Long, _
                                       ByVal lpReserved As Long, _
                                       lpType As Long, _
                                       lpData As Any, _
                                       lpcbData As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (dest As Any, _
                                       Source As Any, _
                                       ByVal numBytes As Long)
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234
Const KEY_READ = &H20019 ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
Public m_cIni As New cInifile
Public Type UserInfo
    UserName As String
    FullName As String
End Type
Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean
Public total     As Currency
Public Ctr1      As Currency, Ctr2 As Currency, Freq As Currency
Global cn_global As New ADODB.Connection
Public Declare Function RoundRect _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long, _
                            ByVal X3 As Long, _
                            ByVal Y3 As Long) As Long
Public strCurrentPacketCreator As String, strCurrentPacketOwner As String
Public Function DBConcurrent() As Integer 'Does the state of the packet stored locally match what's in the DB? NO = 0, YES = 1, NOTFOUND = 2
    Dim strSQL1 As String
    Dim rs      As New ADODB.Recordset
    On Error GoTo errs
    DBConcurrent = 0
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.packetentrydb LEFT JOIN (ticketdb.packetlist) ON (packetlist.idJobNum=packetentrydb.idJobNum) WHERE packetlist.idJobNum = '" & Form1.txtJobNo.Text & "' ORDER BY packetentrydb.idDate DESC"
    Set rs = cn_global.Execute(strSQL1)
    With rs
        If rs.RecordCount < 1 Then
            DBConcurrent = 2
        Else
            If strTicketAction <> !idAction Or strTicketStatus <> !idStatus Then
                DBConcurrent = 0
            ElseIf strTicketAction = !idAction And strTicketStatus = !idStatus Then
                DBConcurrent = 1
            End If
        End If
    End With
    Exit Function
errs:
    If Err.Number = -2147467259 Then
        Form1.CommsDown
    Else
        Form1.CommsUp
    End If
End Function

Public Function GetEmail(strUsername As String) As String
    Dim i As Integer
    For i = 0 To UBound(strUserIndex, 2)
        If strUserIndex(0, i) = strUsername Then
            GetEmail = UCase$(strUserIndex(2, i))
            Exit Function
        End If
    Next i
End Function

Public Function GetFullName(strUsername As String) As String
    Dim i As Integer
    For i = 0 To UBound(strUserIndex, 2)
        If strUserIndex(0, i) = strUsername Then
            GetFullName = UCase$(strUserIndex(1, i))
            Exit Function
        End If
    Next i
End Function
Public Sub SendEmailToQueue(SendRec As String, _
                            MailFrom As String, _
                            MailTo As String, _
                            JobNum As String, _
                            strComment As String)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    strSQL1 = "INSERT INTO emailqueue (idSendOrRec,idFrom,idTo,idJobNum,idComment)" & " VALUES ('" & SendRec & "','" & MailFrom & "','" & MailTo & "','" & JobNum & "','" & strComment & "')"
    Set rs = New ADODB.Recordset
    cn_global.CursorLocation = adUseClient
    Set rs = cn_global.Execute(strSQL1)
End Sub
Public Sub StartTimer()
    total = 0
    QueryPerformanceFrequency Freq
    QueryPerformanceCounter Ctr1
End Sub
Public Function StopTimer() As Double
    StopTimer = 0
    QueryPerformanceCounter Ctr2
    total = total + (Ctr2 - Ctr1)
    StopTimer = Round(CDbl(total / Freq) * 1000, 3)
End Function
Public Function ReturnEmpInfo(strUsername As Variant) As UserInfo
    ReturnEmpInfo.UserName = vbNull
    ReturnEmpInfo.FullName = vbNull
    Dim i As Integer
    For i = 0 To UBound(strUserIndex, 2)
        If strUserIndex(0, i) = strUsername Then
            ReturnEmpInfo.UserName = strUserIndex(0, i)
            ReturnEmpInfo.FullName = strUserIndex(1, i)
            Exit Function
        End If
    Next i
    MsgBox (strUsername & " not found")
End Function
Public Sub MySort(ByRef pvarArray As Variant)
    Dim i               As Long
    Dim c               As Integer
    Dim v               As Integer
    Dim lngHighValIndex As Long
    Dim varSwap()       As Variant
    Dim lngMax          As Long
    ReDim varSwap(UBound(pvarArray, 1))
    lngMax = UBound(pvarArray, 2)
    For c = 0 To lngMax
        lngHighValIndex = lngMax - c
        For v = 0 To UBound(varSwap)
            varSwap(v) = pvarArray(v, lngMax - c)
        Next v
        For i = 0 To lngMax - c
            If pvarArray(0, i) < pvarArray(0, lngHighValIndex) Then lngHighValIndex = i
        Next
        For v = 0 To UBound(varSwap)
            pvarArray(v, lngMax - c) = pvarArray(v, lngHighValIndex)
            pvarArray(v, lngHighValIndex) = varSwap(v)
        Next v
    Next c
End Sub
Public Function GetINIValue(sUser As Variant) As Integer
    With m_cIni
        .Path = strINILoc
        .Section = "HITS"
        .Key = sUser
        .Default = "0"
        GetINIValue = .Value
        ' If Not (.Success) Then
        '  GetINIValue = "0"
        'End If
    End With
End Function
Public Sub SetINIValue(sUser As String, iHits As Integer)
    With m_cIni
        .Path = strINILoc
        .Section = "HITS"
        .Key = sUser
        .Value = iHits
        If Not (.Success) Then
            MsgBox "Failed to set value.", vbInformation
        End If
    End With
    'ShowIniAndParameters
End Sub
Public Sub CreateINI()
    Dim i   As Long
    Dim iNd As Long
    ' Create an Ini File to play with:
    On Error Resume Next
    If Dir$(strINILoc) = "" Then
        MkDir Environ$("APPDATA") & "\JPT\"
        With m_cIni
            .Path = strINILoc
            .Section = "INFO"
            .Key = "VERSION"
            .Value = App.Major & App.Minor & App.Revision
        End With
    End If
End Sub
Public Sub DeleteEntry(strGUID As String, strDesc As String)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    On Error Resume Next
    Dim blah
    blah = MsgBox("Are you sure you want to delete this entry?" & vbCrLf & vbCrLf & "      Job #: " & Form1.txtJobNo.Text & vbCrLf & "Description:  " & strDesc & vbCrLf & vbCrLf & "      This cannot be undone!", vbCritical + vbYesNo, "Are you sure?")
    If blah = vbNo Then
        Exit Sub
    ElseIf blah = vbYes Then
    End If
    Form1.ShowData
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * From packetentrydb Where idGUIDEntry = '" & strGUID & "'"
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .Delete
        .Update
    End With
    Form1.HideData
    If Err.Number = 0 And DBConcurrent = 0 Then
        ShowBanner colInTransit, "Single entry deleted successfully."
    Else
        blah = MsgBox("An update was attempted but the result was unexpected!" & vbCrLf & "(The state of the packet did not change as expected)", vbExclamation + vbOKOnly, "Something's wrong...")
    End If
    Form1.RefreshAfterEdit
    Form1.GetMyPackets
    Form1.SetControls
    Form1.cmdSubmit.Enabled = False
    Form1.optMove.Value = False
    Form1.optReceive.Value = False
    Form1.optMove.Value = False
    Form1.optClose.Value = False
    Form1.optCreate.Value = False
    Form1.optReOpen.Value = False
    Form1.optFile.Value = False
    bolOptionClicked = False
    Form1.imgComment.Enabled = False
    Form1.RefreshAll
    Form1.RefreshHistory
    Form1.GetMyPackets
End Sub
Public Sub DeletePacket(JobNum As String)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Form1.ShowData
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT idJobnum From packetlist Where idJobNum = '" & JobNum & "'"
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    'Do Until rs.EOF
    With rs
        .Delete
        '.MoveNext
    End With
    'Loop
    Form1.HideData
    Form1.ClearFields
    ShowBanner colInTransit, "Packet Deleted!"
End Sub
' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
' SYNCHRONIZE))
' Enumerate values under a given registry key
'
' returns a collection, where each element of the collection
' is a 2-element array of Variants:
' element(0) is the value name, element(1) is the value's value
Function EnumRegistryValues(ByVal hKey As Long, ByVal KeyName As String) As Collection
    Dim handle            As Long
    Dim Index             As Long
    Dim valueType         As Long
    Dim Name              As String
    Dim nameLen           As Long
    Dim resLong           As Long
    Dim resString         As String
    Dim dataLen           As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retVal            As Long
    ' initialize the result
    Set EnumRegistryValues = New Collection
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    Do
        ' this is the max length for a key name
        nameLen = 260
        Name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        ' retrieve the value's name
        valueInfo(0) = Left$(Name, nameLen)
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        Index = Index + 1
    Loop
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
End Function
' get the list of ODBC drivers through the registry
'
' returns a collection of strings, each one holding the
' name of a driver, e.g. "Microsoft Access Driver (*.mdb)"
'
' requires the EnumRegistryValues function
Function GetODBCDrivers() As Collection
    Dim res    As Collection
    Dim values As Variant
    ' initialize the result
    Set GetODBCDrivers = New Collection
    ' the names of all the ODBC drivers are kept as values
    ' under a registry key
    ' the EnumRegistryValue returns a collection
    For Each values In EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBCINST.INI\ODBC Drivers")
        ' each element is a two-item array:
        ' values(0) is the name, values(1) is the data
        If StrComp(values(1), "Installed", 1) = 0 Then
            ' if installed, add to the result collection
            GetODBCDrivers.Add values(0), values(0)
        End If
    Next
End Function
Public Sub FindMySQLDriver()
    GetODBCDrivers
    Dim i           As Integer
    Dim strPossis() As String
    Dim blah
    ReDim strPossis(0)
    For i = 1 To GetODBCDrivers.Count
        If InStr(1, GetODBCDrivers.Item(i), "MySQL") Then
            strPossis(UBound(strPossis)) = GetODBCDrivers.Item(i)
            ReDim Preserve strPossis(UBound(strPossis) + 1)
        End If
    Next i
    If UBound(strPossis) > 1 Then
        blah = MsgBox("Multiple MySQL Drivers detected!", vbExclamation + vbOKOnly, "Gasp!")
        strSQLDriver = strPossis(0)
    Else
        strSQLDriver = strPossis(0)
    End If
End Sub
Public Function CheckForAdmin(strLocalUser As String) As Boolean
    CheckForAdmin = False
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim i
    strSQL1 = "SELECT idAdmins FROM users"
    cn_global.CursorLocation = adUseClient
    Set rs = cn_global.Execute(strSQL1)
    With rs
        Do Until .EOF
            If UCase$(!idAdmins) = UCase$(strLocalUser) Then CheckForAdmin = True
            .MoveNext
        Loop
    End With
End Function
Public Sub CopyGridHistory(Source As MSHFlexGrid, dest As MSHFlexGrid)
    Dim R, c As Integer
    Dim GridImg As Image
    bolNewHistWindow = True
    dest.Redraw = False
    Source.Redraw = False
    'copy properties
    dest.FocusRect = Source.FocusRect
    dest.HighLight = Source.HighLight
    dest.BandDisplay = Source.BandDisplay
    dest.FillStyle = Source.FillStyle
    dest.FocusRect = Source.FocusRect
    dest.GridLines = Source.GridLines
    dest.GridLinesFixed = Source.GridLinesFixed
    dest.GridLinesUnpopulated = Source.GridLinesUnpopulated
    dest.MergeCells = Source.MergeCells
    dest.ScrollBars = Source.ScrollBars
    dest.SelectionMode = Source.SelectionMode
    dest.WordWrap = Source.WordWrap
    dest.Font = Source.Font
    dest.Font.Size = Source.Font.Size
    dest.Cols = 4
    dest.Rows = Source.Rows
    dest.FixedRows = 0
    dest.FixedCols = 0
    For R = 0 To Source.Rows - 1
        For c = 0 To 2
            dest.TextMatrix(R, c) = Source.TextMatrix(R, c)
            dest.Row = R
            dest.col = c
            Source.Row = R
            Source.col = c
            dest.CellFontBold = Source.CellFontBold
            dest.CellFontItalic = Source.CellFontItalic
            dest.CellAlignment = Source.CellAlignment
            dest.CellFontSize = Source.CellFontSize
        Next c
        Call Form1.FlexGridRowColor(dest, R, GetFlexGridRowColor(Source, R))
        dest.Row = R
        dest.col = 0
        Set dest.CellPicture = Source.CellPicture 'HistoryIcons(1)
        dest.CellPictureAlignment = flexAlignCenterCenter
        dest.RowHeight(R) = Source.RowHeight(R)
    Next R
    'Dest.RowHeight(0) = 0
    dest.ColWidth(0) = 1000
    dest.ColWidth(1) = 8500
    dest.ColWidth(3) = 0
    'Dest.Rows = Dest.Rows - 1
    'Form1.SizeTheSheet Dest
    dest.Redraw = True
    Source.Redraw = True
End Sub
Public Sub CopyGrid(Source As MSHFlexGrid, dest As MSHFlexGrid)
    Dim R, c As Integer
    Dim GridImg As Image
    bolNewHistWindow = False
    dest.Redraw = False
    Source.Redraw = False
    dest.Cols = Source.Cols
    dest.Rows = Source.Rows
    dest.FixedRows = Source.FixedRows
    dest.FixedCols = Source.FixedCols
    'Dest.Font.Size = Source.Font.Size
    For R = 0 To Source.Rows - 1
        For c = 0 To Source.Cols - 1
            dest.TextMatrix(R, c) = Source.TextMatrix(R, c)
        Next c
        Call Form1.FlexGridRowColor(dest, R, GetFlexGridRowColor(Source, R))
    Next R
    Form1.SizeTheSheet dest
    dest.Redraw = True
    Source.Redraw = True
End Sub
Public Function GetFlexGridRowColor(FlexGrid As MSHFlexGrid, ByVal lngRow As Long) As Long
    Dim lngPrevCol       As Long
    Dim lngPrevColSel    As Long
    Dim lngPrevRow       As Long
    Dim lngPrevRowSel    As Long
    Dim lngPrevFillStyle As Long
    If lngRow > FlexGrid.Rows - 1 Then
        Exit Function
    End If
    '    lngPrevCol = FlexGrid.col
    '    lngPrevRow = FlexGrid.Row
    '    lngPrevColSel = FlexGrid.ColSel
    '    lngPrevRowSel = FlexGrid.RowSel
    '    lngPrevFillStyle = FlexGrid.FillStyle
    FlexGrid.col = FlexGrid.FixedCols
    FlexGrid.Row = lngRow
    'FlexGrid.ColSel = FlexGrid.Cols - 1
    FlexGrid.RowSel = lngRow
    ' FlexGrid.FillStyle = flexFillRepeat
    GetFlexGridRowColor = FlexGrid.CellBackColor
    '    FlexGrid.col = lngPrevCol
    '    FlexGrid.Row = lngPrevRow
    '    FlexGrid.ColSel = lngPrevColSel
    '    FlexGrid.RowSel = lngPrevRowSel
    '    FlexGrid.FillStyle = lngPrevFillStyle
End Function
Public Function ColorCodeToRGB(lColorCode As Long, _
                               iRed As Integer, _
                               iGreen As Integer, _
                               iBlue As Integer) As Boolean
    ' 1996/01/16 Return the individual colors for lColorCode.
    ' 1996/07/15 Use Tip 171: Determining RGB Color Values, MSDN July 1996.
    ' Enter with:
    '     lColorCode contains the color to be converted
    '
    ' Return:
    '     iRed    contains the red component
    '     iGreen  the green component
    '     iBlue   the blue component
    '
    Dim lColor As Long
    lColor = lColorCode      'work long
    iRed = lColor Mod &H100  'get red component
    lColor = lColor \ &H100  'divide
    iGreen = lColor Mod &H100 'get green component
    lColor = lColor \ &H100  'divide
    iBlue = lColor Mod &H100 'get blue component
    ColorCodeToRGB = True
End Function
Public Function CheckForBlanks(CurrentControl As String) As Boolean
    CheckForBlanks = True
    With Form1
        If Trim$(Form1.Controls(CurrentControl).Text) = "" Then
            CheckForBlanks = True
        Else
            CheckForBlanks = False
        End If
    End With
End Function
Public Sub QryEntryNumbers()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim i
    strSQL1 = "SELECT idJobNum, COUNT(*) FROM packetentrydb GROUP BY idJobNum"
    cn_global.CursorLocation = adUseClient
    Set rs = cn_global.Execute(strSQL1)
    For i = 0 To rs.RecordCount
        ReDim Preserve strEntries(i + 1)
        ReDim Preserve intNumOfEntries(i + 1)
        strEntries(i) = rs.Fields(0)
        intNumOfEntries(i) = rs.Fields(1)
        rs.MoveNext
        If rs.EOF Then Exit For
    Next i
End Sub
Public Function GetNumberOfEntries(JobNum As String) As Integer
    Dim iPos As Integer
    GetNumberOfEntries = 0
    iPos = ArrayPosition(JobNum, strEntries)
    GetNumberOfEntries = intNumOfEntries(iPos)
End Function
Public Function ArrayPosition(ByVal FindValue As Variant, arrSearch As Variant) As Long
    ArrayPosition = -1  'Set default value of "not found"
    On Error GoTo LocalError
    If Not IsArray(arrSearch) Then Exit Function
    FindValue = UCase$(FindValue)  'no need for the If, you can UCase anything (faster this way)
    Dim lngLoop As Long
    For lngLoop = LBound(arrSearch) To UBound(arrSearch)
        If UCase$(arrSearch(lngLoop)) = FindValue Then
            ArrayPosition = lngLoop
            Exit Function
        End If
    Next lngLoop
    Exit Function
LocalError:
    'Nothing
End Function
Public Function ChangesMade() As Boolean
    ChangesMade = False
    If UCase$(Form1.txtPartNoRev) <> PrevPartNum Or UCase$(Form1.txtDrawNoRev) <> PrevDrawNoRev Or UCase$(Form1.txtCustPoNo) <> PrevCustPoNo Or UCase$(Form1.txtSalesNo) <> PrevSalesNo Or UCase$(Form1.txtTicketDescription) <> PrevDescription Then
        ChangesMade = True
    Else
        ChangesMade = False
    End If
End Function
Public Sub CloseBanner()
    bolWaitToClose = False
    intTicksWaited = intTicksToWait
End Sub
Public Sub Wait(ByVal DurationMS As Long)
    Dim EndTime As Long
    EndTime = GetTickCount + DurationMS
    Do While EndTime > GetTickCount
        DoEvents
        Sleep 1
    Loop
End Sub
Public Sub ClearBanners()
    bolBannerCleared = True
    CloseBanner
    Form1.tmrBannerWait.Enabled = False
    Erase BannerColor
    Erase BannerText
    Erase BannerTicks
    Erase BannerCase
    Erase BannerFontColor
    intCachedBanners = 0
    intCurrentCache = -1
End Sub
Public Sub ShowBanner(Color As Long, _
                      Text As String, _
                      Optional Ticks As Integer, _
                      Optional ClickCase As String, _
                      Optional FontColor As Long)
    bolBannerCleared = False
    BannerColor(intCachedBanners) = Color
    BannerText(intCachedBanners) = Text
    BannerTicks(intCachedBanners) = Ticks
    BannerCase(intCachedBanners) = ClickCase
    BannerFontColor(intCachedBanners) = FontColor
    intCachedBanners = intCachedBanners + 1
    If bolBannerOpen = False Then
        intCurrentCache = intCurrentCache + 1
        OpenCloseBanner BannerColor(intCurrentCache), BannerText(intCurrentCache), BannerTicks(intCurrentCache), BannerCase(intCurrentCache), BannerFontColor(intCurrentCache)
    Else
        Form1.tmrBannerWait.Enabled = True
    End If
End Sub

Public Sub OpenCloseBanner(Color As Long, _
                           Text As String, _
                           Optional Ticks As Integer, _
                           Optional ClickCase As String, _
                           Optional FontColor As Long)
    bolBannerOpen = True
    If ClickCase <> "" Then
        strConfirmClickCase = ClickCase
    Else
        strConfirmClickCase = ""
    End If
    With Form1
        .lblConfirm.WordWrap = False
        If Ticks > 0 Then
            intTicksToWait = Ticks
        Else
            intTicksToWait = 170
        End If
        dTimer.Width = intShpTimerStartWidth
        .frmConfirm.BackColor = Color
        If FontColor <> 0 Then
            .lblConfirm.ForeColor = FontColor
            .Border.BorderColor = FontColor
            dTimer.Color = FontColor
        Else
            .lblConfirm.ForeColor = vbBlack
            .Border.BorderColor = vbBlack
            dTimer.Color = vbBlack
        End If
        .lblConfirm.Caption = Text
        .lblConfirm.Left = 240
        If .lblConfirm.Width >= .Width Then
            .lblConfirm.WordWrap = True
            .lblConfirm.Width = .Width - 550
            .Border.Height = .frmConfirm.Height - 240
            .frmConfirm.Height = .Border.Height + 240  '585
        Else
            .lblConfirm.WordWrap = False
            .lblConfirm.Height = 270
            .Border.Height = 855 '615
            .frmConfirm.Height = .Border.Height + 240
        End If
        .frmConfirm.Width = .lblConfirm.Width + 500
        If .frmConfirm.Width <= dTimer.Width Then 'make sure banner is no smaller than count down bar
            .frmConfirm.Width = dTimer.Width + 300
        End If
        .Border.Width = .frmConfirm.Width - 235
        .frmConfirm.Left = .Width / 2 - .frmConfirm.Width / 2 - 50
        .frmConfirm.Visible = True
        intConfirmMovement = 5
        intTicksWaited = 0
        bolOpenConfirm = True
        .tmrConfirmSlider.Enabled = True
        dTimer.Top = .Border.Height + 70
        dTimer.Left = .frmConfirm.Width / 2 - dTimer.Width / 2
        sngShapeResize = dTimer.Width / intTicksToWait
        .lblConfirm.Left = .frmConfirm.Width / 2 - .lblConfirm.Width / 2
        .frmConfirm.Line (dTimer.Left, dTimer.Top)-(dTimer.Left + dTimer.Width, dTimer.Top + 85), dTimer.Color, BF
        RoundRect .frmConfirm.hdc, (.Border.Left / Screen.TwipsPerPixelY), (.Border.Top / Screen.TwipsPerPixelY), ((.Border.Left / Screen.TwipsPerPixelY) + (.Border.Width / Screen.TwipsPerPixelY)), ((.Border.Top / Screen.TwipsPerPixelY) + (.Border.Height / Screen.TwipsPerPixelY)), 10, 10
        .frmConfirm.CurrentX = (.lblConfirm.Left) '/ Screen.TwipsPerPixelX)
        .frmConfirm.CurrentY = (.lblConfirm.Top) '/ Screen.TwipsPerPixelY)
        .frmConfirm.ForeColor = .lblConfirm.ForeColor
        .frmConfirm.DrawStyle = 0
        .frmConfirm.Font.Name = "Tahoma"
        .frmConfirm.Font.Size = 11
        .frmConfirm.FontTransparent = True
        .frmConfirm.Print .lblConfirm.Caption
        .lblClose.Left = .Border.Width - 140
    End With
End Sub
Public Sub Hook(ByVal gHW As Long, HKflg As Boolean)
    Static IsHooked As Boolean
    If HKflg Xor IsHooked Then
        If HKflg Then
            lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
        Else
            SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
        End If
        IsHooked = HKflg
    End If
End Sub
Public Function WindowProc(ByVal hw As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    If uMsg = WM_POWERBROADCAST Then
        If wParam = PBT_APMRESUMEAUTOMATIC Then
            Form1.tmrRefresher.Enabled = True
            Form1.tmrDateTime.Enabled = True
        ElseIf wParam = PBT_APMSUSPEND Then
            Form1.tmrRefresher.Enabled = False
            Form1.tmrDateTime.Enabled = False
        End If
    End If
    If uMsg = WM_ACTIVATEAPP Then
        'Check to see if Activating the application
        If wParam = 0 Then      'Application Received Focus
            'ProgHasFocus = False
        Else
            'Application Lost Focus
        End If
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
Public Sub MouseWheel(ByVal MouseKeys As Long, _
                      ByVal Rotation As Long, _
                      ByVal Xpos As Long, _
                      ByVal Ypos As Long)
    Dim NewValue As Long
    Dim LStep    As Single
    On Error Resume Next
    With WhichGrid
        LStep = .Height / .RowHeight(1)
        LStep = Int(LStep)
        LStep = LStep - 20
        If LStep < 20 Then
            LStep = 1
        End If
        If Rotation > 0 Then
            NewValue = .TopRow - LStep
            If NewValue < 1 Then
                NewValue = 0
            End If
        Else
            NewValue = .TopRow + LStep
            If NewValue > .Rows - 1 Then
                NewValue = .Rows - 1
            End If
        End If
        .TopRow = NewValue
        ' FlexHistLastTopRow = NewValue
    End With
End Sub
Private Function WindowProc2(ByVal Lwnd As Long, _
                             ByVal Lmsg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long
    Dim MouseKeys As Long
    Dim Rotation  As Long
    Dim Xpos      As Long
    Dim Ypos      As Long
    If Lmsg = WM_MOUSEWHEEL Then
        MouseKeys = wParam And 65535
        Rotation = wParam / 65536
        Xpos = lParam And 65535
        Ypos = lParam / 65536
        MouseWheel MouseKeys, Rotation, Xpos, Ypos
    End If
    WindowProc2 = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function
Public Sub WheelHook(PassedForm As Form)
    On Error Resume Next
    Set MyForm = PassedForm
    LocalHwnd = PassedForm.hwnd
    LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc2)
End Sub
Public Sub WheelUnHook()
    Dim WorkFlag As Long
    On Error Resume Next
    WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
    Set MyForm = Nothing
End Sub
