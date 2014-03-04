Attribute VB_Name = "RFQMod"
Private Const lngRFQStart As Long = 130000
Private Const lngRFQEnd As Long = 140000

Public Function FindFreeRFQNum() As String
    Dim arrRFQList As Variant
    Dim i          As Long
    Dim TryRFQNum  As Long
    FindFreeRFQNum = ""
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT idRFQNum FROM rfqmain ORDER BY idRFQNum DESC"
    Set rs = cn_global.Execute(strSQL1)
    ReDim arrRFQList(rs.RecordCount)
    With rs
        Do Until .EOF
            arrRFQList(.AbsolutePosition - 1) = !idRFQNum
            .MoveNext
        Loop
        .Close
    End With
    TryRFQNum = lngRFQStart
    Do
        If Not InArray(TryRFQNum, arrRFQList) Then Exit Do
        TryRFQNum = TryRFQNum + 1
    Loop Until TryRFQNum >= lngRFQEnd
    FindFreeRFQNum = TryRFQNum
End Function
Private Function InArray(FindValue As Variant, Arr As Variant) As Boolean
    Dim i As Long
    InArray = False
    For i = 0 To UBound(Arr)
        If Str(FindValue) = Str(Arr(i)) Then
            InArray = True
            Exit Function
        End If
    Next i
End Function
Public Function GetAttachmentList(strJobNum As String, Grid As MSHFlexGrid)
    Dim rs As New ADODB.Recordset
    cn_global.CursorLocation = adUseClient
    Dim strSQL1 As String
    strSQL1 = "SELECT idFilename, idFileType, idFileSize, idDateStamp,idGUID FROM attachments WHERE idJobNum = '" & strJobNum & "' order by idDateStamp DESC"
    Set rs = cn_global.Execute(strSQL1)
    If rs.RecordCount = 0 Then
    Grid.Clear
    Grid.Visible = False
    Form1.SSTab1.TabCaption(1) = "Attachments"
    Exit Function
    End If
    
    Grid.Cols = 6
    Grid.Rows = rs.RecordCount + 1
    
    Grid.TextMatrix(0, 1) = "Filename"
     Grid.TextMatrix(0, 2) = "Size"
     Grid.TextMatrix(0, 3) = "Type"
      Grid.TextMatrix(0, 4) = "Date/Time"
     Grid.TextMatrix(0, 5) = "GUID"
    
    
    With rs
        Do Until .EOF
            Grid.TextMatrix(.AbsolutePosition, 1) = !idFileName & "." & !idFileType
            Grid.TextMatrix(.AbsolutePosition, 2) = Round((!idFileSize / 1024), 2) & " KB"
            Grid.TextMatrix(.AbsolutePosition, 3) = !idFileType
            Grid.TextMatrix(.AbsolutePosition, 4) = !idDateStamp
            Grid.TextMatrix(.AbsolutePosition, 5) = !idGUID
            .MoveNext
        Loop
    End With
    Form1.SSTab1.TabCaption(1) = "Attachments (" & rs.RecordCount & ")"
    Form1.SizeTheSheet Grid
    Grid.Visible = True
    
End Function
Public Function LoadAttachment(strGUID As String)
    On Error GoTo errs
    Dim strFullFileName As String
    Dim strSQL1         As String
    Dim rs              As New ADODB.Recordset
    Form1.ShowData
    frmWait.Show
    DoEvents
    If Dir$(strTempFileLoc, vbDirectory) = "" Then MkDir strTempFileLoc
    Set rs = New ADODB.Recordset
    cn_global.CursorLocation = adUseClient
    strSQL1 = "Select * from attachments where idGUID = '" & strGUID & "'"
    Set rs = cn_global.Execute(strSQL1)
    ' On Error GoTo procNoPicture
    'If Recordset is Empty, Then Exit
    '    If RS Is Nothing Then
    '        GoTo procNoPicture
    '    End If
    Set strStream = New ADODB.Stream
    With rs
        strStream.Type = adTypeBinary
        strStream.Open
        strStream.Write rs.Fields("idFile").Value
        strFullFileName = strTempFileLoc & !idFileName & "." & !idFileType
        strStream.SaveToFile strFullFileName, adSaveCreateOverWrite
        Form1.HideData
        frmWait.Hide
        DoEvents
        Debug.Print ShellExecute(Form1.hwnd, "open", strFullFileName, vbNullString, vbNullString, 4) 'SW_SHOWNORMAL
    
        'strStream.Close
        
    End With
    Set rs = Nothing
    Set strStream = Nothing
    'LoadAttachment = True
    Exit Function
errs:
 Form1.HideData
        frmWait.Hide
        DoEvents
    ErrHandle Err.Number, Err.Description, "LoadAttachment"
End Function
Public Function SaveAttachment(sFileName As String, strFileTitle As String, strJobNum As String)
    'On Error GoTo procNoPicture
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    Dim FileExtSplit() As String
    Dim strSQL1        As String
    'Set rs = New ADODB.Recordset
    Form1.ShowData
    frmWait.Show
    DoEvents
    cn_global.CursorLocation = adUseClient
    mystream.Open

    mystream.LoadFromFile sFileName
    If mystream.Size > lngMaxFileSize Then
    
     Form1.HideData
    frmWait.Hide
    DoEvents
    Dim blah
    blah = MsgBox("File is too large." & vbCrLf & vbCrLf & "Max size is " & Round((lngMaxFileSize / 1024), 2) & " KB the file is " & Round((mystream.Size / 1024), 2) & " KB.", vbExclamation + vbOKOnly, "We're gonna need a bigger hard-drive...")
    Set mystream = Nothing
    Set cmd = Nothing
    
    Exit Function
    End If
    
    
    FileExtSplit = Split(strFileTitle, ".")
    strSQL1 = "INSERT INTO attachments (idFile, idFolder, idFileName, idFileType,idFileSize, idJobNum) VALUES (?,?,?,?,?,?)"
    cmd.ActiveConnection = cn_global
    cmd.CommandText = strSQL1
    cmd.Parameters.Append cmd.CreateParameter("@idFile", adVarBinary, adParamInput, mystream.Size, mystream.Read)
    cmd.Parameters.Append cmd.CreateParameter("@idFolder", adBSTR, adParamInput, , "ROOT")
    cmd.Parameters.Append cmd.CreateParameter("@idFileName", adBSTR, adParamInput, , FileExtSplit(0))
    cmd.Parameters.Append cmd.CreateParameter("@idFileType", adBSTR, adParamInput, , FileExtSplit((UBound(FileExtSplit))))
    cmd.Parameters.Append cmd.CreateParameter("@idFileType", adBigInt, adParamInput, , mystream.Size)
    cmd.Parameters.Append cmd.CreateParameter("@idFileName", adBSTR, adParamInput, , strJobNum)
    cmd.CommandType = adCmdText
    cmd.Execute
    mystream.Close
    Set mystream = Nothing
    Set cmd = Nothing
    
    SaveAttachment = True
    GetAttachmentList strCurJobNum, Form1.FlexAttach
procExitSub:
    Form1.HideData
    frmWait.Hide
    DoEvents
    
    Exit Function
procNoPicture:
  ErrHandle Err.Number, Err.Description, "SaveAttachment"
  
    SaveAttachment = False
    GoTo procExitSub
End Function
