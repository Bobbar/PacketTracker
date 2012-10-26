VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGrid 
   Caption         =   "Form2"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrid.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9135
   ScaleWidth      =   14280
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmGrid.frx":0CCA
      Top             =   480
      Visible         =   0   'False
      Width           =   8025
   End
   Begin VB.Timer tmrGridResize 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   13920
      Top             =   2640
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGrid 
      Height          =   8415
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   14843
      _Version        =   393216
      FixedRows       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLinesUnpopulated=   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   8835
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const EM_GETLINECOUNT = &HBA
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
Private LeftOffset, TopOffset As Integer
Sub FlexSort(Mode As String)
    If FlexGrid.MouseRow = 0 And Mode = "A" Then
        FlexGrid.col = FlexGrid.MouseCol
        If FlexGrid.col = 10 Then
            FlexGrid.Sort = flexSortGenericAscending
        Else
            FlexGrid.Sort = flexSortStringAscending
        End If
    Else
        'do nothing
    End If
    If FlexGrid.MouseRow = 0 And Mode = "D" Then
        FlexGrid.col = FlexGrid.MouseCol
        If FlexGrid.col = 10 Then
            FlexGrid.Sort = flexSortGenericDescending
        Else
            FlexGrid.Sort = flexSortStringDescending
        End If
    Else
        'do nothing
    End If
End Sub
Private Sub FlexGrid_Click()
    On Error Resume Next
    Set WhichGrid = frmGrid.FlexGrid
    If bolNewHistWindow = True Then Exit Sub
    If strSortMode = "A" Then
        Call FlexSort("D")
        strSortMode = "D"
    ElseIf strSortMode = "D" Then
        Call FlexSort("A")
        strSortMode = "A"
    End If
End Sub
Private Sub FlexGrid_DblClick()
    On Error Resume Next
    If bolNewHistWindow = True Then Exit Sub
    Screen.MousePointer = vbHourglass
    DoEvents
    'ClearFields
    Form1.OpenPacket FlexGrid.TextMatrix(FlexGrid.RowSel, 1)
    Form1.SSTab1.Tab = 0
    'Form1.Show
    Form1.tmrRefresher.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
Private Sub FlexGrid_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Form1.OpenPacket FlexGrid.TextMatrix(FlexGrid.RowSel, 1)
        Form1.SSTab1.Tab = 0
    End If
End Sub
Private Sub Form_Load()
    If bolHook Then Call WheelHook(frmGrid)
    LeftOffset = frmGrid.Width - (FlexGrid.Left + FlexGrid.Width)
    TopOffset = frmGrid.Height - (FlexGrid.Top + FlexGrid.Height)
End Sub
Public Sub FlexGridRedrawHeight()
    Dim ColLoop As Long
    Dim RowLoop As Long
    'Turn off redrawing to avoid flickering
    FlexGrid.Redraw = False
    'For ColLoop = 0 To FlexGridHist.Cols - 1
    'FlexGridHist.ColWidth(ColLoop) = 2500
    For RowLoop = 0 To FlexGrid.Rows - 1
        ReSizeCellHeight RowLoop, 1
    Next RowLoop
    'Next ColLoop
    'Turn redrawing back on
    FlexGrid.Redraw = True
End Sub
Public Sub ReSizeCellHeight(MyRow As Long, MyCol As Long)
    Dim LinesOfText  As Long
    Dim HeightOfLine As Long
    On Error Resume Next
    'Set MSFlexGrid to appropriate Cell
    FlexGrid.Row = MyRow
    FlexGrid.col = MyCol
    'Set textbox width to match current width of selected cell
    Text1.Width = FlexGrid.ColWidth(MyCol)
    Text1.FontSize = FlexGrid.CellFontSize
    Text1.FontBold = FlexGrid.CellFontBold
    Text1.FontItalic = FlexGrid.CellFontItalic
    Text1.Text = FlexGrid.Text
    'Get the height of the text in the textbox
    HeightOfLine = 285 'Me.TextHeight(Text1.Text) '285
    'Call API to determine how many lines of text are in text box
    LinesOfText = SendMessage(Text1.hwnd, EM_GETLINECOUNT, 0&, 0&)
    'Check to see if row is not tall enough
    ' If FlexGrid.RowHeight(MyRow) < (LinesOfText * HeightOfLine) Then
    'Adjust the RowHeight based on the number of lines in textbox
    FlexGrid.RowHeight(MyRow) = LinesOfText * HeightOfLine + 200
    ' End If
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    FlexGrid.Width = frmGrid.Width - LeftOffset
    FlexGrid.Height = frmGrid.Height - TopOffset
    If bolNewHistWindow = True Then
        Dim ColW, i As Integer
        For i = 0 To FlexGrid.Cols - 1
            ColW = ColW + FlexGrid.ColWidth(i)
        Next i
        If FlexGrid.Width < ColW + 600 Then
            FlexGrid.ColWidth(1) = FlexGrid.Width - 1500
            FlexGridRedrawHeight
        Else
            'If FlexGrid.ColWidth(1) >= 12615 Then
            ' FlexGrid.ColWidth(1) = 12615
            ' Else
            FlexGrid.ColWidth(1) = FlexGrid.Width - 1500
            FlexGridRedrawHeight
            ' End If
        End If
    End If
End Sub
Private Sub tmrGridResize_Timer()
    Dim ColW, i As Integer
    For i = 0 To FlexGrid.Cols - 1
        ColW = ColW + FlexGrid.ColWidth(i)
    Next i
    If FlexGrid.Width < ColW + 600 Then
        FlexGrid.ColWidth(1) = FlexGrid.Width - 1200
        FlexGridRedrawHeight
    Else
        If FlexGrid.ColWidth(1) >= 12615 Then
            FlexGrid.ColWidth(1) = 12615
        Else
            FlexGrid.ColWidth(1) = FlexGrid.Width - 1200
            FlexGridRedrawHeight
        End If
    End If
End Sub
