VERSION 5.00
Begin VB.Form fTestIni 
   Caption         =   "cINIFile Class Demonstration Application"
   ClientHeight    =   5670
   ClientLeft      =   3645
   ClientTop       =   2310
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fTestIni.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   7200
   Begin VB.PictureBox pnlButtons 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   5340
      ScaleHeight     =   2895
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   180
      Width           =   1875
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Re-&Create Test Ini"
         Height          =   430
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1800
      End
      Begin VB.CommandButton cmdEnumSection 
         Caption         =   "&Enumerate Section...."
         Height          =   430
         Left            =   0
         TabIndex        =   15
         Top             =   480
         Width           =   1800
      End
      Begin VB.CommandButton cmdGetValue 
         Caption         =   "&Get Value for Key..."
         Height          =   430
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   1800
      End
      Begin VB.CommandButton cmdDeleteKey 
         Caption         =   "&Delete Key..."
         Height          =   430
         Left            =   0
         TabIndex        =   13
         Top             =   1920
         Width           =   1800
      End
      Begin VB.CommandButton cmdDeleteSection 
         Caption         =   "Delete &Section..."
         Height          =   430
         Left            =   0
         TabIndex        =   12
         Top             =   2400
         Width           =   1800
      End
      Begin VB.CommandButton cmdSetValue 
         Caption         =   "Set &Value for Key..."
         Height          =   430
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   1800
      End
   End
   Begin VB.ListBox lstIni 
      Height          =   5325
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   5115
   End
   Begin VB.Frame fraInfo 
      Caption         =   "cIniFile Parameters:"
      Height          =   2535
      Left            =   5340
      TabIndex        =   0
      Top             =   3060
      Width           =   1815
      Begin VB.TextBox txtInfo 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   2100
         Width           =   1635
      End
      Begin VB.TextBox txtInfo 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txtInfo 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox txtInfo 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label lblInfo 
         Caption         =   "Value:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1860
         Width           =   1995
      End
      Begin VB.Label lblInfo 
         Caption         =   "Key:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label lblInfo 
         Caption         =   "Section:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   1995
      End
      Begin VB.Label lblInfo 
         Caption         =   "Path:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   1695
      End
   End
End
Attribute VB_Name = "fTestIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cIni As New cInifile

Private Sub ShowIniAndParameters()
Dim sGet As String
Dim sSections() As String
Dim iSectionCount As Long
Dim sKeys() As String
Dim iKeycount As Long
Dim iSection As Long
Dim iKey As Long
Dim lSect As Long

    lstIni.Clear
    With m_cIni
        .EnumerateAllSections sSections(), iSectionCount
        For iSection = 1 To iSectionCount
            lstIni.AddItem "[" & sSections(iSection) & "]"
            lSect = lstIni.NewIndex
            lstIni.ItemData(lSect) = -1
            .Section = sSections(iSection)
            .EnumerateCurrentSection sKeys(), iKeycount
            For iKey = 1 To iKeycount
                .Key = sKeys(iKey)
                lstIni.AddItem .Key & "=" & .Value
                lstIni.ItemData(lstIni.NewIndex) = lSect
            Next iKey
        Next iSection
        If (lstIni.ListCount > 0) Then
            lstIni.ListIndex = lstIni.ListCount - 1
        End If
        txtInfo(0) = .Path
        txtInfo(1) = .Section
        txtInfo(2) = .Key
        txtInfo(3) = .Value
    End With
End Sub
Private Sub CreateTestIni()
Dim i As Long
Dim iNd As Long
    ' Create an Ini File to play with:
    On Error Resume Next
        Kill App.Path & "\TEST.INI"
    On Error GoTo 0
    With m_cIni
        .Path = App.Path & "\TEST.INI"
        .Section = "Window"
        .Key = "Toolbar"
        .Value = "1"
        .Key = "Statusbar"
        .Value = "1"
        .Key = "Maximised"
        .Value = (Me.WindowState = vbMaximized)
        .Key = "Left"
        .Value = Me.Left
        .Key = "Top"
        .Value = Me.Top
        .Key = "Width"
        .Value = Me.Width
        .Key = "Height"
        .Value = Me.Height
        .Key = "Title"
        .Value = App.Title
        .Section = "Options"
        .Key = "ShowTips"
        .Value = "1"
        .Key = "OpenInNewWindow"
        .Value = "0"
        .Section = "Controls"
        For i = 0 To Me.Controls.Count - 1
            .Key = "Control" & i
            Err.Clear
            On Error Resume Next
            iNd = Me.Controls(i).Index
            If (Err.Number = 0) Then
                .Value = Me.Controls(i).Name & "," & CStr(iNd)
            Else
                .Value = Me.Controls(i).Name
            End If
        Next i
    End With
    ShowIniAndParameters

End Sub

Private Sub cmdCreate_Click()
    CreateTestIni
End Sub

Private Sub cmdDeleteKey_Click()
    With m_cIni
        .Path = txtInfo(0)
        .Section = txtInfo(1)
        .Key = txtInfo(2)
        .DeleteKey
        If Not (.Success) Then
            MsgBox "Delete Key Failed.", vbInformation
        End If
    End With
    ShowIniAndParameters
End Sub

Private Sub cmdDeleteSection_Click()
    With m_cIni
        .Path = txtInfo(0)
        .Section = txtInfo(1)
        .DeleteSection
        If Not (.Success) Then
            MsgBox "Delete Section Failed.", vbInformation
        End If
    End With
    ShowIniAndParameters
End Sub

Private Sub cmdEnumSection_Click()
Dim sKey() As String, iCount As Long, i As Long
Dim sOut As String
    With m_cIni
        .Path = txtInfo(0)
        .Section = txtInfo(1)
        .EnumerateCurrentSection sKey(), iCount
        If (iCount > 0) Then
            For i = 1 To iCount
                sOut = sOut & vbCrLf & "    " & sKey(i)
            Next i
            MsgBox "Section contains:" & sOut, vbInformation
        Else
            MsgBox "Section is empty.", vbInformation
        End If
    End With
End Sub


Private Sub cmdGetValue_Click()
    With m_cIni
        .Path = txtInfo(0)
        .Section = txtInfo(1)
        .Key = txtInfo(2)
        .Default = "THIS IS THE DEFAULT"
        txtInfo(3) = .Value
        If Not (.Success) Then
            MsgBox "Failed to get value.", vbInformation
        End If
    End With

End Sub

Private Sub cmdSetValue_Click()
    With m_cIni
        .Path = txtInfo(0)
        .Section = txtInfo(1)
        .Key = txtInfo(2)
        .Value = txtInfo(3)
        If Not (.Success) Then
            MsgBox "Failed to set value.", vbInformation
        End If
    End With
    ShowIniAndParameters
End Sub

Private Sub Form_Load()
    With m_cIni
        .Path = App.Path & "\PTESTINI.INI"
        .Section = Me.Name
        .LoadFormPosition Me, Me.Width, Me.Height
    End With
    CreateTestIni
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With m_cIni
        .Path = App.Path & "\PTESTINI.INI"
        .Section = Me.Name
        .SaveFormPosition Me
    End With
    
End Sub

Private Sub Form_Resize()
Dim lL As Long
    On Error Resume Next
    lL = Me.ScaleWidth - fraInfo.Width - 2 * Screen.TwipsPerPixelX
    lstIni.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, lL - 2 * Screen.TwipsPerPixelY, Me.ScaleHeight - 4 * Screen.TwipsPerPixelY
    pnlButtons.Move lL, lstIni.Top
    fraInfo.Move lL
End Sub

Private Sub lstIni_Click()
Dim sItem As String
Dim iPos As Long
    If (lstIni.ListIndex > -1) Then
        If (lstIni.ItemData(lstIni.ListIndex) = -1) Then
            sItem = lstIni.List(lstIni.ListIndex)
            ' Key:
            txtInfo(1) = Mid$(sItem, 2, Len(sItem) - 2)
            txtInfo(2) = ""
            cmdGetValue.Enabled = False
            cmdSetValue.Enabled = False
            cmdDeleteKey.Enabled = False
        Else
            sItem = lstIni.List(lstIni.ItemData(lstIni.ListIndex))
            txtInfo(1) = Mid$(sItem, 2, Len(sItem) - 2)
            sItem = lstIni.List(lstIni.ListIndex)
            iPos = InStr(sItem, "=")
            txtInfo(2) = Left$(sItem, (iPos - 1))
            txtInfo(3) = Mid$(sItem, (iPos + 1))
            cmdGetValue.Enabled = True
            cmdSetValue.Enabled = True
            cmdDeleteKey.Enabled = True
        End If
    End If
End Sub
