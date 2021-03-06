VERSION 5.00
Begin VB.Form frmRedirect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Redirect Packet"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRedirect.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGo 
      Caption         =   "Redirect"
      Height          =   420
      Left            =   4440
      TabIndex        =   14
      Top             =   1620
      Width           =   1170
   End
   Begin VB.ComboBox cmbOwner 
      Height          =   315
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtOwner 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      TabIndex        =   8
      Text            =   "Owner"
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox cmbUserFrom 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox cmbUserTo 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox cmbAction 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtUserFrom 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Text            =   "UserFrom"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtUserTo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Text            =   "UserTo"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtAction 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "Action"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packet Owner:"
      Height          =   195
      Left            =   7800
      TabIndex        =   13
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User From:"
      Height          =   195
      Left            =   5880
      TabIndex        =   12
      Top             =   240
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User To:"
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Action:"
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1320
   End
End
Attribute VB_Name = "frmRedirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strUserFrom As String
Private strUserTo   As String
Public Sub GetPacket()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    On Error Resume Next
    Form1.ShowData
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.packetentrydb LEFT JOIN (ticketdb.packetlist) ON (packetlist.idJobNum=packetentrydb.idJobNum) WHERE packetlist.idJobNum = '" & Form1.txtJobNo.Text & "' ORDER BY packetentrydb.idDate DESC"
    Set rs = cn_global.Execute(strSQL1)
    With rs
        txtAction.Text = !idAction
        txtUserTo.Text = !idUserTo
        txtUserFrom.Text = !idUserFrom
        txtOwner.Text = !idUser
    End With
    Form1.HideData
End Sub
Private Sub cmbAction_Click()
    If cmbAction.ListIndex = 2 Then 'Received
        cmbUserTo.Enabled = False
        strUserTo = "NULL"
    ElseIf cmbAction.ListIndex = 3 Then 'Filed
        cmbUserTo.Enabled = False
        strUserTo = "NULL"
        cmbUserFrom.Enabled = False
        strUserFrom = "NULL"
    ElseIf cmbAction.ListIndex = 4 Then 'Closed
        cmbUserTo.Enabled = False
        strUserTo = "NULL"
        cmbUserFrom.Enabled = False
        strUserFrom = "NULL"
    ElseIf cmbAction.ListIndex = 5 Then 'Reopened
        cmbUserTo.Enabled = False
        strUserTo = "NULL"
        cmbUserFrom.Enabled = False
        strUserFrom = "NULL"
    Else
        cmbUserTo.Enabled = True
        strUserTo = ""
        cmbUserFrom.Enabled = True
        strUserFrom = "NULL"
    End If
End Sub
Private Sub cmbUserFrom_Change()
    strUserFrom = UCase$(strUserIndex(cmbUserFrom.ListIndex))
End Sub
Private Sub cmbUserFrom_Click()
    strUserFrom = UCase$(strUserIndex(0, cmbUserFrom.ListIndex))
End Sub
Private Sub cmbUserTo_Change()
    strUserTo = UCase$(strUserIndex(cmbUserTo.ListIndex))
End Sub
Private Sub cmbUserTo_Click()
    strUserTo = UCase$(strUserIndex(0, cmbUserTo.ListIndex))
End Sub
Private Sub cmdGo_Click()
    Dim rs             As New ADODB.Recordset
    Dim strSQL1        As String
    Dim FormatDateTime As String
    On Error Resume Next
    FormatDateTime = Format$(Form1.txtCreateDate.Text, strDBDateTimeFormat)
    Form1.ShowData
    cn_global.CursorLocation = adUseClient
    strSQL1 = "INSERT INTO packetentrydb" & " (idUser,idAction,idUserFrom,idUserTo,idComment,idJobNum) VALUES" & " ('" & UCase$(strUserIndex(0, cmbOwner.ListIndex)) & "','" & (IIf(cmbAction.Text = "CLOSED", "NULL", cmbAction.Text)) & "','" & strUserFrom & "','" & strUserTo & "','" & "Packet redirected by " & strLocalUser & "','" & Form1.txtJobNo.Text & "')"
    Set rs = cn_global.Execute(strSQL1)
    Form1.HideData
    ShowBanner colInTransit, "Packet updated successfully."
    Form1.RefreshAll
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
    Form1.imgComment.Picture = ButtonPics(4)
    Form1.imgComment.Enabled = False
    frmRedirect.Hide
End Sub
Private Sub Form_Load()
    cmbAction.Clear
    cmbAction.AddItem "", 0
    cmbAction.AddItem "INTRANSIT", 1
    cmbAction.AddItem "RECEIVED", 2
    cmbAction.AddItem "FILED", 3
    cmbAction.AddItem "REOPENED", 4
End Sub

