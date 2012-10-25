VERSION 5.00
Begin VB.Form frmUserSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Faux Local User"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserSelect.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbUsers 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "frmUserSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbUsers_Click()

    If strUserIndex(0, cmbUsers.ListIndex) <> "" Then

        strLocalUser = UCase$(strUserIndex(0, cmbUsers.ListIndex))
        Form1.txtLocalUser.Enabled = True
        Form1.txtLocalUser.Locked = True

        Form1.txtLocalUser.BackColor = &HFF&
        Form1.txtLocalUser.Text = strLocalUser
        'Form1.lblFauxUser.Visible = True

        Form1.GetMyPackets
        Form1.SetControls

        frmUserSelect.Hide

        Form1.mnuFauxUser.Checked = True

        ShowBanner colClosed, "Faux user set to " & strLocalUser

    End If

End Sub

