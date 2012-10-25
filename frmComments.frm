VERSION 5.00
Begin VB.Form frmComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notes"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComments.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblChars 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 / 200"
      Height          =   195
      Left            =   4920
      TabIndex        =   2
      Top             =   1680
      Width           =   510
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    txtComment.Text = ""

End Sub

Private Sub cmdOK_Click()
    frmComments.Hide

End Sub

Private Sub Form_Activate()
    frmComments.txtComment.SetFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True

    Me.Hide

End Sub

Private Sub txtComment_Change()
    strTicketComment = Replace(txtComment.Text, vbCrLf, " ")

    lblChars.Caption = Len(txtComment.Text) & " / 200"

End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then frmComments.Hide

End Sub
