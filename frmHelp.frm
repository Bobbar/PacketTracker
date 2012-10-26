VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timeline Help"
   ClientHeight    =   7710
   ClientLeft      =   2295
   ClientTop       =   2325
   ClientWidth     =   11820
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show at Startup"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Page"
      Height          =   375
      Left            =   10440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7020
      Left            =   120
      ScaleHeight     =   6960
      ScaleWidth      =   10140
      TabIndex        =   1
      Top             =   120
      Width           =   10200
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Index in collection of tip currently being displayed.
Dim CurrentTip As Long
Private Sub DoNextTip()
    CurrentTip = CurrentTip + 1
    If CurrentTip > UBound(HelpPics) Then
        HelpClosed = True
        Unload Me
    Else
        Picture1.Picture = HelpPics(CurrentTip)
    End If
End Sub
Private Sub chkLoadTipsAtStartup_Click()
    ' save whether or not this form should be displayed at startup
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub
Private Sub cmdNextTip_Click()
    DoNextTip
End Sub
Private Sub cmdOK_Click()
    HelpClosed = True
    Unload Me
End Sub
Private Sub Form_Load()
    Dim ShowAtStartup As Long
    HelpClosed = False
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    If ShowAtStartup = 0 Then
        HelpClosed = True
        Unload Me
        Exit Sub
    End If
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkLoadTipsAtStartup.Value = vbChecked
    Picture1.Picture = HelpPics(1)
    CurrentTip = 1
End Sub
