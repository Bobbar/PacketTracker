VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim blah As New CrystalDesignerPluginLib.Designer


Dim crxApplication As New CRAXDRT.Application
    Dim crxReport As CRAXDRT.Report
    Dim crxConnectionInfo As CRAXDRT.ConnectionProperties
    
    Set crxReport = crxApplication.OpenReport(sReportName, 1)
    
    For iLoop = 1 To crxReport.Database.Tables.Count - 1
        With crxReport.Database.Tables(iLoop)
            .Location = sDatabaseName
        End With
    Next
    
    crxViewer.ReportSource = crxReport
    crxViewer.ViewReport
End Sub
