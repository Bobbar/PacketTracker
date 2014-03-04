VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAttachments 
   Caption         =   "Attachments"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridFileList 
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   540
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8916
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadAttachments()

End Sub
