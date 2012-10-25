VERSION 5.00
Begin VB.Form Opening 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   5715
      ItemData        =   "Opening.frx":0000
      Left            =   120
      List            =   "Opening.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "Opening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim X(9) As String

X(0) = "11/10/2004"
X(1) = "12/13/1995"
X(2) = "01/10/1992"
X(3) = "10/05/2005"
X(4) = "06/15/1996"
X(5) = "09/04/2002"
X(6) = "04/09/1993"
X(7) = "03/05/1995"
X(8) = "12/12/1992"
X(9) = "11/10/2001"


Call B2S_BSort_Date(X(), List1)
MsgBox "Done"
End Sub
