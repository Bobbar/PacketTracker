VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autocomplete"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   Icon            =   "Autocomplete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim textIn As String
Dim selText As Integer
Dim listKey As Boolean


Private Sub Form_Load()

    List1.AddItem "aaa"
    List1.AddItem "aab"
    List1.AddItem "aac"
    List1.AddItem "aba"
    List1.AddItem "abb"
    List1.AddItem "abc"
    List1.AddItem "aca"
    List1.AddItem "acb"
    List1.AddItem "acc"
    List1.AddItem "baa"
    List1.AddItem "bab"
    List1.AddItem "bac"
    List1.AddItem "bba"
    List1.AddItem "bbb"
    List1.AddItem "bbc"
    List1.AddItem "bca"
    List1.AddItem "bcb"
    List1.AddItem "bcc"
    List1.AddItem "caa"
    List1.AddItem "cab"
    List1.AddItem "cac"
    List1.AddItem "cba"
    List1.AddItem "cbb"
    List1.AddItem "cbc"
    List1.AddItem "cca"
    List1.AddItem "ccb"
    List1.AddItem "ccc"
    
   ' List1.Visible = False
   ' List2.Visible = False

End Sub

Private Sub Text1_Change()

    textIn = Text1.Text
    List2.Clear
    List2.Visible = False
    selText = 0
    listKey = True

    If Text1.Text <> "" Then
    
        For i = 0 To List1.ListCount
        
            cad = List1.List(i)
            
            If InStr(1, cad, textIn, vbTextCompare) = 1 Then
                List2.AddItem cad
                List2.Visible = True
            End If
            
            
            If List2.ListCount = 1 Then
                List2.Height = 100
            ElseIf List2.ListCount = 2 Then
                List2.Height = 500
            ElseIf List2.ListCount = 3 Then
                List2.Height = 700
            ElseIf List2.ListCount >= 4 Then
                List2.Height = 1000
            End If
        
        Next
    
    End If

End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38 Then
    
        If selText >= 1 Then
            If listKey = True Then
                selText = selText - 2
            Else
                selText = selText - 1
            End If
            
            List2.Selected(selText) = True
            listKey = False
        End If
        
    End If
    
    If KeyCode = 40 Then
    
        If List2.ListCount >= 0 Then
            If selText < List2.ListCount Then
                If listKey = True Then
                    selText = selText + 1
                Else
                    selText = selText + 2
                End If
                
                List2.Selected(selText - 1) = True
                listKey = True
            End If
        End If
    
    End If
    
    If KeyCode = 13 Then
    
    If List2.Visible = True And List2.Text <> "" Then
        Text1.Text = List2.Text
        List2.Visible = False
        Text1.SelStart = Len(Text1.Text)
    End If
    
    End If
    
    
    If KeyCode = 27 Then
    
        List2.Visible = False
        List2.Clear
    End If

End Sub

Private Sub List2_Click()

    Text1.Text = List2.Text
    List2.Visible = False
    List2.Clear

End Sub
