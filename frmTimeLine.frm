VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimeLine 
   BackColor       =   &H00808080&
   Caption         =   "History Timeline"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTimeLine.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8760
   ScaleWidth      =   13395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   240
      Left            =   5460
      TabIndex        =   3
      Top             =   6540
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8400
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.PictureBox picWindow 
      BorderStyle     =   0  'None
      Height          =   7635
      Left            =   0
      ScaleHeight     =   7635
      ScaleWidth      =   13395
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13395
      Begin VB.VScrollBar VScroll1 
         Height          =   6400
         Left            =   12000
         Max             =   1000
         SmallChange     =   2
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pbDrawArea 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   9855
         Left            =   0
         ScaleHeight     =   9855
         ScaleWidth      =   11895
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   11895
         Begin VB.Timer tmrActionShow 
            Interval        =   10
            Left            =   9240
            Top             =   5880
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   120
            TabIndex        =   6
            Top             =   5840
            Width           =   4215
            Begin VB.CheckBox chkDayLines 
               BackColor       =   &H00808080&
               Caption         =   "Show Day Lines"
               Height          =   195
               Left            =   240
               TabIndex        =   11
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox chkShowAll 
               BackColor       =   &H00808080&
               Caption         =   "Show All Actions"
               Height          =   195
               Left            =   2040
               TabIndex        =   10
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Visual Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   255
               Width           =   1560
            End
            Begin VB.Shape Cls 
               BorderColor     =   &H00C0C0FF&
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00C0C0FF&
               FillStyle       =   0  'Solid
               Height          =   240
               Left            =   1430
               Top             =   250
               Width           =   265
            End
            Begin VB.Shape UFl 
               BorderColor     =   &H00FFC0FF&
               FillColor       =   &H00FFC0FF&
               FillStyle       =   0  'Solid
               Height          =   230
               Left            =   1170
               Top             =   250
               Width           =   262
            End
            Begin VB.Shape Fil 
               BorderColor     =   &H00FFC0C0&
               FillColor       =   &H00FFC0C0&
               FillStyle       =   0  'Solid
               Height          =   230
               Left            =   916
               Top             =   250
               Width           =   262
            End
            Begin VB.Shape Rec 
               BorderColor     =   &H00C0FFFF&
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   230
               Left            =   654
               Top             =   255
               Width           =   262
            End
            Begin VB.Shape Snd 
               BorderColor     =   &H00C0FFC0&
               FillColor       =   &H00C0FFC0&
               FillStyle       =   0  'Solid
               Height          =   230
               Left            =   392
               Top             =   250
               Width           =   265
            End
            Begin VB.Shape Cr 
               BorderColor     =   &H00C0E0FF&
               FillColor       =   &H00C0E0FF&
               FillStyle       =   0  'Solid
               Height          =   230
               Left            =   130
               Top             =   250
               Width           =   262
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   """Mouse-over for action"""
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   1680
               TabIndex        =   9
               Top             =   480
               Width           =   2415
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   " Packet ACTION | Time in State "
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   1800
               TabIndex        =   8
               Top             =   240
               Width           =   2280
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00808080&
               BackStyle       =   1  'Opaque
               FillColor       =   &H000080FF&
               Height          =   255
               Left            =   120
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   0
            Picture         =   "frmTimeLine.frx":08CA
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblPacketAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Packet Age"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5640
            TabIndex        =   4
            Top             =   5880
            Width           =   810
         End
         Begin VB.Line lnScale 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   5
            Visible         =   0   'False
            X1              =   840
            X2              =   12175
            Y1              =   5520
            Y2              =   5520
         End
      End
      Begin ComctlLib.Slider sldEntries 
         Height          =   435
         Left            =   0
         TabIndex        =   13
         ToolTipText     =   "Change "
         Top             =   0
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   767
         _Version        =   327682
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
   End
   Begin VB.CommandButton cmdCantSeeMe 
      Caption         =   "Command1"
      Height          =   195
      Left            =   11760
      TabIndex        =   5
      Top             =   4800
      Width           =   135
   End
End
Attribute VB_Name = "frmTimeLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient _
                Lib "user32" (ByVal hwnd As Long, _
                              lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim ret As Long
Dim a   As POINTAPI
Dim b   As Long
Dim c   As Long

Private Function lngTotaltime(Ents As Integer) As Long
Dim i As Integer
Dim tmpTime As Long

tmpTime = DateDiff("n", TicketDate(0), TicketDate(Ents - 1))

'For i = 0 To Ents
'
'tmpTime = tmpTime + TicketHours(i)
'
'
'
'Next i
lngTotaltime = tmpTime


End Function
Public Sub DrawTimeLine()
    Dim i, Days As Integer
    Dim DStep As Single
    'Dim intLastEntry As Integer
    'intLastEntry = Entry - 20
    On Error Resume Next
    LStep = (frmTimeLine.lnScale.X2 - frmTimeLine.lnScale.X1) / (lngTotaltime(intLastEntry))  ' + TicketHours(Entry - 1))
    frmTimeLine.pbDrawArea.FillColor = &H80C0FF
    ReDim dLine(intLastEntry - 1)
    dLine(0).Color = &H80C0FF
    dLine(0).Height = 300
    dLine(0).Left = 470
    dLine(0).Top = 120
    dLine(0).Width = 315
    ReDim dGrid(intLastEntry - 1)
    dGrid(0).Color = &HE0E0E0
    dGrid(0).Height = 300
    dGrid(0).Left = 0
    dGrid(0).Top = 120
    dGrid(0).Width = 11895
    ReDim dAction(intLastEntry - 1)
    ReDim dNote(intLastEntry - 1)
    dAction(0).Text = TicketActionText(0)
    dAction(0).Color = &H80C0FF
    dAction(0).Left = dLine(0).Left + dLine(0).Width + 200
    dAction(0).Top = dLine(0).Top + 20
    dAction(0).Height = 210
    dAction(0).Visible = True
    dNote(0).Height = 210
    dGrid(0).Width = frmTimeLine.Width
    Days = lngTotaltime(intLastEntry) / 1440 'TotalTime / 1440
    Days = Round(Days, 1)
    For i = 0 To intLastEntry - 1
        With frmTimeLine
            If i Mod 2 <> 0 Then 'number is odd
                dGrid(i).Color = &HC0C0C0
            Else
                dGrid(i).Color = &HE0E0E0
            End If
            dGrid(i).Width = .Width + 200
            dGrid(i).Top = dGrid(i - 1).Top + dGrid(0).Height
            dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width
            dLine(i).Top = dLine(i - 1).Top + dLine(0).Height
            If TicketAction(i) = "CREATED" Then
                dLine(i).Color = colCreate
            ElseIf TicketAction(i) = "INTRANSIT" Then
                dLine(i).Color = colInTransit
            ElseIf TicketAction(i) = "RECEIVED" Then
                dLine(i).Color = colReceived
            ElseIf TicketAction(i) = "NULL" Then
                dLine(i).Color = colClosed
            ElseIf TicketAction(i) = "FILED" Then
                dLine(i).Color = colFiled
            ElseIf TicketAction(i) = "REOPENED" Then
                dLine(i).Color = colReopened
            End If
            If TicketHours(i) * LStep < 38 Then 'Less than 1 pixel wide
                dLine(i).Width = 38
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width - 38
            Else
                dLine(i).Width = TicketHours(i) * LStep
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width
            End If
            If i = intLastEntry - 1 Then
                '                dLine(i).Width = 38
                '                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width - 38
                dLine(i).Width = 200
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width ' - 100
                dLine(i).FillStyle = 7
            End If
            dNote(i).Text = strTimelineComments(i)
            dNote(i).Width = frmTimeLine.pbDrawArea.TextWidth(dNote(i).Text)
            dAction(i).Text = TicketActionText(i)
            dAction(i).Width = frmTimeLine.pbDrawArea.TextWidth(dAction(i).Text)
            If dNote(i).Width > dAction(i).Width Then
                Dim strNote As String
                strNote = dNote(i).Text
                dNote(i).Width = .pbDrawArea.TextWidth(strNote)
                Do Until dNote(i).Width <= dAction(i).Width
                    strNote = Left$(strNote, Len(strNote) - 1)
                    dNote(i).Width = .pbDrawArea.TextWidth(strNote)
                Loop
                dNote(i).Text = Left$(strNote, Len(strNote) - 4) & "..." & Chr$(34)
            End If
            If dLine(i).Left - dAction(i).Width - 240 < 0 And (dLine(i).Left + dLine(i).Width) + dAction(i).Width + 400 < .Width Then
                dAction(i).Left = (dLine(i).Left + dLine(i).Width) + 200
            ElseIf (dLine(i).Left + dLine(i).Width) + dAction(i).Width + 400 > .Width And dLine(i).Left - dAction(i).Width - 240 > 0 Then
                dAction(i).Left = dLine(i).Left - dAction(i).Width - 200
            ElseIf (dLine(i).Left + dLine(i).Width) + dAction(i).Width + 400 > .Width And dLine(i).Left - dAction(i).Width - 240 < 0 Then
                dAction(i).Left = ((dLine(i).Left + dLine(i).Width) / 2) - (dAction(i).Width / 2)  '+  dLine(i).X1
            ElseIf (dLine(i).Left + dLine(i).Width) + dAction(i).Width + 400 < .Width And dLine(i).Left - dAction(i).Width - 240 > 0 Then
                dAction(i).Left = (dLine(i).Left + dLine(i).Width) + 200
            End If
            dAction(i).Top = dGrid(i).Top + dGrid(0).Height / 2 - dAction(0).Height / 2
            dAction(i).Color = dLine(i).Color
            Printer.FontSize = 9
            If frmTimeLine.chkShowAll.Value = 1 Then
                dAction(i).Visible = True
            Else
                dAction(i).Visible = False
            End If
        End With
    Next i
    If Days > 0 Then
        DStep = ((dLine(UBound(dLine)).Left + dLine(UBound(dLine)).Width) - frmTimeLine.lnScale.X1) / Days
    Else
    End If
    ReDim dDayLine(Days)
    dDayLine(0).Y1 = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200
    dDayLine(0).Y2 = dGrid(0).Top
    dDayLine(0).X1 = 470
    dDayLine(0).X2 = 470
    For i = 1 To Days
        '        dDayLine(i).Y1 = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200 'frmTimeLine.lnScale.Y1
        '        dDayLine(i).Y2 = dGrid(0).Top
        '        dDayLine(i).X1 = dDayLine(i - 1).X1 + DStep
        '        dDayLine(i).X2 = dDayLine(i - 1).X2 + DStep
        dDayLine(i).Y1 = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200
        dDayLine(i).Y2 = dGrid(0).Top
        dDayLine(i).X1 = dDayLine(0).X1 + DStep * i
        dDayLine(i).X2 = dDayLine(0).X2 + DStep * i
    Next i
    DrawLines
    Image1.ZOrder 0
    frmTimeLine.lblPacketAge.Top = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200 + 40
    If frmTimeLine.Width <= 10755 Then frmTimeLine.lblPacketAge.Left = frmTimeLine.Frame1.Left + frmTimeLine.Frame1.Width + 10
    frmTimeLine.lblPacketAge.Caption = "Packet Age: " & (IIf((lngTotaltime(intLastEntry) + TicketHours(intLastEntry - 1)) > 1440, Round((lngTotaltime(intLastEntry) + TicketHours(intLastEntry - 1)) / 1440, 1) & "days", Round((lngTotaltime(intLastEntry) + TicketHours(intLastEntry - 1)) / 60, 1) & "hrs"))
    '    If frmTimeLine.lblPacketAge.Top + 30 >= frmTimeLine.Height Then
    '        frmTimeLine.pbDrawArea.Height = frmTimeLine.lblPacketAge.Top + 40
    '    Else
    '        frmTimeLine.pbDrawArea.Height = frmTimeLine.picWindow.Height
    '    End If
    'If frmTimeLine.lblPacketAge.Top + 1100 >= frmTimeLine.Height Then
    frmTimeLine.pbDrawArea.Height = frmTimeLine.lblPacketAge.Top + 2500
    'Else
    '   frmTimeLine.pbDrawArea.Height = frmTimeLine.picWindow.Height
    'End If
    ' If frmTimeLine.Visible = True Then
    '    frmTimeLine.VScroll1.Max = frmTimeLine.VScroll1.Max
    ' Else
    If frmTimeLine.pbDrawArea.ScaleHeight - frmTimeLine.picWindow.ScaleHeight > 0 Then
        frmTimeLine.VScroll1.Max = frmTimeLine.pbDrawArea.ScaleHeight - frmTimeLine.picWindow.ScaleHeight
    Else
        frmTimeLine.VScroll1.Max = 0
    End If
    ' End If
    frmTimeLine.Frame1.Top = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 500
End Sub
Public Sub ReDrawTimeLine()
    Dim i, Days As Integer
    Dim DStep As Single
    'On Error Resume Next
    LStep = (frmTimeLine.lnScale.X2 - frmTimeLine.lnScale.X1) / (lngTotaltime(intLastEntry)) ' + TicketHours(Entry - 1))
    dGrid(0).Width = frmTimeLine.Width + 20
    Days = lngTotaltime(intLastEntry) / 1440 'TotalTime / 1440
    Days = Round(Days, 1)
    For i = 1 To intLastEntry - 1
        With frmTimeLine
            dGrid(i).Width = .Width
            If TicketHours(i) * LStep < 30 Then 'Less than 1 pixel wide
                dLine(i).Width = 30
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width - 30
            Else
                dLine(i).Width = TicketHours(i) * LStep
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width
            End If
            If i = intLastEntry - 1 Then
                dLine(i).Width = 200
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width ' - 100
                dLine(i).FillStyle = 7
            End If
            If chkShowAll.Value = 1 Then
                dAction(i).Top = dGrid(i).Top + dGrid(0).Height / 2 - dAction(0).Height / 2
                If dLine(i).Left - dAction(i).Width - 200 < 0 And (dLine(i).Left + dLine(i).Width) + dAction(i).Width + VScroll1.Width + 200 < pbDrawArea.Width Then
                    dAction(i).Left = (dLine(i).Left + dLine(i).Width) + 200
                ElseIf (dLine(i).Left + dLine(i).Width) + dAction(i).Width + VScroll1.Width > pbDrawArea.Width And dLine(i).Left - dAction(i).Width - 200 > 0 Then
                    dAction(i).Left = dLine(i).Left - dAction(i).Width - 200
                ElseIf (dLine(i).Left + dLine(i).Width) + dAction(i).Width > pbDrawArea.Width And dLine(i).Left - dAction(i).Width - 200 < 0 Then
                    dAction(i).Left = ((dLine(i).Left + dLine(i).Width) / 2) - (dAction(i).Width / 2)  '+  dLine(i).X1
                ElseIf (dLine(i).Left + dLine(i).Width) + dAction(i).Width + VScroll1.Width < pbDrawArea.ScaleWidth And dLine(i).Left - dAction(i).Width - 200 > 0 Then
                    dAction(i).Left = (dLine(i).Left + dLine(i).Width) + 200
                Else
                    'dAction(i).Left = pbDrawArea.Width / 2 - dAction(i).Width / 2
                End If
                If dAction(i).Left < 0 Or (dAction(i).Left + dAction(i).Width) > pbDrawArea.Width Then dAction(i).Left = pbDrawArea.Width / 2 - dAction(i).Width / 2
            Else
            End If
        End With
    Next i
    'Day Lines
    If DrawDayLines = True Then
        If Days > 0 Then
            DStep = ((dLine(UBound(dLine)).Left + dLine(UBound(dLine)).Width) - frmTimeLine.lnScale.X1) / Days
        Else
        End If
        dDayLine(0).Y1 = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200
        dDayLine(0).Y2 = dGrid(0).Top
        dDayLine(0).X1 = 470
        dDayLine(0).X2 = 470
        For i = 1 To Days
            dDayLine(i).Y1 = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200
            dDayLine(i).Y2 = dGrid(0).Top
            dDayLine(i).X1 = dDayLine(0).X1 + DStep * i
            dDayLine(i).X2 = dDayLine(0).X2 + DStep * i
        Next i
    End If
    ' If frmTimeLine.lblPacketAge.Top + 1100 >= frmTimeLine.Height Then
    frmTimeLine.pbDrawArea.Height = frmTimeLine.lblPacketAge.Top + 2500
    'Else
    '  frmTimeLine.pbDrawArea.Height = frmTimeLine.picWindow.Height
    ' End If
    If frmTimeLine.pbDrawArea.ScaleHeight - frmTimeLine.picWindow.ScaleHeight > 0 Then
        frmTimeLine.VScroll1.Max = frmTimeLine.pbDrawArea.ScaleHeight - frmTimeLine.picWindow.ScaleHeight
    Else
        frmTimeLine.VScroll1.Max = 0
    End If
    DrawLines
End Sub
Public Sub DrawLines()
    Dim i As Integer
    With frmTimeLine
        .pbDrawArea.Cls
        For i = 0 To UBound(dGrid) 'draw grid
            .pbDrawArea.DrawStyle = 0
            .pbDrawArea.FillColor = dGrid(i).Color
            .pbDrawArea.Line (dGrid(i).Left, dGrid(i).Top)-(dGrid(i).Left + dGrid(i).Width, dGrid(i).Top + dGrid(0).Height), dGrid(i).Color, B
        Next
        If chkDayLines.Value = 1 Then
            For i = 0 To UBound(dDayLine) 'draw day lines
                .pbDrawArea.DrawStyle = 2
                .pbDrawArea.Line (dDayLine(i).X1, dDayLine(i).Y1)-(dDayLine(i).X2, dDayLine(i).Y2), &H878181 '&H9B9997  '&H404040
            Next
        End If
        .pbDrawArea.DrawStyle = 0
        If dPointLine.Visible Then .pbDrawArea.Line (dPointLine.X1, dPointLine.Y1)-(dPointLine.X2, dPointLine.Y2)
        For i = 0 To UBound(dLine) 'draw bars
            If i = UBound(dLine) Then
                .pbDrawArea.Font.Size = 25
                .pbDrawArea.Font.Name = "Wingdings 3"
                .pbDrawArea.DrawStyle = 0
                .pbDrawArea.FillStyle = dLine(i).FillStyle
                .pbDrawArea.FillColor = dLine(i).Color
                .pbDrawArea.Line (dLine(i).Left, dLine(i).Top)-(dLine(i).Left + dLine(i).Width, dLine(i).Top + dLine(0).Height), vbBlack, B
                .pbDrawArea.CurrentX = dLine(i).Left + dLine(i).Width + 40
                .pbDrawArea.CurrentY = dLine(i).Top - 100
                .pbDrawArea.Print Chr$(52)
            Else
                .pbDrawArea.DrawStyle = 0
                .pbDrawArea.FillStyle = dLine(i).FillStyle
                .pbDrawArea.FillColor = dLine(i).Color
                .pbDrawArea.Line (dLine(i).Left, dLine(i).Top)-(dLine(i).Left + dLine(i).Width, dLine(i).Top + dLine(0).Height), vbBlack, B
            End If
        Next
        .pbDrawArea.FillStyle = 0
        .pbDrawArea.Font.Size = 9
        .pbDrawArea.Font.Name = "Tahoma"
        .pbDrawArea.FontSize = 9
        For i = 0 To UBound(dAction)
            If dAction(i).Visible Then
                .pbDrawArea.DrawStyle = 0
                .pbDrawArea.FillColor = dAction(i).Color
                .pbDrawArea.Line (dAction(i).Left, dAction(i).Top)-(dAction(i).Left + dAction(i).Width, dAction(i).Top + dAction(0).Height), dAction(i).Color, B
                If dNote(i).Visible Then
                    .pbDrawArea.DrawStyle = 0
                    .pbDrawArea.FillColor = dNote(i).Color
                    .pbDrawArea.Line (dNote(i).Left, dNote(i).Top)-(dNote(i).Left + dAction(i).Width, dNote(i).Top + dNote(0).Height), dNote(i).Color, B
                    .pbDrawArea.CurrentX = dNote(i).Left + (dAction(i).Width / 2) - (dNote(i).Width / 2)
                    .pbDrawArea.CurrentY = dNote(i).Top ' + (dAction(0).Height / 2)
                    If chkShowAll.Value = False Then
                        .pbDrawArea.ForeColor = &H80000012
                    Else
                        .pbDrawArea.ForeColor = &H80000012 '&HA4A4A4
                    End If
                    .pbDrawArea.DrawStyle = 0
                    .pbDrawArea.FontTransparent = True
                    .pbDrawArea.Font.Italic = True
                    .pbDrawArea.Print dNote(i).Text
                End If
                .pbDrawArea.CurrentX = dAction(i).Left + (dAction(i).Width / 2) - (dAction(i).Width / 2)
                .pbDrawArea.CurrentY = dAction(i).Top ' + (dAction(0).Height / 2)
                If chkShowAll.Value = False Then
                    .pbDrawArea.ForeColor = &H80000012
                Else
                    .pbDrawArea.ForeColor = &H80000012 '&HA4A4A4
                End If
                .pbDrawArea.DrawStyle = 0
                .pbDrawArea.FontTransparent = True
                .pbDrawArea.Font.Italic = False
                .pbDrawArea.Print dAction(i).Text
            End If
        Next i
        .pbDrawArea.DrawWidth = 5
        .pbDrawArea.ForeColor = vbBlack
        .pbDrawArea.Line (470, dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200)-((dLine(UBound(dLine)).Left + dLine(UBound(dLine)).Width), dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200)
        .pbDrawArea.DrawWidth = 1
    End With
End Sub
Private Sub UnloadControls()
End Sub
Private Sub chkDayLines_Click()
    If chkDayLines.Value = 1 Then
        DrawDayLines = True
        ' DayLine(0).Visible = True
        DrawTimeLine
        ReDrawTimeLine
    Else
        DrawDayLines = False
        'DayLine(0).Visible = False
        DrawTimeLine
        ReDrawTimeLine
    End If
End Sub
Private Sub chkShowAll_Click()
    If chkShowAll.Value = 0 Then
        tmrActionShow.Enabled = True
        DrawTimeLine
        'ReDrawTimeLine
    Else
        tmrActionShow.Enabled = False
        DrawTimeLine
        'ReDrawTimeLine
    End If
End Sub
Private Sub cmdDone_Click()
    tmrActionShow.Enabled = False
    UnloadControls
    Unload Me
End Sub
Private Sub Form_Load()
    picWindow.Left = 0
    picWindow.Width = (frmTimeLine.Width - VScroll1.Width) - 225
    VScroll1.Height = picWindow.Height
    VScroll1.Max = pbDrawArea.Height - picWindow.Height
    VScroll1.SmallChange = 100
    VScroll1.LargeChange = picWindow.Height
    VScroll1.Left = (frmTimeLine.Width - VScroll1.Width) ' - 225
    VScroll1.Top = 0
    pbDrawArea.Left = 0
    pbDrawArea.Width = picWindow.Width
    MouseXPrev = 0
    MouseYPrev = 0
    If Not DrawDayLines Then
    frmTimeLine.chkDayLines.Value = 0
    Else
    frmTimeLine.chkDayLines.Value = 1
    End If
    
End Sub
Private Sub CoordinateMouse()
    On Error Resume Next
    ret = GetCursorPos(a)
    ScreenToClient Me.hwnd, a
    b = a.X * Screen.TwipsPerPixelX
    c = a.Y * Screen.TwipsPerPixelY
    MouseX = b
    MouseY = c + VScroll1.Value - sldEntries.Height
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnloadControls
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    'Me.Height = 7320
    If frmTimeLine.Width < (Frame1.Width + Frame1.Left + 550) Then frmTimeLine.Width = (Frame1.Width + Frame1.Left + 550)
    picWindow.Left = 0
    picWindow.Width = (frmTimeLine.Width) ' - VScroll1.Width)
    picWindow.Height = frmTimeLine.Height - StatusBar1.Height
    sldEntries.Width = frmTimeLine.Width - (VScroll1.Width * 2)
    'pbDrawArea.Height = picWindow.Height
    pbDrawArea.Left = 0
    pbDrawArea.Width = picWindow.Width - VScroll1.Width * 2
    VScroll1.Height = frmTimeLine.Height - StatusBar1.Height - 500
    VScroll1.SmallChange = 2
    VScroll1.LargeChange = picWindow.Height
    VScroll1.Left = picWindow.Width - VScroll1.Width * 2 '(frmTimeLine.Width - VScroll1.Width * 2) '- 100
    VScroll1.Top = 0
    lnScale.X2 = pbDrawArea.Width - 500
    lblPacketAge.Left = (lnScale.X2 / 2) - (lblPacketAge.Width / 2)
    cmdDone.Left = (StatusBar1.Width / 2) - (cmdDone.Width / 2)
    cmdDone.Top = frmTimeLine.Height - cmdDone.Height - 600
    'Form1.DrawTimeLine
    ReDrawTimeLine
    'Me.Refresh
    pbDrawArea.Refresh
    'cmdCantSeeMe.SetFocus
End Sub

Private Sub pbDrawArea_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    If chkShowAll.Value = 0 Then
        Dim intOffset As Integer
        intOffset = 800
        Dim i               As Integer
        Dim intNumofActions As Integer
        'Call CoordinateMouse
        If chkShowAll.Value = False Then
            If MouseX <> MouseXPrev Or MouseY <> MouseYPrev Then
                MouseXPrev = MouseX
                MouseYPrev = MouseY
                For i = 0 To UBound(dGrid)
                    If MouseY > dLine(i).Top And MouseY < (dLine(i).Top + dLine(0).Height) And MouseX > dLine(i).Left - intOffset And MouseX < (dLine(i).Left + dLine(i).Width) + intOffset Then
                        If MouseX + 20 + dAction(i).Width >= frmTimeLine.pbDrawArea.Width Then
                            dAction(i).Left = (MouseX - dAction(i).Width) - 300
                        Else
                            dAction(i).Left = MouseX + 300
                        End If
                        dAction(i).Top = MouseY - dAction(0).Height
                        dAction(i).Visible = True
                        dNote(i).Top = dAction(i).Top + dAction(0).Height
                        dNote(i).Left = dAction(i).Left
                        dNote(i).Color = dAction(i).Color
                        If dNote(i).Text <> "" Then dNote(i).Visible = True
                        dPointLine.X1 = dLine(i).Left + dLine(i).Width ' / 2
                        dPointLine.Y1 = dLine(i).Top + dLine(0).Height / 2
                        dPointLine.X2 = dAction(i).Left + dAction(i).Width / 2  'MouseX
                        dPointLine.Y2 = dAction(i).Top + dAction(0).Height / 2 'MouseY
                    Else
                        dAction(i).Visible = False
                        dNote(i).Visible = False
                        dPointLine.Visible = False
                        intNumofActions = intNumofActions + 1
                    End If
                Next i
                If intNumofActions > UBound(dGrid) Then 'if no actions are visible, hide pointer line.
                    dPointLine.Visible = False
                Else
                    dPointLine.Visible = True
                End If
                DrawLines
            End If
        Else
            For i = 0 To UBound(dGrid)
                dAction(i).Visible = True
            Next i
            'tmrActionShow.Enabled = False
        End If
    End If
End Sub

Private Sub sldEntries_Scroll()
intLastEntry = sldEntries.Value


DrawTimeLine

'ReDrawTimeLine

End Sub

Private Sub tmrActionShow_Timer()
Call CoordinateMouse
End Sub

Private Sub VScroll1_Change()
    pbDrawArea.Top = -(VScroll1.Value) + sldEntries.Height
    pbDrawArea.Refresh
    cmdCantSeeMe.SetFocus
End Sub
Private Sub VScroll1_Scroll()
    pbDrawArea.Top = -(VScroll1.Value) + sldEntries.Height
    pbDrawArea.Refresh
End Sub
