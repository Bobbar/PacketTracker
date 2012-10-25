VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimeLine 
   BackColor       =   &H00808080&
   Caption         =   "History Timeline"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13395
   DrawMode        =   14  'Copy Pen
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
      Left            =   5400
      TabIndex        =   5
      Top             =   6450
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   11895
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   11895
      Begin VB.PictureBox pbDrawArea 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   9855
         Left            =   0
         ScaleHeight     =   9855
         ScaleWidth      =   11895
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
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
            TabIndex        =   8
            Top             =   5800
            Width           =   4215
            Begin VB.CheckBox chkDayLines 
               BackColor       =   &H00808080&
               Caption         =   "Show Day Lines"
               Height          =   195
               Left            =   240
               TabIndex        =   13
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox chkShowAll 
               BackColor       =   &H00808080&
               Caption         =   "Show All Actions"
               Height          =   195
               Left            =   2040
               TabIndex        =   12
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   """Mouse-over for action"""
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   1680
               TabIndex        =   11
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
               TabIndex        =   10
               Top             =   240
               Width           =   2280
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
               Left            =   75
               TabIndex        =   9
               Top             =   255
               Width           =   1620
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00808080&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   120
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Label lblNote 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Note"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Index           =   0
            Left            =   1470
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   615
            WordWrap        =   -1  'True
         End
         Begin VB.Line lnPoint 
            Visible         =   0   'False
            X1              =   720
            X2              =   1560
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line lnVisScale 
            BorderWidth     =   5
            X1              =   470
            X2              =   11805
            Y1              =   5760
            Y2              =   5760
         End
         Begin VB.Image Image1 
            Height          =   300
            Left            =   0
            Picture         =   "frmTimeLine.frx":0CCA
            Stretch         =   -1  'True
            Top             =   120
            Width           =   445
         End
         Begin VB.Label lblPacketAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Packet Age"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5640
            TabIndex        =   6
            Top             =   5880
            Width           =   810
         End
         Begin VB.Line lnScale 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   5
            Visible         =   0   'False
            X1              =   470
            X2              =   11805
            Y1              =   5160
            Y2              =   5160
         End
         Begin VB.Label lblAction 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Action"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Index           =   0
            Left            =   1500
            TabIndex        =   3
            Top             =   150
            Width           =   555
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6400
      Left            =   12000
      Max             =   1000
      SmallChange     =   2
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdCantSeeMe 
      Caption         =   "Command1"
      Height          =   195
      Left            =   11760
      TabIndex        =   7
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

Public Sub ReDrawTimeLine()

    Dim i, Days As Integer

    Dim DStep As Single

    On Error Resume Next
   
    LStep = (frmTimeLine.lnScale.X2 - frmTimeLine.lnScale.X1) / (TotalTime + TicketHours(Entry - 1))
 
    
    dGrid(0).Width = frmTimeLine.Width + 20
    
    Days = (TotalTime + TicketHours(Entry - 1)) / 1440
    Days = Round(Days, 1)
    
    For i = 0 To Entry ' - 1

        With frmTimeLine
           
            
            dGrid(i).Width = .Width
            If TicketHours(i) * LStep < 38 Then 'Less than 1 pixel wide
                
           
                dLine(i).Width = 38
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width - 38
                
            Else
                
                
                dLine(i).Width = TicketHours(i) * LStep
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width
                
            End If

            If chkShowAll.Value = 1 Then

                .lblAction(i).Top = dGrid(i).Top + dGrid(0).Height / 2 - .lblAction(i).Height / 2
                
                If dLine(i).Left - .lblAction(i).Width - 240 < 0 And (dLine(i).Left + dLine(i).Width) + .lblAction(i).Width + 400 < .Width Then

                    .lblAction(i).Left = (dLine(i).Left + dLine(i).Width) + 200

                ElseIf (dLine(i).Left + dLine(i).Width) + .lblAction(i).Width + 400 > .Width And dLine(i).Left - .lblAction(i).Width - 240 > 0 Then

                    .lblAction(i).Left = dLine(i).Left - .lblAction(i).Width - 200

                ElseIf (dLine(i).Left + dLine(i).Width) + .lblAction(i).Width + 400 > .Width And dLine(i).Left - .lblAction(i).Width - 240 < 0 Then

                    .lblAction(i).Left = ((dLine(i).Left + dLine(i).Width) / 2) - (.lblAction(i).Width / 2)  '+  dLine(i).X1

                ElseIf (dLine(i).Left + dLine(i).Width) + .lblAction(i).Width + 400 < .Width And dLine(i).Left - .lblAction(i).Width - 240 > 0 Then

                    .lblAction(i).Left = (dLine(i).Left + dLine(i).Width) + 200

                End If

            Else
            
            End If
        
        End With
        
    Next i
  
    'Day Lines
  
    If DrawDayLines = True Then
  
        If Days > 0 Then
            DStep = ((dLine(UBound(dLine)).Left + dLine(UBound(dLine)).Width) - frmTimeLine.lnScale.X1) / Days
            DStep = Round(DStep, 0)
            
        Else

        End If
    
        dDayLine(0).Y1 = frmTimeLine.lnVisScale.Y1
      
        dDayLine(0).Y2 = dGrid(0).Top
        
        dDayLine(0).X1 = frmTimeLine.lnVisScale.X1
        dDayLine(0).X2 = frmTimeLine.lnVisScale.X1
      
    
        For i = 1 To Days
        
            dDayLine(i).Y1 = frmTimeLine.lnVisScale.Y1
            dDayLine(i).Y2 = dGrid(0).Top
            dDayLine(i).X1 = dDayLine(i - 1).X1 + DStep
            dDayLine(i).X2 = dDayLine(i - 1).X2 + DStep
        
        Next i
  
    End If
    
    lnVisScale.X2 = (dLine(UBound(dLine)).Left + dLine(UBound(dLine)).Width)

    If frmTimeLine.lblPacketAge.Top + 1100 >= frmTimeLine.Height Then
        frmTimeLine.pbDrawArea.Height = frmTimeLine.lblPacketAge.Top + 2500
    Else
        frmTimeLine.pbDrawArea.Height = frmTimeLine.picWindow.Height

    End If

    frmTimeLine.VScroll1.Max = frmTimeLine.pbDrawArea.ScaleHeight - frmTimeLine.picWindow.ScaleHeight
    
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
    .pbDrawArea.DrawMode = 3
    
        For i = 0 To UBound(dDayLine) 'draw day lines

            .pbDrawArea.DrawStyle = 2
        
            .pbDrawArea.Line (dDayLine(i).X1, dDayLine(i).Y1)-(dDayLine(i).X2, dDayLine(i).Y2), &H404040

        Next
 .pbDrawArea.DrawMode = 13
    End If
    
    

        For i = 0 To UBound(dLine) 'draw bars
          
            .pbDrawArea.DrawStyle = 0
            .pbDrawArea.FillColor = dLine(i).Color

            .pbDrawArea.Line (dLine(i).Left, dLine(i).Top)-(dLine(i).Left + dLine(i).Width, dLine(i).Top + dLine(0).Height), vbBlack, B

        Next

    End With

End Sub
Private Sub UnloadControls()

    Dim i As Integer

    For i = 1 To UBound(dLine)
     
        Unload lblAction(i)

    Next

End Sub

Private Sub chkDayLines_Click()

    If chkDayLines.Value = 1 Then
        DrawDayLines = True
       ' DayLine(0).Visible = True

        Form1.DrawTimeLine
        ReDrawTimeLine
    Else
        DrawDayLines = False
        'DayLine(0).Visible = False

        Form1.DrawTimeLine
        ReDrawTimeLine

    End If

End Sub

Private Sub chkShowAll_Click()

    If chkShowAll.Value = 0 Then
        tmrActionShow.Enabled = True

    Else
        tmrActionShow.Enabled = False
        Form1.DrawTimeLine
        ReDrawTimeLine

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
    VScroll1.Left = (frmTimeLine.Width - VScroll1.Width) - 225
    VScroll1.Top = 0
    
    pbDrawArea.Left = 0
    pbDrawArea.Width = picWindow.Width
    MouseXPrev = 0
    MouseYPrev = 0
    
End Sub

Private Sub CoordinateMouse()

    On Error Resume Next

    ret = GetCursorPos(a)
    ScreenToClient Me.hwnd, a
    b = a.X * Screen.TwipsPerPixelX
    c = a.Y * Screen.TwipsPerPixelY
    
    MouseX = b
    MouseY = c + VScroll1.Value

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnloadControls

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    'Me.Height = 7320

    picWindow.Left = 0
    picWindow.Width = (frmTimeLine.Width - VScroll1.Width) - 100
    picWindow.Height = frmTimeLine.Height
    'pbDrawArea.Height = picWindow.Height
    pbDrawArea.Left = 0
    pbDrawArea.Width = picWindow.Width
    
    VScroll1.Height = frmTimeLine.Height - StatusBar1.Height - 500
   
    VScroll1.SmallChange = 2
    VScroll1.LargeChange = picWindow.Height
    VScroll1.Left = (frmTimeLine.Width - VScroll1.Width) - 100
    VScroll1.Top = 0

    lnScale.X2 = pbDrawArea.Width - 500

    lblPacketAge.Left = (lnScale.X2 / 2) - (lblPacketAge.Width / 2)

    cmdDone.Left = (Me.Width / 2) - 500
    cmdDone.Top = frmTimeLine.Height - cmdDone.Height - 550

    'Form1.DrawTimeLine
    ReDrawTimeLine

    'Me.Refresh
    pbDrawArea.Refresh
    cmdCantSeeMe.SetFocus

End Sub

Private Sub pbDrawArea_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
End Sub

Private Sub tmrActionShow_Timer()

    Dim intOffset As Integer

    intOffset = 800
    
    Dim i               As Integer
    Dim intNumofActions As Integer
    Call CoordinateMouse

    If chkShowAll.Value = False Then
        
        If MouseX <> MouseXPrev Or MouseY <> MouseYPrev Then
        
            MouseXPrev = MouseX
            MouseYPrev = MouseY
        
            For i = 0 To UBound(dGrid)
                'If MouseY > Lines(i).Top And MouseY < (Lines(i).Top + Lines(i).Height) And MouseX > Lines(i).Left - intOffset And MouseX < (Lines(i).Left + Lines(i).Width) + intOffset Then
                If MouseY > dLine(i).Top And MouseY < (dLine(i).Top + dLine(0).Height) And MouseX > dLine(i).Left - intOffset And MouseX < (dLine(i).Left + dLine(i).Width) + intOffset Then
                
                    If MouseX + 20 + lblAction(i).Width >= frmTimeLine.pbDrawArea.Width Then
                        lblAction(i).Left = (MouseX - lblAction(i).Width) - 300

                    Else
                        lblAction(i).Left = MouseX + 300

                    End If

                    lblAction(i).Top = MouseY - lblAction(i).Height
                    lblAction(i).Visible = True
                    lblNote(i).Top = lblAction(i).Top + lblAction(i).Height
                    lblNote(i).Left = lblAction(i).Left
                    lblNote(i).Width = lblAction(i).Width
                    lblNote(i).BackColor = lblAction(i).BackColor
            
                    If lblNote(i).Caption <> "" Then lblNote(i).Visible = True
                
                    'DoEvents
                    lnPoint.X1 = dLine(i).Left + dLine(i).Width / 2
                    lnPoint.Y1 = dLine(i).Top + dLine(0).Height / 2
                    lnPoint.X2 = lblAction(i).Left + lblAction(i).Width / 2 'MouseX
                    lnPoint.Y2 = lblAction(i).Top + lblAction(i).Height / 2 'MouseY
                
                Else
                    lblAction(i).Visible = False
                    lblNote(i).Visible = False
                    lnPoint.Visible = False
                
                    intNumofActions = intNumofActions + 1
                
                End If

            Next i
     
            If intNumofActions > UBound(dGrid) Then 'if no actions are visible, hide pointer line.
        
                lnPoint.Visible = False
            Else
                lnPoint.Visible = True
       
            End If
           
        End If
    Else

        For i = 0 To UBound(dGrid)

            lblAction(i).Visible = True

        Next i

        tmrActionShow.Enabled = False

    End If

End Sub

Private Sub VScroll1_Change()
    pbDrawArea.Top = -(VScroll1.Value)
    pbDrawArea.Refresh
    cmdCantSeeMe.SetFocus

End Sub

Private Sub VScroll1_Scroll()
    pbDrawArea.Top = -(VScroll1.Value)
    pbDrawArea.Refresh

End Sub
