VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Search"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportFilter.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.CheckBox chkHeatMap 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Heat Map"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   39
         ToolTipText     =   "Shows heat map of packet entries. (More entries = hotter)"
         Top             =   4920
         Width           =   1035
      End
      Begin VB.Frame Frame5 
         Caption         =   "Plant"
         Height          =   2775
         Left            =   3120
         TabIndex        =   30
         Top             =   1800
         Width           =   2055
         Begin VB.CheckBox chkW 
            Appearance      =   0  'Flat
            Caption         =   "Wooster"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox chkC 
            Appearance      =   0  'Flat
            Caption         =   "Controls"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkRMT 
            Appearance      =   0  'Flat
            Caption         =   "Rocky Mountain"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox chkN 
            Appearance      =   0  'Flat
            Caption         =   "Nuclear"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkSF 
            Appearance      =   0  'Flat
            Caption         =   "Steel Fab"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkIM 
            Appearance      =   0  'Flat
            Caption         =   "Industrial Machine"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   2160
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Date Range"
         Height          =   1335
         Left            =   2760
         TabIndex        =   22
         Top             =   360
         Width           =   5535
         Begin VB.CheckBox chkAllTickets 
            Appearance      =   0  'Flat
            Caption         =   "All Dates"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2400
            TabIndex        =   29
            Top             =   960
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker MonthStart 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy MM dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   -2147483635
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "MM-dd-yyyy"
            Format          =   254279683
            CurrentDate     =   40405
            MinDate         =   40405
         End
         Begin MSComCtl2.DTPicker MonthEnd 
            Height          =   375
            Left            =   3240
            TabIndex        =   26
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   -2147483635
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "MM-dd-yyyy"
            Format          =   254279683
            CurrentDate     =   40405
            MinDate         =   40405
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ending Date:"
            Height          =   195
            Left            =   3240
            TabIndex        =   27
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "è"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   27.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2280
            TabIndex        =   25
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Date:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Search Text (Beginning with...)"
         Height          =   2775
         Left            =   5280
         TabIndex        =   9
         Top             =   1800
         Width           =   5295
         Begin VB.ComboBox cmbPacketType 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ComboBox cmbUsers 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2520
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox txtSearchCust 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   15
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox txtSearchDraw 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   14
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox txtSearchSales 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   13
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtSearchPart 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtSearchDesc 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   11
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtSearchJobNum 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Packets..."
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   2355
            Width           =   750
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Number:"
            Height          =   195
            Left            =   2400
            TabIndex        =   21
            Top             =   1560
            Width           =   1350
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Drawing Number:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Number:"
            Height          =   195
            Left            =   2400
            TabIndex        =   19
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Number:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   960
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description: (Containing...)"
            Height          =   195
            Left            =   2400
            TabIndex        =   17
            Top             =   360
            Width           =   1965
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Number:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Packet Status"
         Height          =   3000
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   2655
         Begin VB.CheckBox chkReOpened 
            Appearance      =   0  'Flat
            Caption         =   "Re-opened Packets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   40
            Top             =   2520
            Width           =   1815
         End
         Begin VB.CheckBox chkFiled 
            Appearance      =   0  'Flat
            Caption         =   "Filed Packets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CheckBox chkOpened 
            Appearance      =   0  'Flat
            Caption         =   "Open Packets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkClosed 
            Appearance      =   0  'Flat
            Caption         =   "Closed Packets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkInTransit 
            Appearance      =   0  'Flat
            Caption         =   "In-transit Packets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   5
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox chkReceived 
            Appearance      =   0  'Flat
            Caption         =   "Received Packets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   4
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkCreated 
            Appearance      =   0  'Flat
            Caption         =   "Created Packets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   3
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Line Line8 
            X1              =   360
            X2              =   600
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line7 
            X1              =   360
            X2              =   360
            Y1              =   2280
            Y2              =   2640
         End
         Begin VB.Line Line6 
            X1              =   360
            X2              =   600
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line5 
            X1              =   360
            X2              =   360
            Y1              =   1920
            Y2              =   2280
         End
         Begin VB.Line Line4 
            X1              =   360
            X2              =   600
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line3 
            X1              =   360
            X2              =   600
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line2 
            X1              =   360
            X2              =   600
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            X1              =   360
            X2              =   360
            Y1              =   960
            Y2              =   1920
         End
      End
      Begin VB.CommandButton cmdRunReport 
         Caption         =   "Search"
         Height          =   465
         Left            =   4680
         TabIndex        =   1
         Top             =   4800
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmReportFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ClearFields()
    On Error Resume Next
    chkAllTickets.Value = 0
    MonthStart.Enabled = True
    MonthEnd.Enabled = True
    MonthStart.Value = Date
    MonthEnd.Value = Date
    chkClosed.Value = 0
    chkFiled.Value = 0
    chkOpened.Value = 0
    chkInTransit.Value = 0
    chkReceived.Value = 0
    chkCreated.Value = 0
    chkReOpened.Value = 0
    chkClosed.Enabled = True
    chkFiled.Enabled = True
    chkOpened.Enabled = True
    chkInTransit.Enabled = True
    chkReceived.Enabled = True
    chkCreated.Enabled = True
    chkReOpened.Enabled = True
    chkSF.Value = 0
    chkN.Value = 0
    chkRMT.Value = 0
    chkC.Value = 0
    chkW.Value = 0
    chkIM.Value = 0
    txtSearchJobNum.Text = ""
    txtSearchDesc.Text = ""
    txtSearchPart.Text = ""
    txtSearchCust.Text = ""
    txtSearchDraw.Text = ""
    txtSearchCust.Text = ""
    txtSearchSales.Text = ""
    cmbUsers.Enabled = True
    'cmbUsers.ComboItems.Item(1).Selected = True
    'cmbUsers.Enabled = False
    chkHeatMap.Value = 0
End Sub
Private Sub chkAllTickets_Click()
    If chkAllTickets.Value = 1 Then
        MonthStart.Enabled = False
        MonthEnd.Enabled = False
    Else
        MonthStart.Enabled = True
        MonthEnd.Enabled = True
    End If
End Sub
Private Sub chkCreated_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    If chkCreated.Value = 0 Then
        chkOpened.Value = 1
        chkOpened.Enabled = False
    ElseIf chkCreated.Value = 1 And chkInTransit.Value = 0 And chkReceived.Value = 0 And chkFiled.Value = 0 And chkReOpened.Value = 0 Then
        chkOpened.Value = 0
        chkOpened.Enabled = True
    End If
End Sub
Private Sub chkFiled_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    If chkFiled.Value = 0 Then
        chkOpened.Value = 1
        chkOpened.Enabled = False
    ElseIf chkCreated.Value = 0 And chkInTransit.Value = 0 And chkReceived.Value = 0 And chkFiled.Value = 1 And chkReOpened.Value = 0 Then
        chkOpened.Value = 0
        chkOpened.Enabled = True
    End If
End Sub
Private Sub chkInTransit_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    If chkInTransit.Value = 0 Then
        chkOpened.Value = 1
        chkOpened.Enabled = False
    ElseIf chkCreated.Value = 0 And chkInTransit.Value = 1 And chkReceived.Value = 0 And chkFiled.Value = 0 And chkReOpened.Value = 0 Then
        chkOpened.Value = 0
        chkOpened.Enabled = True
    End If
End Sub
Private Sub chkOpened_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    If chkOpened.Value = 0 Then
        chkReceived.Value = 1
        chkReceived.Enabled = False
        chkCreated.Value = 1
        chkCreated.Enabled = False
        chkInTransit.Value = 1
        chkInTransit.Enabled = False
        chkFiled.Value = 1
        chkFiled.Enabled = False
        chkReOpened.Value = 1
        chkReOpened.Enabled = False
    Else
        chkReceived.Value = 0
        chkReceived.Enabled = True
        chkCreated.Value = 0
        chkCreated.Enabled = True
        chkInTransit.Value = 0
        chkInTransit.Enabled = True
        chkFiled.Value = 0
        chkFiled.Enabled = True
        chkReOpened.Value = 0
        chkReOpened.Enabled = True
    End If
End Sub
Private Sub chkReceived_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    If chkReceived.Value = 0 Then
        chkOpened.Value = 1
        chkOpened.Enabled = False
    ElseIf chkCreated.Value = 0 And chkInTransit.Value = 0 And chkReceived.Value = 1 And chkFiled.Value = 0 And chkReOpened.Value = 0 Then
        chkOpened.Value = 0
        chkOpened.Enabled = True
    End If
End Sub
Private Sub chkReOpened_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    If chkReOpened.Value = 0 Then
        chkOpened.Value = 1
        chkOpened.Enabled = False
    ElseIf chkCreated.Value = 0 And chkInTransit.Value = 0 And chkReceived.Value = 0 And chkFiled.Value = 0 And chkReOpened.Value = 1 Then
        chkOpened.Value = 0
        chkOpened.Enabled = True
    End If
End Sub
Private Sub cmbUsers_Click()
    strSearchUser = UCase$(strUserIndex(0, frmReportFilter.cmbUsers.ListIndex))
End Sub
Private Sub cmdRunReport_Click()
    Unload frmReportFilter
    dtStartDate = (MonthStart.Year & "-" & (IIf(Len(MonthStart.Month) < 2, "0" & MonthStart.Month, MonthStart.Month)) & "-" & (IIf(Len(MonthStart.Day) < 2, "0" & MonthStart.Day, MonthStart.Day)))
    dtEndDate = (MonthEnd.Year & "-" & (IIf(Len(MonthEnd.Month) < 2, "0" & MonthEnd.Month, MonthEnd.Month)) & "-" & (IIf(Len(MonthEnd.Day) < 2, "0" & MonthEnd.Day, MonthEnd.Day)))
    DateRangeReport
    sAddlMsg = "Filtered by : " & (IIf(chkCreated.Value = 1, "Created, ", "")) & (IIf(chkReceived.Value = 1, "Received, ", "")) & (IIf(chkInTransit.Value = 1, "In transit, ", "")) & (IIf(chkOpened.Value = 1, "Opened, ", "")) & (IIf(chkFiled.Value = 1, "Filed, ", "")) & (IIf(chkClosed.Value = 1, "Closed ", "")) & vbCrLf & "      Plants: " & (IIf(chkSF.Value = 1, "Steel Fab, ", "")) & (IIf(chkN.Value = 1, "Nuclear, ", "")) & (IIf(chkRMT.Value = 1, "Rocky Mt, ", "")) & (IIf(chkC.Value = 1, "Controls, ", "")) & (IIf(chkW.Value = 1, "Wooster, ", "")) & (IIf(chkIM.Value = 1, "Industrial Mach, ", "")) & vbCrLf & "      User: " & (IIf(cmbUsers.Text <> "", cmbUsers.Text, "All"))
    ClearFields
End Sub
Private Sub Form_Load()
    MonthStart.Value = Date
    MonthEnd.Value = Date
    chkAllTickets.Value = 1
    cmbPacketType.Clear
    cmbPacketType.AddItem "", 0
    cmbPacketType.AddItem "Owned by:", 1
    cmbPacketType.AddItem "In transit to:", 2
    cmbPacketType.AddItem "Sent by:", 3
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    frmReportFilter.Hide
End Sub
Private Sub MonthEnd_Change()
    dtEndDate = MonthEnd.Value
End Sub
Private Sub MonthStart_Change()
    dtStartDate = MonthStart.Value
End Sub
