VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Packet Tracker"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmConfirm 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3120
      TabIndex        =   84
      Top             =   0
      Visible         =   0   'False
      Width           =   5565
      Begin VB.Shape shpTimer 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   690
         Width           =   3000
      End
      Begin VB.Shape Border 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5325
      End
      Begin VB.Label lblConfirm 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%Info Bar%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   85
         Top             =   285
         Width           =   1260
      End
   End
   Begin VB.Frame frmTimers 
      Caption         =   "Timers"
      Height          =   5055
      Left            =   10755
      TabIndex        =   88
      Top             =   4425
      Visible         =   0   'False
      Width           =   795
      Begin MSComctlLib.ImageList ImgList 
         Left            =   120
         Top             =   4380
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Timer tmrLiveSearch 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   120
         Top             =   495
      End
      Begin VB.Timer tmrBannerWait 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   120
         Top             =   3840
      End
      Begin VB.Timer tmrReSizer 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   1440
      End
      Begin VB.Timer tmrButtonFlasher 
         Interval        =   50
         Left            =   120
         Top             =   2880
      End
      Begin VB.Timer tmrRefresher 
         Interval        =   7000
         Left            =   120
         Top             =   1920
      End
      Begin VB.Timer tmrDateTime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   2400
      End
      Begin VB.Timer tmrScroll 
         Interval        =   1
         Left            =   120
         Top             =   960
      End
      Begin VB.Timer tmrConfirmSlider 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   3360
      End
   End
   Begin VB.Frame frmpBar 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   3960
      TabIndex        =   78
      Top             =   6000
      Visible         =   0   'False
      Width           =   5355
      Begin ComctlLib.ProgressBar pBar 
         Height          =   405
         Left            =   120
         TabIndex        =   79
         Top             =   840
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   714
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Progress..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   5190
      End
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   370
      Left            =   7080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "Edit Field"
      Top             =   720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   9765
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   21484
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      CausesValidation=   0   'False
      Height          =   5175
      Left            =   120
      TabIndex        =   38
      ToolTipText     =   "Click to expand"
      Top             =   4320
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   706
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "History"
      TabPicture(0)   =   "Form1.frx":124E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "Form1.frx":177E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Incoming Packets"
      TabPicture(2)   =   "Form1.frx":1BF0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "On Hand Packets"
      TabPicture(3)   =   "Form1.frx":1D8A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   53
         Top             =   480
         Width           =   11775
         Begin VB.Frame frmKey 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   3
            Left            =   10920
            TabIndex        =   115
            Top             =   3000
            Visible         =   0   'False
            Width           =   768
            Begin VB.Label lblCreated 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080C0FF&
               Caption         =   "Created"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   121
               Top             =   0
               Width           =   765
            End
            Begin VB.Label lblInTransit 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080FF80&
               Caption         =   "In-Transit"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   120
               Top             =   240
               Width           =   765
            End
            Begin VB.Label lblReceived 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080FFFF&
               Caption         =   "Received"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   119
               Top             =   480
               Width           =   765
            End
            Begin VB.Label lblClosed 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H008080FF&
               Caption         =   "Closed"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   118
               Top             =   720
               Width           =   765
            End
            Begin VB.Label lblFiled 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Filed"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   117
               Top             =   960
               Width           =   765
            End
            Begin VB.Label lblReopened 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FF80FF&
               Caption         =   "Reopened"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   116
               Top             =   1200
               Width           =   765
            End
         End
         Begin VB.CommandButton cmdPrintOnPack 
            Caption         =   "&Print"
            Height          =   840
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Form1.frx":247C
            Style           =   1  'Graphical
            TabIndex        =   62
            TabStop         =   0   'False
            ToolTipText     =   "Print Report"
            Top             =   3600
            UseMaskColor    =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdGetOutBox 
            Caption         =   "Refresh Packets"
            Height          =   360
            Left            =   120
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "Maually refresh my packets"
            Top             =   480
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGridOUT 
            Height          =   4215
            Left            =   1560
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7435
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Image imgNewWindowOut 
            Appearance      =   0  'Flat
            Height          =   450
            Left            =   600
            Picture         =   "Form1.frx":4010
            ToolTipText     =   "Open grid in a new window"
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label lblColorKey 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Color Key"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   122
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "On-hand Packets"
            Height          =   195
            Left            =   6120
            TabIndex        =   73
            Top             =   2160
            Width           =   1230
         End
         Begin VB.Shape Shape4 
            Height          =   4215
            Left            =   1560
            Top             =   240
            Width           =   10095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   11775
         Begin VB.Frame frmKey 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   2
            Left            =   10920
            TabIndex        =   107
            Top             =   3000
            Visible         =   0   'False
            Width           =   768
            Begin VB.Label lblReopened 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FF80FF&
               Caption         =   "Reopened"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   113
               Top             =   1200
               Width           =   765
            End
            Begin VB.Label lblFiled 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Filed"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   112
               Top             =   960
               Width           =   765
            End
            Begin VB.Label lblClosed 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H008080FF&
               Caption         =   "Closed"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   111
               Top             =   720
               Width           =   765
            End
            Begin VB.Label lblReceived 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080FFFF&
               Caption         =   "Received"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   110
               Top             =   480
               Width           =   765
            End
            Begin VB.Label lblInTransit 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080FF80&
               Caption         =   "In-Transit"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   109
               Top             =   240
               Width           =   765
            End
            Begin VB.Label lblCreated 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080C0FF&
               Caption         =   "Created"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   108
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.CommandButton cmdPrintInPack 
            Caption         =   "&Print"
            Height          =   840
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Form1.frx":4105
            Style           =   1  'Graphical
            TabIndex        =   61
            TabStop         =   0   'False
            ToolTipText     =   "Print Report"
            Top             =   3600
            UseMaskColor    =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdGetInBox 
            Caption         =   "Refresh Packets"
            Height          =   360
            Left            =   120
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "Maually refresh my packets"
            Top             =   480
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGridIN 
            Height          =   4215
            Left            =   1560
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7435
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            FillStyle       =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Image imgNewWindowIn 
            Appearance      =   0  'Flat
            Height          =   450
            Left            =   600
            Picture         =   "Form1.frx":5C99
            ToolTipText     =   "Open grid in a new window"
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label lblColorKey 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Color Key"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   114
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Shape Shape3 
            Height          =   4215
            Left            =   1560
            Top             =   240
            Width           =   10095
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Incoming Packets"
            Height          =   195
            Left            =   6120
            TabIndex        =   72
            Top             =   2160
            Width           =   1245
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4575
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   11775
         Begin VB.Frame frmKey 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   1
            Left            =   10920
            TabIndex        =   99
            Top             =   3000
            Visible         =   0   'False
            Width           =   768
            Begin VB.Label lblCreated 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080C0FF&
               Caption         =   "Created"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   105
               Top             =   0
               Width           =   765
            End
            Begin VB.Label lblInTransit 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080FF80&
               Caption         =   "In-Transit"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   104
               Top             =   240
               Width           =   765
            End
            Begin VB.Label lblReceived 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080FFFF&
               Caption         =   "Received"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   103
               Top             =   480
               Width           =   765
            End
            Begin VB.Label lblClosed 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H008080FF&
               Caption         =   "Closed"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   102
               Top             =   720
               Width           =   765
            End
            Begin VB.Label lblFiled 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Filed"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   101
               Top             =   960
               Width           =   765
            End
            Begin VB.Label lblReopened 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FF80FF&
               Caption         =   "Reopened"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   100
               Top             =   1200
               Width           =   765
            End
         End
         Begin VB.CommandButton cmdAllOpenReport 
            Caption         =   "All Opened"
            Height          =   360
            Left            =   120
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "Display all currently open packets"
            Top             =   795
            Width           =   1335
         End
         Begin VB.CommandButton cmdAllClosedReport 
            Caption         =   "All Closed"
            Height          =   360
            Left            =   120
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Display all currently closed packets"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton cmdFilterReport 
            Caption         =   "Custom Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Run a custom filtered search"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdPrintReport 
            Caption         =   "&Print"
            Height          =   840
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Form1.frx":5D8E
            Style           =   1  'Graphical
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Print Report"
            Top             =   3600
            UseMaskColor    =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdHeatMap 
            Caption         =   "Entry Heat Map"
            Height          =   360
            Left            =   0
            TabIndex        =   90
            ToolTipText     =   "Shows heat map of packet entries. (More entries = hotter)"
            Top             =   4200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flexgrid1 
            Height          =   4215
            Left            =   1560
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7435
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            RowHeightMin    =   285
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            GridLinesUnpopulated=   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblColorKey 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Color Key"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   106
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Image imgNewWindow 
            Appearance      =   0  'Flat
            Height          =   450
            Left            =   600
            Picture         =   "Form1.frx":7922
            ToolTipText     =   "Open grid in a new window"
            Top             =   1680
            Width           =   450
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Packet Search"
            Height          =   195
            Left            =   6120
            TabIndex        =   86
            Top             =   2160
            Width           =   555
         End
         Begin VB.Shape Shape5 
            Height          =   4215
            Left            =   1560
            Top             =   240
            Width           =   10095
         End
      End
      Begin VB.Frame Frame1 
         ClipControls    =   0   'False
         Height          =   4575
         Left            =   -74880
         TabIndex        =   39
         Top             =   480
         Width           =   11775
         Begin VB.Frame frmKey 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   0
            Left            =   10920
            TabIndex        =   91
            Top             =   3000
            Visible         =   0   'False
            Width           =   768
            Begin VB.Label lblReopened 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FF80FF&
               Caption         =   "Reopened"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   97
               Top             =   1200
               Width           =   765
            End
            Begin VB.Label lblFiled 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Filed"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   96
               Top             =   960
               Width           =   765
            End
            Begin VB.Label lblClosed 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H008080FF&
               Caption         =   "Closed"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   95
               Top             =   720
               Width           =   765
            End
            Begin VB.Label lblReceived 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080FFFF&
               Caption         =   "Received"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   94
               Top             =   480
               Width           =   765
            End
            Begin VB.Label lblInTransit 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080FF80&
               Caption         =   "In-Transit"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   93
               Top             =   240
               Width           =   765
            End
            Begin VB.Label lblCreated 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H0080C0FF&
               Caption         =   "Created"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   92
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.PictureBox picOlder 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1560
            Picture         =   "Form1.frx":7A17
            ScaleHeight     =   300
            ScaleWidth      =   9810
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   4140
            Visible         =   0   'False
            Width           =   9810
         End
         Begin VB.CommandButton cmdTimeLine 
            Caption         =   "View Timeline"
            Height          =   480
            Left            =   120
            TabIndex        =   66
            TabStop         =   0   'False
            ToolTipText     =   "Displays a visual representation of packet activity"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdPrintHistory 
            Caption         =   "&Print"
            Height          =   840
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Form1.frx":93CA
            Style           =   1  'Graphical
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Print History"
            Top             =   3600
            UseMaskColor    =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdRefreshHist 
            Caption         =   "Refresh History"
            Height          =   360
            Left            =   120
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Manually refresh history data"
            Top             =   360
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGridHist 
            Height          =   4215
            Left            =   1560
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7435
            _Version        =   393216
            BackColor       =   16777215
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            GridLines       =   0
            GridLinesFixed  =   0
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   65
            Text            =   "Form1.frx":AF5E
            Top             =   240
            Visible         =   0   'False
            Width           =   8025
         End
         Begin VB.Label lblColorKey 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Color Key"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   98
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Image imgNewWindowHist 
            Appearance      =   0  'Flat
            Height          =   450
            Left            =   600
            Picture         =   "Form1.frx":AF66
            ToolTipText     =   "Open grid in a new window"
            Top             =   1560
            Width           =   450
         End
         Begin VB.Shape Shape2 
            Height          =   4215
            Left            =   1560
            Top             =   240
            Width           =   10095
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "History Viewer"
            Height          =   195
            Left            =   6120
            TabIndex        =   51
            Top             =   2160
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tracking Info."
      Height          =   3975
      Left            =   7380
      TabIndex        =   23
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   2490
         TabIndex        =   75
         Top             =   2730
         Width           =   2175
         Begin VB.PictureBox pbData 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   750
            Left            =   1320
            ScaleHeight     =   750
            ScaleWidth      =   765
            TabIndex        =   124
            TabStop         =   0   'False
            Top             =   360
            Width           =   765
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   360
            Left            =   120
            TabIndex        =   77
            TabStop         =   0   'False
            ToolTipText     =   "Manually refresh all data"
            Top             =   510
            Width           =   990
         End
         Begin VB.CheckBox chkAutoRefresh 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto Refresh"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   195
            Value           =   1  'Checked
            Width           =   1260
         End
         Begin VB.Label lblQryTime 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "%AVG QRY TIME MS%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   135
            TabIndex        =   89
            ToolTipText     =   "Avg. Query Time"
            Top             =   960
            Width           =   1080
         End
      End
      Begin VB.PictureBox pbScrollBox 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   4185
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtLocalUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "%USERNAME%"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtActionDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtDateTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Text            =   "%DATETIME%"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox txtTicketStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtTicketAction 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtCreateDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtCreator 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtTicketOwner 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   71
         Top             =   2760
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latest Note"
         Height          =   195
         Left            =   240
         TabIndex        =   69
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "@"
         Height          =   195
         Left            =   2280
         TabIndex        =   60
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local User"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action Date"
         Height          =   195
         Left            =   2520
         TabIndex        =   59
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Status"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latest Action"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create Date"
         Height          =   195
         Left            =   2520
         TabIndex        =   32
         Top             =   2160
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packet Creator"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Packet Owner"
         Height          =   195
         Left            =   2520
         TabIndex        =   30
         Top             =   1560
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Packet Info."
      Height          =   3975
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7215
      Begin MSComctlLib.ImageCombo cmbUsers 
         Height          =   330
         Left            =   1680
         TabIndex        =   123
         Top             =   2580
         Visible         =   0   'False
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cmbPlant 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton cmdSubmit 
         Appearance      =   0  'Flat
         Caption         =   "Submit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2280
         MaskColor       =   &H000000FF&
         TabIndex        =   7
         ToolTipText     =   "Submit update"
         Top             =   3270
         Width           =   1815
      End
      Begin VB.OptionButton optFile 
         Appearance      =   0  'Flat
         Caption         =   "File Packet"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   42
         Top             =   3480
         Width           =   1455
      End
      Begin VB.OptionButton optReOpen 
         Appearance      =   0  'Flat
         Caption         =   "Reopen/Unfile Packet"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtTicketDescription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear All"
         Height          =   360
         Left            =   1440
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtJobNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   15
         TabIndex        =   0
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optCreate 
         Appearance      =   0  'Flat
         Caption         =   "New Packet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optClose 
         Appearance      =   0  'Flat
         Caption         =   "Close Packet"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   990
      End
      Begin VB.TextBox txtPartNoRev 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtCustPoNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   5
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtSalesNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.OptionButton optMove 
         Appearance      =   0  'Flat
         Caption         =   "Send Packet To"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   15
         Top             =   2730
         Width           =   1695
      End
      Begin VB.OptionButton optReceive 
         Appearance      =   0  'Flat
         Caption         =   "Receive Packet"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtDrawNoRev 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   4
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   5010
         TabIndex        =   81
         Top             =   3205
         Width           =   2175
         Begin VB.CommandButton cmdShowMore 
            Caption         =   "Show Tabs"
            Height          =   360
            Left            =   480
            TabIndex        =   82
            TabStop         =   0   'False
            ToolTipText     =   "Show additional features"
            Top             =   230
            Width           =   1575
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "â"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   15.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   80
            TabIndex        =   83
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Image imgComment 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   4080
         Picture         =   "Form1.frx":B05B
         ToolTipText     =   "Add Note"
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label lblChars 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(0/0)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   165
         Left            =   4080
         TabIndex        =   67
         ToolTipText     =   "Current Chars / Max Chars"
         Top             =   1800
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Shape shpButtonFlash 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         Height          =   615
         Left            =   2190
         Shape           =   4  'Rounded Rectangle
         Top             =   3195
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2160
         TabIndex        =   63
         Top             =   2940
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plant"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4800
         TabIndex        =   50
         Top             =   2520
         Width           =   435
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   37
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4800
         TabIndex        =   22
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer && PO No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4800
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drawing No. && Rev."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part No. && Rev."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   18
         Top             =   1080
         Width           =   1260
      End
   End
   Begin VB.Label lblAppVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%APP VERSION%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   165
      Left            =   120
      TabIndex        =   87
      Top             =   9540
      Width           =   1290
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by Bobby Lovell"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   165
      Left            =   10560
      TabIndex        =   36
      Top             =   9540
      Width           =   1470
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administrator"
      Begin VB.Menu mnuRedirect 
         Caption         =   "Redirect Current Packet"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Current Packet"
      End
      Begin VB.Menu mnuFauxUser 
         Caption         =   "Faux Local User"
      End
      Begin VB.Menu mnuStats 
         Caption         =   "DB Stats"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuDeleteEntry 
         Caption         =   "^ Delete ^"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_TAB = &H9
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Const EM_GETLINECOUNT = &HBA
Private TheX As Long
Private TheY As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect _
                Lib "user32" (ByVal hwnd As Long, _
                              lpRect As RECT) As Long
Private Declare Function ScreenToClientAny _
                Lib "user32" _
                Alias "ScreenToClient" (ByVal hwnd As Long, _
                                        lpPoint As Any) As Long
Private Declare Function MoveWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal nWidth As Long, _
                              ByVal nHeight As Long, _
                              ByVal bRepaint As Long) As Long
Private Declare Function TranslateColor _
                Lib "olepro32.dll" _
                Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, _
                                           ByVal palet As Long, _
                                           col As Long) As Long
Private bolNoHits As Boolean
Private intRowSel As Integer
Private strCommentText As String
Private Function GetRealColor(ByVal Color As OLE_COLOR) As Long
    Dim R As Long
    R = TranslateColor(Color, 0, GetRealColor)
    If R <> 0 Then 'raise an error
    End If
End Function
Public Sub SetComboBoxHeight(ComboBox As ImageCombo, ByVal NewHeight As Long)
    Dim lpRect As RECT
    Dim wi     As Long
    GetWindowRect ComboBox.hwnd, lpRect
    wi = lpRect.Right - lpRect.Left
    ScreenToClientAny ComboBox.Parent.hwnd, lpRect
    MoveWindow ComboBox.hwnd, lpRect.Left, lpRect.Top, wi, NewHeight, True
End Sub
Private Function GetTabState() As Boolean
    GetTabState = False
    If GetKeyState(VK_TAB) And -256 Then
        GetTabState = True
    End If
End Function
Public Sub BannerClick(Optional ClickCase As String)
    On Error Resume Next
    Select Case ClickCase
        Case "VIEWPACK"
            CloseBanner
            If bolOpenForm = True Then
                Call cmdShowMore_Click
                SSTab1.Tab = 2
            Else
                SSTab1.Tab = 2
            End If
        Case "NEWPACK"
            CloseBanner
            optCreate.Value = True
            optCreate.Enabled = True
            SetBoxesForEdit "All"
            EnableBoxes
            'ClearFields
            bolOptionClicked = True
            cmbUsers.Visible = False
            lblUser.Visible = False
            imgComment.Picture = ButtonPics(3)
            imgComment.Enabled = True
            frmComments.txtComment.Text = ""
            frmComments.txtComment.Locked = False
            If txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtSalesNo.Text <> "" And txtTicketDescription.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Then
                cmdSubmit.Enabled = True
            End If
            txtTicketDescription.SetFocus
            txtTicketDescription.BackColor = &HC0FFC0
            cmbPlant.BackColor = &HC0FFC0
            txtCustPoNo.BackColor = &HC0FFC0
            txtDrawNoRev.BackColor = &HC0FFC0
            txtSalesNo.BackColor = &HC0FFC0
            txtPartNoRev.BackColor = &HC0FFC0
            txtJobNo.BackColor = vbWindowBackground
        Case ""
            CloseBanner
        Case "CLOSE"
            CloseBanner
    End Select
End Sub
Public Sub FillFlexHist(strAction As String, _
                        strStatus As String, _
                        strComment As String, _
                        strDate As String, _
                        strCreator As String, _
                        strUserFrom As String, _
                        strUserTo As String, _
                        strUser As String, _
                        strGUID As String)
    Dim lngFontSize As Long
    lngFontSize = 9.5
    FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    If strAction = "CREATED" Then
        If strComment <> "" And bolPrinting = False Then
            FlexGridHist.Rows = FlexGridHist.Rows + 1 ' Add new row for comments
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 4) = "com"
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = "    " & Chr$(34) & strComment & Chr$(34)
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 1
            FlexGridHist.CellFontSize = lngFontSize
            FlexGridHist.CellFontItalic = True
            'FlexGridHist.CellFontBold = True
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &H80C0FF)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | Job packet was CREATED by " & strCreator
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(1)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &H80C0FF)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    ElseIf strAction = "INTRANSIT" Then
        If strComment <> "" And bolPrinting = False Then
            FlexGridHist.Rows = FlexGridHist.Rows + 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 4) = "com"
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = "    " & Chr$(34) & strComment & Chr$(34)
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 1
            FlexGridHist.CellFontSize = lngFontSize
            FlexGridHist.CellFontItalic = True
            'FlexGridHist.CellFontBold = True
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &H80FF80)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUserFrom & " SENT the job packet to " & strUserTo
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(2)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &H80FF80)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    ElseIf strAction = "RECEIVED" Then
        If strComment <> "" And bolPrinting = False Then
            FlexGridHist.Rows = FlexGridHist.Rows + 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 4) = "com"
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = "    " & Chr$(34) & strComment & Chr$(34)
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 1
            FlexGridHist.CellFontSize = lngFontSize
            FlexGridHist.CellFontItalic = True
            'FlexGridHist.CellFontBold = True
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &H80FFFF)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUser & " RECEIVED the job packet from " & strUserFrom
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(3)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &H80FFFF)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    ElseIf strStatus = "CLOSED" Then
        If strComment <> "" And bolPrinting = False Then
            FlexGridHist.Rows = FlexGridHist.Rows + 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 4) = "com"
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = "    " & Chr$(34) & strComment & Chr$(34)
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 1
            FlexGridHist.CellFontSize = lngFontSize
            FlexGridHist.CellFontItalic = True
            'FlexGridHist.CellFontBold = True
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &H8080FF)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUser & " CLOSED the job packet."
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(5)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &H8080FF)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    ElseIf strStatus = "OPEN" And strAction = "FILED" Then
        If strComment <> "" And bolPrinting = False Then
            FlexGridHist.Rows = FlexGridHist.Rows + 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 4) = "com"
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = "    " & Chr$(34) & strComment & Chr$(34)
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 1
            FlexGridHist.CellFontSize = lngFontSize
            FlexGridHist.CellFontItalic = True
            'FlexGridHist.CellFontBold = True
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &HFF8080)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUser & " FILED the job packet."
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(4)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &HFF8080)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    ElseIf strAction = "REOPENED" Then
        If strComment <> "" And bolPrinting = False Then
            FlexGridHist.Rows = FlexGridHist.Rows + 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 4) = "com"
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = "    " & Chr$(34) & strComment & Chr$(34)
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 1
            FlexGridHist.CellFontSize = lngFontSize
            FlexGridHist.CellFontItalic = True
            'FlexGridHist.CellFontBold = True
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &HFF80FF)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUser & " REOPENED the job packet."
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(7)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, &HFF80FF)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    End If
End Sub
Public Sub OpenPacket(JobNum As String) 'Opens Packet - Fills HistoryGrid, Fills Fields, Does not refresh MyPackets
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim b       As Integer
    Dim R       As Integer
    Dim CRow    As Integer
    On Error GoTo errhandle
    If Trim$(JobNum) = "" Then Exit Sub
    txtJobNo.Text = JobNum
    SetBoxesForEdit "All"
    txtJobNo.Text = UCase$(txtJobNo.Text)
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    ShowData
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From ticketdatabase Where idTicketJobNum = '" & JobNum & "' Order By ticketdatabase.idTicketDate Desc"
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    List1.Clear
    If rs.RecordCount <= 0 Then Err.Raise vbObjectError + 513, "ADO Open", "Zero Records Returned For Query"
    With rs
        dtLatestHistDate = Format$(!idTicketDate, strDBDateTimeFormat)
        txtPartNoRev.Text = !idTicketPartNum
        txtDrawNoRev.Text = !idTicketDrawingNum
        txtCustPoNo.Text = !idTicketCustPoNum
        txtSalesNo.Text = !idTicketSalesNum
        txtCreator.Text = !idTicketCreator
        txtCreateDate.Text = !idTicketCreateDate
        txtTicketOwner.Text = !idTicketUser
        txtActionDate.Text = !idTicketDate
        strTicketAction = !idTicketAction
        strUserFrom = !idTicketUserFrom
        strUserTo = !idTicketUserTo
        strTicketStatus = !idTicketStatus
        txtTicketStatus.Text = UCase$(!idTicketStatus)
        strCurUser = !idTicketUser
        txtTicketAction.Text = !idTicketAction
        txtTicketDescription.Text = !idTicketDescription
        strLatestComment = !idTicketComment
        frmComments.txtComment.Text = strLatestComment
        frmComments.txtComment.Locked = True
        strPlant = !idTicketPlant
        cmbPlant.Text = strPlant
        If !idTicketComment <> "" Then
            TheX = pbScrollBox.ScaleWidth
            strCommentText = !idTicketComment
            tmrScroll.Enabled = True
        Else
            pbScrollBox.Cls
            strCommentText = ""
            tmrScroll.Enabled = False
        End If
        If rs.RecordCount >= 1 Then
            bolHasTicket = True
        ElseIf rs.RecordCount <= 0 Then
            bolHasTicket = False
        End If
    End With
    FlexGridHist.Redraw = False
    FlexGridHist.Visible = False
    FlexHistLastTopRow = 0
    FlexGridHist.Clear
    FlexGridHist.Cols = 6
    FlexGridHist.Rows = 0
    rs.MoveLast
    Do Until rs.BOF
        With rs
            Call FillFlexHist(!idTicketAction, !idTicketStatus, !idTicketComment, !idTicketDate, !idTicketCreator, !idTicketUserFrom, !idTicketUserTo, !idTicketUser, !idGUID)
            rs.MovePrevious
        End With
    Loop
    For b = 0 To FlexGridHist.Cols - 1
        FlexGridHist.ColAlignment(b) = flexAlignLeftCenter
    Next b
    FlexGridHist.ColWidth(0) = 1000
    FlexGridHist.ColWidth(1) = 8500
    FlexGridHist.ColWidth(3) = 0
    FlexGridHist.ColWidth(4) = 0
    FlexGridHist.RowHeight(0) = 0
    FlexGridHist.TopRow = FlexHistLastTopRow
    Call FlexFlipHist("D")
    DisplayArrows
    FlexBoldFirst FlexGridHist
    'FlexGridHist.Rows = FlexGridHist.Rows - 1
    FlexGridRedrawHeight
    FlexGridHist.Redraw = True
    FlexGridHist.Visible = True
    rs.Close
    cn.Close
    HideData
    DisableBoxes
    SetControls
    optMove.Value = False
    optReceive.Value = False
    optClose.Value = False
    optFile.Value = False
    optReOpen.Value = False
    optCreate.Value = False
    cmbUsers.Visible = False
    lblUser.Visible = False
    cmbUsers.ComboItems.Item(1).Selected = True
    bolOptionClicked = False
    cmdSubmit.Enabled = False
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    FlexHistLastTopRow = 0
    Screen.MousePointer = vbDefault
    lblChars.Visible = False
    Exit Sub
errhandle:
    If Hex(Err.Number) = 80040201 Then
        bolHasTicket = False
        Screen.MousePointer = vbDefault
        optMove.Value = 0
        optReceive.Value = 0
        optClose.Value = 0
        optCreate.Value = 0
        optReOpen.Value = 0
        optFile.Value = 0
        optMove.Enabled = False
        optReceive.Enabled = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = False
        If bolBannerOpen = False Then ShowBanner &HFFFFC0, "No Job Packet found with that job number. Click here to create a new one.", 1000, "NEWPACK"
        lblChars.Visible = False
        Set pbData.Picture = picDataPics(0)
        Err.Clear
        HideData
    ElseIf Err.Number = -2147467259 Then
        Screen.MousePointer = vbDefault
        CommsDown
    ElseIf Err.Number = 0 Then
        Screen.MousePointer = vbDefault
        CommsUp
    Else
        Dim blah
        blah = MsgBox("An error was detected!" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Yikes!")
        'Resume Next
        ClearFields
    End If
End Sub
Public Sub PositionMaxChar(ActiveBox As Control)
    lblChars.Top = ActiveBox.Top - 200
    lblChars.Left = (ActiveBox.Left + ActiveBox.Width) - 450
    lblChars.Caption = "(" & Len(ActiveBox.Text) & "/" & ActiveBox.MaxLength & ")"
    If Len(ActiveBox.Text) >= ActiveBox.MaxLength Then
        lblChars.ForeColor = &HFF&
    Else
        lblChars.ForeColor = &H8000&
    End If
    If optCreate.Value = True Or EditMode = True Then
        lblChars.Visible = True
    Else
        lblChars.Visible = False
    End If
End Sub
Public Sub GetTimeLineData()
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim dtTicketDate1, dtTicketDate2 As Date
    On Error Resume Next
    ShowData
    strSQL1 = "SELECT * From ticketdatabase Where idTicketJobNum = '" & txtJobNo.Text & "' Order By ticketdatabase.idTicketDate"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    Entry = 0
    With rs
        ReDim strTimelineComments(.RecordCount)
        dtTicketDate1 = !idTicketDate
        .MoveLast
        dtTicketDate2 = !idTicketDate
        TotalTime = DateDiff("n", dtTicketDate1, dtTicketDate2)
        .MoveFirst
    End With
    Do Until rs.EOF
        With rs
            If !idTicketComment <> "" Then strTimelineComments(Entry) = Chr(34) & !idTicketComment & Chr(34)
            dtTicketDate1 = !idTicketDate
            .MoveNext
            If .EOF Then
                .MovePrevious
                dtTicketDate1 = !idTicketDate
                dtTicketDate2 = Date & " " & Time
                TicketHours(Entry) = DateDiff("n", dtTicketDate1, dtTicketDate2)
                .MoveNext
            Else
                dtTicketDate2 = !idTicketDate
                TicketHours(Entry) = DateDiff("n", dtTicketDate1, dtTicketDate2)
            End If
            .MovePrevious
            If !idTicketAction = "CREATED" Then
                TicketActionText(Entry) = " Job packet was CREATED by " & !idTicketCreator & " | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days ", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idTicketAction = "INTRANSIT" Then
                TicketActionText(Entry) = " " & !idTicketUserFrom & " SENT the job packet to " & !idTicketUserTo & " | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days ", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idTicketAction = "RECEIVED" Then
                TicketActionText(Entry) = " " & !idTicketUser & " RECEIVED the job packet from " & !idTicketUserFrom & " | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days ", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idTicketStatus = "CLOSED" Then
                TicketActionText(Entry) = " " & !idTicketUser & " CLOSED the job packet. | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idTicketStatus = "OPEN" And !idTicketAction = "FILED" Then
                TicketActionText(Entry) = " " & !idTicketUser & " FILED the job packet. | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idTicketAction = "REOPENED" Then
                TicketActionText(Entry) = " " & !idTicketUser & " REOPENED the job packet. | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            End If
            TicketDate(Entry) = !idTicketDate
            TicketAction(Entry) = !idTicketAction
            .MoveNext
            Entry = Entry + 1
        End With
    Loop
    rs.Close
    cn.Close
    HideData
End Sub
Public Sub DrawTimeLine()
    Dim i, Days As Integer
    Dim DStep As Single
    On Error Resume Next
    LStep = (frmTimeLine.lnScale.X2 - frmTimeLine.lnScale.X1) / (TotalTime + TicketHours(Entry - 1))
    frmTimeLine.pbDrawArea.FillColor = &H80C0FF
    ReDim dLine(Entry - 1)
    dLine(0).Color = &H80C0FF
    dLine(0).Height = 300
    dLine(0).Left = 470
    dLine(0).Top = 120
    dLine(0).Width = 315
    ReDim dGrid(Entry - 1)
    dGrid(0).Color = &HE0E0E0
    dGrid(0).Height = 300
    dGrid(0).Left = 0
    dGrid(0).Top = 120
    dGrid(0).Width = 11895
    ReDim dAction(Entry - 1)
    ReDim dNote(Entry - 1)
    dAction(0).Text = TicketActionText(0)
    dAction(0).Color = &H80C0FF
    dAction(0).Left = dLine(0).Left + dLine(0).Width + 200
    dAction(0).Top = dLine(0).Top + 20
    dAction(0).Height = 210
    dAction(0).Visible = True
    dNote(0).Height = 210
    dGrid(0).Width = frmTimeLine.Width
    Days = (TotalTime + TicketHours(Entry - 1)) / 1440
    Days = Round(Days, 1)
    For i = 0 To Entry - 1
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
                dLine(i).Color = &H80C0FF
            ElseIf TicketAction(i) = "INTRANSIT" Then
                dLine(i).Color = &H80FF80
            ElseIf TicketAction(i) = "RECEIVED" Then
                dLine(i).Color = &H80FFFF
            ElseIf TicketAction(i) = "NULL" Then
                dLine(i).Color = &H8080FF
            ElseIf TicketAction(i) = "FILED" Then
                dLine(i).Color = &HFF8080
            ElseIf TicketAction(i) = "REOPENED" Then
                dLine(i).Color = &HFF80FF
            End If
            If TicketHours(i) * LStep < 38 Then 'Less than 1 pixel wide
                dLine(i).Width = 38
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width - 38
            Else
                dLine(i).Width = TicketHours(i) * LStep
                dLine(i).Left = dLine(i - 1).Left + dLine(i - 1).Width
            End If
            dNote(i).Text = strTimelineComments(i)
            dNote(i).Width = Printer.TextWidth(dNote(i).Text)
            dAction(i).Text = TicketActionText(i)
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
            dAction(i).Width = frmTimeLine.pbDrawArea.TextWidth(dAction(i).Text)
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
        dDayLine(i).Y1 = frmTimeLine.lnScale.Y1
        dDayLine(i).Y2 = dGrid(0).Top
        dDayLine(i).X1 = dDayLine(i - 1).X1 + DStep
        dDayLine(i).X2 = dDayLine(i - 1).X2 + DStep
    Next i
    With frmTimeLine
        frmTimeLine.DrawLines
        .Image1.ZOrder 0
    End With
    frmTimeLine.lblPacketAge.Top = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 200 + 40
    If frmTimeLine.Width <= 10755 Then frmTimeLine.lblPacketAge.Left = frmTimeLine.Frame1.Left + frmTimeLine.Frame1.Width + 10
    frmTimeLine.lblPacketAge.Caption = "Packet Age: " & (IIf((TotalTime + TicketHours(Entry - 1)) > 1440, Round((TotalTime + TicketHours(Entry - 1)) / 1440, 1) & "days", Round((TotalTime + TicketHours(Entry - 1)) / 60, 1) & "hrs"))
    If frmTimeLine.lblPacketAge.Top + 30 >= frmTimeLine.Height Then
        frmTimeLine.pbDrawArea.Height = frmTimeLine.lblPacketAge.Top + 40
    Else
        frmTimeLine.pbDrawArea.Height = frmTimeLine.picWindow.Height
    End If
    If frmTimeLine.Visible = True Then
        frmTimeLine.VScroll1.Max = frmTimeLine.VScroll1.Max
    Else
        frmTimeLine.VScroll1.Max = frmTimeLine.pbDrawArea.Height - frmTimeLine.picWindow.Height
    End If
    frmTimeLine.Frame1.Top = dGrid(UBound(dGrid)).Top + dGrid(0).Height + 500
End Sub
Public Sub GetMyPackets()
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim LineIN, LineOUT, Row As Integer
    On Error GoTo errs
    strSQL1 = "SELECT * FROM ticketdb.ticketdatabase ticketdatabase_0" & " WHERE (ticketdatabase_0.idTicketAction='CREATED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='RECEIVED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='REOPENED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='INTRANSIT') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') AND (ticketdatabase_0.idTicketUserTo='" & strLocalUser & "')" & " ORDER BY ticketdatabase_0.idTicketDate"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    FlexGridOUT.Clear
    FlexGridOUT.Redraw = False
    FlexGridOUT.Rows = 2
    FlexGridOUT.FixedCols = 1
    FlexGridOUT.FixedRows = 1
    FlexGridIN.Clear
    FlexGridIN.Redraw = False
    FlexGridIN.Rows = 2
    FlexGridIN.FixedCols = 1
    FlexGridIN.FixedRows = 1
    ShowData
    rs.Open strSQL1, cn, adOpenKeyset
    If rs.RecordCount <= 0 Then
        intPrevInPackets = 0
        SSTab1.TabCaption(3) = "On-hand Packets (0)"
        SSTab1.TabCaption(2) = "Incoming Packets (0)"
        FlexGridOUT.Visible = False
        FlexGridOUT.Redraw = True
        FlexGridIN.Visible = False
        FlexGridIN.Redraw = True
        rs.Close
        cn.Close
        HideData
        FlexGridOUT.Clear
        FlexGridIN.Clear
        Exit Sub
    End If
    LineIN = 1
    LineOUT = 1
    Row = 0
    FlexGridOUT.Rows = rs.RecordCount + 1
    FlexGridOUT.Cols = 10
    FlexGridIN.Rows = rs.RecordCount + 1
    FlexGridIN.Cols = 10
    ' Create header row
    FlexGridOUT.TextMatrix(0, 1) = "Job Number"
    FlexGridOUT.TextMatrix(0, 2) = "Part Number"
    FlexGridOUT.TextMatrix(0, 3) = "Description"
    FlexGridOUT.TextMatrix(0, 4) = "Sales Number"
    FlexGridOUT.TextMatrix(0, 5) = "Customer/PO Number"
    FlexGridOUT.TextMatrix(0, 6) = "Created By"
    FlexGridOUT.TextMatrix(0, 7) = "Create Date"
    FlexGridOUT.TextMatrix(0, 8) = "Last Activity Date"
    FlexGridOUT.TextMatrix(0, 9) = "Last Activity"
    FlexGridIN.TextMatrix(0, 1) = "Job Number"
    FlexGridIN.TextMatrix(0, 2) = "Part Number"
    FlexGridIN.TextMatrix(0, 3) = "Description"
    FlexGridIN.TextMatrix(0, 4) = "Sales Number"
    FlexGridIN.TextMatrix(0, 5) = "Customer/PO Number"
    FlexGridIN.TextMatrix(0, 6) = "Created By"
    FlexGridIN.TextMatrix(0, 7) = "Create Date"
    FlexGridIN.TextMatrix(0, 8) = "Last Activity Date"
    FlexGridIN.TextMatrix(0, 9) = "Last Activity"
    Do Until rs.EOF
        With rs
            If !idTicketAction = "CREATED" And !idTicketUser = strLocalUser Or !idTicketAction = "RECEIVED" And !idTicketUser = strLocalUser Or !idTicketAction = "REOPENED" And !idTicketUser = strLocalUser Then
                Row = Row + 1
                FlexGridOUT.TextMatrix(LineOUT, 0) = LineOUT
                FlexGridOUT.TextMatrix(LineOUT, 1) = !idTicketJobNum
                FlexGridOUT.TextMatrix(LineOUT, 2) = !idTicketPartNum
                FlexGridOUT.TextMatrix(LineOUT, 3) = !idTicketDescription
                FlexGridOUT.TextMatrix(LineOUT, 4) = !idTicketSalesNum
                FlexGridOUT.TextMatrix(LineOUT, 5) = !idTicketCustPoNum
                FlexGridOUT.TextMatrix(LineOUT, 6) = !idTicketCreator
                FlexGridOUT.TextMatrix(LineOUT, 7) = !idTicketCreateDate
                FlexGridOUT.TextMatrix(LineOUT, 8) = !idTicketDate
                If !idTicketAction = "CREATED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &H80C0FF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = "Job packet was CREATED by " & !idTicketCreator
                ElseIf !idTicketAction = "RECEIVED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &H80FFFF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = !idTicketUser & " RECEIVED the job packet from " & !idTicketUserFrom
                ElseIf !idTicketAction = "REOPENED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &HFF80FF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = !idTicketUser & " REOPENED the job packet."
                End If
                LineOUT = LineOUT + 1
            ElseIf !idTicketAction = "INTRANSIT" And !idTicketUserTo = strLocalUser Then '**************************************
                Row = Row + 1
                FlexGridIN.TextMatrix(LineIN, 0) = LineIN
                FlexGridIN.TextMatrix(LineIN, 1) = !idTicketJobNum
                FlexGridIN.TextMatrix(LineIN, 2) = !idTicketPartNum
                FlexGridIN.TextMatrix(LineIN, 3) = !idTicketDescription
                FlexGridIN.TextMatrix(LineIN, 4) = !idTicketSalesNum
                FlexGridIN.TextMatrix(LineIN, 5) = !idTicketCustPoNum
                FlexGridIN.TextMatrix(LineIN, 6) = !idTicketCreator
                FlexGridIN.TextMatrix(LineIN, 7) = !idTicketCreateDate
                FlexGridIN.TextMatrix(LineIN, 8) = !idTicketDate
                Call FlexGridRowColor(FlexGridIN, LineIN, &H80FF80)
                FlexGridIN.TextMatrix(LineIN, 9) = !idTicketUserFrom & " SENT the job packet to " & !idTicketUserTo
                LineIN = LineIN + 1
            ElseIf !idTicketStatus = "CLOSED" Then
NextLoop:
            End If
            Row = Row + 1
            rs.MoveNext
        End With
    Loop
    FlexGridOUT.Rows = LineOUT
    FlexGridIN.Rows = LineIN
    rs.Close
    cn.Close
    HideData
    SizeTheSheet FlexGridOUT
    SizeTheSheet FlexGridIN
    FlexGridOUT.Redraw = True
    FlexGridIN.Redraw = True
    FlexGridIN.Visible = True
    FlexGridOUT.Visible = True
    If LineIN <= 1 Then FlexGridIN.Visible = False
    If LineOUT <= 1 Then FlexGridOUT.Visible = False
    FlexGridIN.TopRow = intFlexGridInLastRow
    FlexGridOUT.TopRow = intFlexGridOutLastRow
    SSTab1.TabCaption(3) = "On-hand Packets (" & FlexGridOUT.Rows - 1 & ")"
    SSTab1.TabCaption(2) = "Incoming Packets (" & FlexGridIN.Rows - 1 & ")"
    intPrevInPackets = FlexGridIN.Rows - 1
    If SSTab1.Tab = 2 And ProgHasFocus = True Then
        If Me.ActiveControl.Name <> "SSTab1" Then
            Exit Sub
        ElseIf Me.ActiveControl.Name <> "FlexGridIN" Then
            Exit Sub
        End If
        FlexGridIN.col = FlexINLastSel(1)
        FlexGridIN.Row = FlexINLastSel(0)
        FlexGridIN.ColSel = FlexINLastSel(1)
        FlexGridIN.RowSel = FlexINLastSel(0)
        FlexGridIN.SetFocus
    ElseIf SSTab1.Tab = 3 And ProgHasFocus = True And Me.ActiveControl.Name = "SSTab2" Or Me.ActiveControl.Name = "FlexGridOUT" Then
        If Me.ActiveControl.Name <> "SSTab2" Then
            Exit Sub
        ElseIf Me.ActiveControl.Name <> "FlexGridOUT" Then
            Exit Sub
        End If
        FlexGridOUT.col = FlexOUTLastSel(1)
        FlexGridOUT.Row = FlexOUTLastSel(0)
        FlexGridOUT.ColSel = FlexOUTLastSel(1)
        FlexGridOUT.RowSel = FlexOUTLastSel(0)
        FlexGridOUT.SetFocus
    End If
    Exit Sub
errs:
    If Err.Number = -2147467259 Then
        CommsDown
    Else
        CommsUp
    End If
    Resume Next
End Sub
Public Sub DisplayArrows()
    If FlexGridHist.Rows - FlexGridHist.TopRow <= 14 Then
        picOlder.Visible = False
    Else
        picOlder.Visible = True
    End If
    If FlexGridHist.Rows <= 14 Then picOlder.Visible = False
End Sub
Public Sub ShowData()

    Set pbData.Picture = picDataPics(2)
    DoEvents
    StartTimer
End Sub
Public Sub HideData()
    Dim lngCurQry As Double, lngAddQry As Double, lngAvgQry As Double
    Dim i As Integer
    lngCurQry = StopTimer
    If intQryIndex >= 20 Then
        intQryIndex = 0
        lngQryTimes(intQryIndex) = lngCurQry
    Else
        intQryIndex = intQryIndex + 1
        lngQryTimes(intQryIndex) = lngCurQry
    End If
    For i = 0 To UBound(lngQryTimes)
        lngAddQry = lngAddQry + lngQryTimes(i)
    Next i
    lngAvgQry = lngAddQry / UBound(lngQryTimes)
    lblQryTime.Caption = Round(lngAvgQry, 2) & " ms"
    Set pbData.Picture = picDataPics(0)
End Sub
Public Sub RefreshAll() 'Main Idle Loop - Refreshes Fields and History, only when newer entries are detected. Always refreshes MyPackets.
    Dim b As Integer
    On Error GoTo errs
    If bolRunning = True Then Exit Sub
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim strSQL2, strSQL3 As String
    'strSQL2 = "SELECT * From ticketdatabase Where idTicketJobNum = '" & txtJobNo.Text & "' Order By ticketdatabase.idTicketDate Desc"
    strSQL2 = "SELECT *" & " FROM ticketdb.ticketdatabase ticketdatabase_0" & " WHERE (ticketdatabase_0.idTicketJobNum='" & txtJobNo.Text & "') AND (ticketdatabase_0.idTicketDate>{ts '" & dtLatestHistDate & "'})" & " ORDER BY ticketdatabase_0.idTicketDate"
    strSQL3 = "SELECT * FROM ticketdb.ticketdatabase ticketdatabase_0" & " WHERE (ticketdatabase_0.idTicketAction='CREATED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='RECEIVED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='REOPENED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='INTRANSIT') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') AND (ticketdatabase_0.idTicketUserTo='" & strLocalUser & "')" & " ORDER BY ticketdatabase_0.idTicketDate"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    If txtJobNo.Text = "" Or optCreate.Value = True Or bolHasTicket = False Then GoTo GetMyPackets
    ShowData
    rs.Open strSQL2, cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <= 0 Then
        rs.Close
        GoTo GetMyPackets 'If no new entries, skip history and field updates
    End If
    FlexGridHist.Redraw = False
    FlexGridHist.Visible = False
    With rs
        rs.MoveLast
        dtLatestHistDate = Format$(!idTicketDate, strDBDateTimeFormat)
        rs.MoveFirst
    End With
    FlexUNBoldFirst FlexGridHist
    Call FlexFlipHist("A")
    '
    Do Until rs.EOF
        With rs
            Call FillFlexHist(!idTicketAction, !idTicketStatus, !idTicketComment, !idTicketDate, !idTicketCreator, !idTicketUserFrom, !idTicketUserTo, !idTicketUser, !idGUID)
            rs.MoveNext
        End With
    Loop
    For b = 0 To FlexGridHist.Cols - 1
        FlexGridHist.ColAlignment(b) = flexAlignLeftCenter
    Next b
    FlexGridHist.ColWidth(0) = 1000
    FlexGridHist.ColWidth(1) = 8500
    FlexGridHist.ColWidth(3) = 500
    FlexGridHist.RowHeight(0) = 0
    FlexGridHist.TopRow = FlexHistLastTopRow
    FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
    Call FlexFlipHist("D")
    DisplayArrows
    FlexBoldFirst FlexGridHist
    FlexGridRedrawHeight
    'FlexGridHist.Rows = FlexGridHist.Rows - 1
    FlexGridHist.Redraw = True
    FlexGridHist.Visible = True
GetFields:
    With rs
        rs.MoveLast
        txtPartNoRev.Text = !idTicketPartNum
        txtDrawNoRev.Text = !idTicketDrawingNum
        txtCustPoNo.Text = !idTicketCustPoNum
        txtSalesNo.Text = !idTicketSalesNum
        txtCreator.Text = !idTicketCreator
        txtCreateDate.Text = !idTicketCreateDate
        txtActionDate.Text = !idTicketDate
        strTicketAction = !idTicketAction
        strUserFrom = !idTicketUserFrom
        strUserTo = !idTicketUserTo
        strCurUser = !idTicketUser
        strTicketStatus = !idTicketStatus
        txtTicketAction.Text = !idTicketAction
        txtTicketOwner.Text = !idTicketUser
        txtTicketDescription.Text = !idTicketDescription
        txtTicketStatus.Text = !idTicketStatus
        strPlant = !idTicketPlant
        cmbPlant.Text = strPlant
        If txtJobNo.Text = "" Then
            DisableBoxes
            tmrRefresher.Enabled = False
        Else
            bolHasTicket = True
            tmrRefresher.Enabled = True
            FlexGridHist.Visible = True
        End If
        If !idTicketComment <> "" Then
            strCommentText = !idTicketComment
            tmrScroll.Enabled = True
        Else
            pbScrollBox.Cls
            strCommentText = ""
            tmrScroll.Enabled = False
        End If
    End With
    SetControls
    rs.Close
    '*************************************** GetMyPackets ******************************
GetMyPackets:
    Dim LineIN, LineOUT, Row As Integer
    FlexGridOUT.Clear
    FlexGridOUT.Redraw = False
    FlexGridOUT.Rows = 2
    FlexGridOUT.FixedCols = 1
    FlexGridOUT.FixedRows = 1
    FlexGridIN.Clear
    FlexGridIN.Redraw = False
    FlexGridIN.Rows = 2
    FlexGridIN.FixedCols = 1
    FlexGridIN.FixedRows = 1
    ShowData
    rs.Open strSQL3, cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <= 0 Then
        intPrevInPackets = 0
        SSTab1.TabCaption(3) = "On-hand Packets (0)"
        SSTab1.TabCaption(2) = "Incoming Packets (0)"
        FlexGridOUT.Visible = False
        FlexGridOUT.Redraw = True
        FlexGridIN.Visible = False
        FlexGridIN.Redraw = True
        rs.Close
        cn.Close
        HideData
        FlexGridOUT.Clear
        FlexGridIN.Clear
        GoTo SkipGetMyPackets
    End If
    LineIN = 1
    LineOUT = 1
    Row = 0
    FlexGridOUT.Rows = rs.RecordCount + 1
    FlexGridOUT.Cols = 10
    FlexGridIN.Rows = rs.RecordCount + 1
    FlexGridIN.Cols = 10
    ' Create header row
    FlexGridOUT.TextMatrix(0, 1) = "Job Number"
    FlexGridOUT.TextMatrix(0, 2) = "Part Number"
    FlexGridOUT.TextMatrix(0, 3) = "Description"
    FlexGridOUT.TextMatrix(0, 4) = "Sales Number"
    FlexGridOUT.TextMatrix(0, 5) = "Customer/PO Number"
    FlexGridOUT.TextMatrix(0, 6) = "Created By"
    FlexGridOUT.TextMatrix(0, 7) = "Create Date"
    FlexGridOUT.TextMatrix(0, 8) = "Last Activity Date"
    FlexGridOUT.TextMatrix(0, 9) = "Last Activity"
    FlexGridIN.TextMatrix(0, 1) = "Job Number"
    FlexGridIN.TextMatrix(0, 2) = "Part Number"
    FlexGridIN.TextMatrix(0, 3) = "Description"
    FlexGridIN.TextMatrix(0, 4) = "Sales Number"
    FlexGridIN.TextMatrix(0, 5) = "Customer/PO Number"
    FlexGridIN.TextMatrix(0, 6) = "Created By"
    FlexGridIN.TextMatrix(0, 7) = "Create Date"
    FlexGridIN.TextMatrix(0, 8) = "Last Activity Date"
    FlexGridIN.TextMatrix(0, 9) = "Last Activity"
    Do Until rs.EOF
        With rs
            If !idTicketAction = "CREATED" And !idTicketUser = strLocalUser Or !idTicketAction = "RECEIVED" And !idTicketUser = strLocalUser Or !idTicketAction = "REOPENED" And !idTicketUser = strLocalUser Then
                Row = Row + 1
                FlexGridOUT.TextMatrix(LineOUT, 0) = LineOUT
                FlexGridOUT.TextMatrix(LineOUT, 1) = !idTicketJobNum
                FlexGridOUT.TextMatrix(LineOUT, 2) = !idTicketPartNum
                FlexGridOUT.TextMatrix(LineOUT, 3) = !idTicketDescription
                FlexGridOUT.TextMatrix(LineOUT, 4) = !idTicketSalesNum
                FlexGridOUT.TextMatrix(LineOUT, 5) = !idTicketCustPoNum
                FlexGridOUT.TextMatrix(LineOUT, 6) = !idTicketCreator
                FlexGridOUT.TextMatrix(LineOUT, 7) = !idTicketCreateDate
                FlexGridOUT.TextMatrix(LineOUT, 8) = !idTicketDate
                If !idTicketAction = "CREATED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &H80C0FF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = "Job packet was CREATED by " & !idTicketCreator
                ElseIf !idTicketAction = "RECEIVED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &H80FFFF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = !idTicketUser & " RECEIVED the job packet from " & !idTicketUserFrom
                ElseIf !idTicketAction = "REOPENED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &HFF80FF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = !idTicketUser & " REOPENED the job packet."
                End If
                LineOUT = LineOUT + 1
            ElseIf !idTicketAction = "INTRANSIT" And !idTicketUserTo = strLocalUser Then '**************************************
                Row = Row + 1
                FlexGridIN.TextMatrix(LineIN, 0) = LineIN
                FlexGridIN.TextMatrix(LineIN, 1) = !idTicketJobNum
                FlexGridIN.TextMatrix(LineIN, 2) = !idTicketPartNum
                FlexGridIN.TextMatrix(LineIN, 3) = !idTicketDescription
                FlexGridIN.TextMatrix(LineIN, 4) = !idTicketSalesNum
                FlexGridIN.TextMatrix(LineIN, 5) = !idTicketCustPoNum
                FlexGridIN.TextMatrix(LineIN, 6) = !idTicketCreator
                FlexGridIN.TextMatrix(LineIN, 7) = !idTicketCreateDate
                FlexGridIN.TextMatrix(LineIN, 8) = !idTicketDate
                Call FlexGridRowColor(FlexGridIN, LineIN, &H80FF80)
                FlexGridIN.TextMatrix(LineIN, 9) = !idTicketUserFrom & " SENT the job packet to " & !idTicketUserTo
                LineIN = LineIN + 1
            ElseIf !idTicketStatus = "CLOSED" Then
NextLoop:
            End If
            Row = Row + 1
            rs.MoveNext
        End With
    Loop
    FlexGridOUT.Rows = LineOUT
    FlexGridIN.Rows = LineIN
    rs.Close
    SizeTheSheet FlexGridOUT
    SizeTheSheet FlexGridIN
    FlexGridOUT.Redraw = True
    FlexGridIN.Redraw = True
    FlexGridIN.Visible = True
    FlexGridOUT.Visible = True
    If LineIN <= 1 Then FlexGridIN.Visible = False
    If LineOUT <= 1 Then FlexGridOUT.Visible = False
    If intFlexGridInLastRow >= 2 Then FlexGridIN.TopRow = intFlexGridInLastRow
    If intFlexGridOutLastRow >= 2 Then FlexGridOUT.TopRow = intFlexGridOutLastRow
    SSTab1.TabCaption(3) = "On-hand Packets (" & FlexGridOUT.Rows - 1 & ")"
    SSTab1.TabCaption(2) = "Incoming Packets (" & FlexGridIN.Rows - 1 & ")"
    If FlexGridIN.Rows - 1 > intPrevInPackets Then
        ShowBanner &HC0C0C0, "You have new incoming Job Packets. Click to view.", 500, "VIEWPACK"
        intPrevInPackets = FlexGridIN.Rows - 1
    Else
        intPrevInPackets = FlexGridIN.Rows - 1
    End If
    If SSTab1.Tab = 2 And ActiveControl.Name = "FlexGridIN" And ProgHasFocus = True Then
        FlexGridIN.col = FlexINLastSel(1)
        FlexGridIN.Row = FlexINLastSel(0)
        FlexGridIN.ColSel = FlexINLastSel(1)
        FlexGridIN.RowSel = FlexINLastSel(0)
        FlexGridIN.SetFocus
    ElseIf SSTab1.Tab = 3 And ActiveControl.Name = "FlexGridOUT" And ProgHasFocus = True Then
        FlexGridOUT.col = FlexOUTLastSel(1)
        FlexGridOUT.Row = FlexOUTLastSel(0)
        FlexGridOUT.ColSel = FlexOUTLastSel(1)
        FlexGridOUT.RowSel = FlexOUTLastSel(0)
        FlexGridOUT.SetFocus
    End If
SkipGetMyPackets:
    ' cn.Close
    HideData
    CommsUp
    Exit Sub
errs:
    If Err.Number = -2147467259 Then
        CommsDown
    Else
        ' CommsUp
    End If
    Err.Clear
    'Resume Next
End Sub
Public Sub CommsDown()
    Set pbData.Picture = picDataPics(1)
    optReceive.Enabled = False
    optMove.Enabled = False
    cmbUsers.Visible = False
    lblUser.Visible = False
    optClose.Enabled = False
    optCreate.Enabled = False
    optReOpen.Enabled = False
    optFile.Enabled = False
    bolCanEdit = False
    cmdSubmit.Enabled = False
    cmdRefreshHist.Enabled = False
    cmdTimeLine.Enabled = False
    cmdFilterReport.Enabled = False
    cmdAllOpenReport.Enabled = False
    cmdAllClosedReport.Enabled = False
    cmdGetInBox.Enabled = False
    cmdGetOutBox.Enabled = False
    If bolMessageDelivered = False Then
        StatusBar1.Panels.Item(1).Text = "Cannot communicate with server! Program suspended until the server has been detected."
        ShowBanner vbRed, "! The program has lost the connection to the server. Packet Tracker has gone into suspend mode. !", 500, , vbWhite
        bolMessageDelivered = True
    Else
    End If
End Sub
Public Sub CommsUp()
    If bolMessageDelivered = True Then
        StatusBar1.Panels.Item(1).Text = ""
        ShowBanner vbGreen, "Connection restored!", 250
        bolMessageDelivered = False
    Else
    End If
    Set pbData.Picture = picDataPics(0)
    cmdRefreshHist.Enabled = True
    cmdTimeLine.Enabled = True
    cmdFilterReport.Enabled = True
    cmdAllOpenReport.Enabled = True
    cmdAllClosedReport.Enabled = True
    cmdGetInBox.Enabled = True
    cmdGetOutBox.Enabled = True
End Sub
Public Sub FlexGridRedrawHeight()
    Dim ColLoop As Long
    Dim RowLoop As Long
    'Turn off redrawing to avoid flickering
    FlexGridHist.Redraw = False
    'For ColLoop = 0 To FlexGridHist.Cols - 1
    'FlexGridHist.ColWidth(ColLoop) = 2500
    For RowLoop = 0 To FlexGridHist.Rows - 1
        ReSizeCellHeight RowLoop, 1
    Next RowLoop
    'Next ColLoop
    'Turn redrawing back on
    FlexGridHist.Redraw = True
End Sub
Public Sub ReSizeCellHeight(MyRow As Long, MyCol As Long)
    Dim LinesOfText  As Long
    Dim HeightOfLine As Long
    On Error Resume Next
    'Set MSFlexGrid to appropriate Cell
    FlexGridHist.Row = MyRow
    FlexGridHist.col = MyCol
    'Set textbox width to match current width of selected cell
    Text1.Width = FlexGridHist.ColWidth(MyCol) - 100
    Text1.FontSize = FlexGridHist.CellFontSize
    Text1.FontBold = FlexGridHist.CellFontBold
    Text1.FontItalic = FlexGridHist.CellFontItalic
    Text1.Text = FlexGridHist.Text
    'Get the height of the text in the textbox
    HeightOfLine = Me.TextHeight(Text1.Text) + 50 '285
    'Call API to determine how many lines of text are in text box
    LinesOfText = SendMessage(Text1.hwnd, EM_GETLINECOUNT, 0&, 0&)
    'Check to see if row is not tall enough
    If FlexGridHist.RowHeight(MyRow) < (LinesOfText * HeightOfLine) Then
        'Adjust the RowHeight based on the number of lines in textbox
        FlexGridHist.RowHeight(MyRow) = LinesOfText * HeightOfLine + 200
    End If
End Sub
Private Sub SetBoxesForEdit(EnabledControl As String)
    If EnabledControl <> "All" Then
        If EnabledControl = "txtTicketDescription" Then
            txtTicketDescription.Enabled = True
        Else
            txtTicketDescription.Enabled = False
        End If
        If EnabledControl = "txtPartNoRev" Then
            txtPartNoRev.Enabled = True
        Else
            txtPartNoRev.Enabled = False
        End If
        If EnabledControl = "txtSalesNo" Then
            txtSalesNo.Enabled = True
        Else
            txtSalesNo.Enabled = False
        End If
        If EnabledControl = "txtDrawNoRev" Then
            txtDrawNoRev.Enabled = True
        Else
            txtDrawNoRev.Enabled = False
        End If
        If EnabledControl = "txtCustPoNo" Then
            txtCustPoNo.Enabled = True
        Else
            txtCustPoNo.Enabled = False
        End If
    ElseIf EnabledControl = "All" Then
        txtTicketDescription.Enabled = True
        txtPartNoRev.Enabled = True
        txtSalesNo.Enabled = True
        txtDrawNoRev.Enabled = True
        txtCustPoNo.Enabled = True
        cmdEdit.Visible = False
        cmdEdit.Picture = ButtonPics(1)
        cmdEdit.ToolTipText = "Edit Field"
        DisableBoxes
        EditMode = False
    End If
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errs
    Dim blah
    Dim EditedFieldPrev, EditedFieldName, EditedFieldCurr As String
    If EditMode = False Then
        PrevPartNum = UCase$(txtPartNoRev)
        PrevDrawNoRev = UCase$(txtDrawNoRev)
        PrevCustPoNo = UCase$(txtCustPoNo)
        PrevSalesNo = UCase$(txtSalesNo)
        PrevDescription = UCase$(txtTicketDescription)
        Me.Controls(ActiveText).Locked = False
        cmdEdit.Picture = ButtonPics(2)
        cmdEdit.ToolTipText = "Confirm Changes"
        EditMode = True
        SetControls
        PositionMaxChar Me.Controls(ActiveText)
        With Me.Controls(ActiveText)
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        SetBoxesForEdit ActiveText
    Else
        '**** Do stuff that updates the database
        If CheckForBlanks(ActiveText) = True Then
            ShowBanner &H8080FF, "This field cannot be left blank! Please fill the field and try again.", 300
            Me.Controls(ActiveText).BackColor = &H8080FF
            Exit Sub
        End If
        If ChangesMade = False Then
            ShowBanner &H8080FF, "No changes detected! Job Packet was not updated.", 250
            Form1.cmdSubmit.Enabled = False
            Form1.optMove.Value = False
            Form1.optReceive.Value = False
            Form1.optMove.Value = False
            Form1.optClose.Value = False
            Form1.optCreate.Value = False
            Form1.optReOpen.Value = False
            Form1.optFile.Value = False
            bolOptionClicked = False
            imgComment.Picture = ButtonPics(4)
            imgComment.Enabled = False
            'RefreshAfterEdit
            cmdEdit.Visible = False
            cmdEdit.Picture = ButtonPics(1)
            cmdEdit.ToolTipText = "Edit Field"
            EditMode = False
            PositionMaxChar Me.Controls(ActiveText)
            SetBoxesForEdit "All"
            DisableBoxes
            SetControls
            Exit Sub
        End If
        Dim rs      As New ADODB.Recordset
        Dim cn      As New ADODB.Connection
        Dim strSQL1 As String
        strSQL1 = "SELECT * From ticketdatabase Where idTicketJobNum = '" & txtJobNo.Text & "' Order By ticketdatabase.idTicketDate Desc"
        Set rs = New ADODB.Recordset
        Set cn = New ADODB.Connection
        cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
        cn.CursorLocation = adUseClient
        rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
        Do Until rs.EOF
            With rs
                If ActiveText = "txtPartNoRev" And txtPartNoRev <> PrevPartNum Then !idTicketPartNum = UCase$(txtPartNoRev.Text)
                If ActiveText = "txtDrawNoRev" And txtDrawNoRev <> PrevDrawNoRev Then !idTicketDrawingNum = UCase$(txtDrawNoRev.Text)
                If ActiveText = "txtCustPoNo" And txtCustPoNo <> PrevCustPoNo Then !idTicketCustPoNum = UCase$(txtCustPoNo.Text)
                If ActiveText = "txtSalesNo" And txtSalesNo <> PrevSalesNo Then !idTicketSalesNum = UCase$(txtSalesNo.Text)
                If ActiveText = "txtTicketDescription" And txtTicketDescription <> PrevDescription Then !idTicketDescription = txtTicketDescription.Text
                rs.Update
                rs.MoveNext
            End With
        Loop
        rs.Close
        cn.Close
        Form1.cmdSubmit.Enabled = False
        Form1.optMove.Value = False
        Form1.optReceive.Value = False
        Form1.optMove.Value = False
        Form1.optClose.Value = False
        Form1.optCreate.Value = False
        Form1.optReOpen.Value = False
        Form1.optFile.Value = False
        bolOptionClicked = False
        imgComment.Picture = ButtonPics(4)
        imgComment.Enabled = False
        RefreshAll
        cmdEdit.Visible = False
        cmdEdit.Picture = ButtonPics(1)
        cmdEdit.ToolTipText = "Edit Field"
        EditMode = False
        SetBoxesForEdit "All"
        DisableBoxes
        SetControls
        ShowBanner &HC0FFC0, "Job Packet was updated successfully!", 200
    End If
    Exit Sub
errs:
    EditMode = False
    SetBoxesForEdit "All"
    DisableBoxes
    blah = MsgBox("An error was detected!" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Yikes!")
    Err.Clear
End Sub
Private Sub cmdPrintInPack_Click()
    If FlexGridIN.Rows < 2 Then
        MsgBox ("Nothing to print!")
        Exit Sub
    End If
    frmPrinters.Show 1
    If bolCancelPrint = True Then
        bolCancelPrint = False
        Exit Sub
    End If
    FlexGridIN.ColWidth(1) = 1005
    FlexGridIN.ColWidth(2) = 1005
    FlexGridIN.ColWidth(3) = 2715
    FlexGridIN.ColWidth(4) = 930
    FlexGridIN.ColWidth(5) = 1170
    FlexGridIN.ColWidth(6) = 885
    FlexGridIN.ColWidth(7) = 1335
    FlexGridIN.ColWidth(8) = 1290
    FlexGridIN.ColWidth(9) = 3525
    strReportType = "Incoming Packets"
    PrintFlexGrid FlexGridIN
    SizeTheSheet FlexGridIN
End Sub
Private Sub cmdPrintOnPack_Click()
    If FlexGridOUT.Rows < 2 Then
        MsgBox ("Nothing to print!")
        Exit Sub
    End If
    frmPrinters.Show 1
    If bolCancelPrint = True Then
        bolCancelPrint = False
        Exit Sub
    End If
    FlexGridOUT.ColWidth(1) = 1005
    FlexGridOUT.ColWidth(2) = 1005
    FlexGridOUT.ColWidth(3) = 2715
    FlexGridOUT.ColWidth(4) = 930
    FlexGridOUT.ColWidth(5) = 1170
    FlexGridOUT.ColWidth(6) = 885
    FlexGridOUT.ColWidth(7) = 1335
    FlexGridOUT.ColWidth(8) = 1290
    FlexGridOUT.ColWidth(9) = 3525
    strReportType = "On-hand Packets"
    PrintFlexGrid FlexGridOUT
    SizeTheSheet FlexGridOUT
End Sub
Private Sub cmdTimeLine_Click()
    On Error Resume Next
    If bolHasTicket = False Then Exit Sub
    Load frmTip
    frmTip.Show
    Do While HelpClosed = False
        Sleep 10
        DoEvents
    Loop
    GetTimeLineData
    DrawDayLines = True
    frmTimeLine.chkDayLines.Value = 1
    DrawTimeLine
    'frmTimeLine.tmrActionShow.Enabled = True
    frmTimeLine.Show
End Sub
Private Sub cmdHeatMap_Click()
    If bolRunning = True Then 'if already running the qry, dont try to start another one. (Prevents server flooding if return key is held down)
        Exit Sub
    Else
        ClearBanners
        ShowAllOpenHeatMap
    End If
End Sub
Private Sub GetFadeColor()
    Dim FadeColor As Long
    Dim Color1, Color2
    FadeColor = GetRealColor(Frame3.BackColor)
    ColorCodeToRGB FadeColor, iRed, iGreen, iBlue
    Color1 = RGB(iRed, iGreen, iBlue)
    r1 = Color1 And (Not &HFFFFFF00)
    g1 = (Color1 And (Not &HFFFF00FF)) \ &H100&
    b1 = (Color1 And (Not &HFF00FFFF)) \ &HFFFF&
    FadeColor = GetRealColor(shpButtonFlash.BackColor)
    ColorCodeToRGB FadeColor, iRed, iGreen, iBlue
    Color2 = RGB(iRed, iGreen, iBlue)
    r2 = Color2 And (Not &HFFFFFF00)
    g2 = (Color2 And (Not &HFFFF00FF)) \ &H100&
    b2 = (Color2 And (Not &HFF00FFFF)) \ &HFFFF&
End Sub
Private Sub FlexGrid1_Click()
    On Error Resume Next
    Set WhichGrid = Flexgrid1
    If strSortMode = "A" Then
        Call FlexSort("D")
        strSortMode = "D"
    ElseIf strSortMode = "D" Then
        Call FlexSort("A")
        strSortMode = "A"
    End If
End Sub
Sub FlexBoldFirst(FlexGrid As MSHFlexGrid)
    Dim intCellHeight As Integer
    On Error Resume Next 'GoTo errs
    intCellHeight = 600
    'FlexGrid.Row = FlexGrid.TopRow
    'FlexGrid.col = 0
    'FlexGrid.CellFontSize = 10.75
    'FlexGrid.CellFontBold = True
    'FlexGrid.CellFontWidth = 6
    FlexGrid.Row = 0
    FlexGrid.col = 1
    FlexGrid.CellFontSize = 10
    FlexGrid.CellFontBold = True
    'FlexGrid.CellFontWidth = 6
    'FlexGrid.CellAlignment = flexAlignCenterCenter
    If FlexGrid.TextMatrix(1, 4) = "com" Then
        FlexGrid.Row = 1
        FlexGrid.col = 1
        FlexGrid.CellFontBold = True
        'FlexGrid.CellAlignment = flexAlignCenterCenter
        FlexGrid.CellFontSize = 10.75
        FlexGrid.RowHeight(1) = intCellHeight - 200
    End If
    Exit Sub
errs:
    If Err.Number = 381 Then FlexGrid.RowHeight(0) = intCellHeight 'if Subscript out of range, it most likely means the grid only has one row. Therefor, no comment, it should fail and finish setting grid height
End Sub
Sub FlexUNBoldFirst(FlexGrid As MSHFlexGrid)
    On Error Resume Next
    FlexGrid.RowHeight(0) = intRowH
    FlexGrid.Row = 0
    FlexGrid.col = 1
    FlexGrid.CellFontSize = 10
    FlexGrid.CellFontBold = False
    FlexGrid.CellFontItalic = False
    FlexGrid.CellAlignment = flexAlignLeftCenter
    If FlexGrid.TextMatrix(1, 4) = "com" Then
        FlexGrid.Row = 1
        FlexGrid.col = 1
        FlexGrid.CellFontBold = False
        FlexGrid.CellAlignment = flexAlignLeftCenter
        FlexGrid.CellFontSize = 10
        FlexGrid.RowHeight(1) = intRowH
    End If
End Sub
Sub FlexFlipHist(Mode As String)
    If Mode = "A" Then
        FlexGridHist.col = 3
        FlexGridHist.Sort = flexSortGenericAscending
    Else
        'do nothing
    End If
    If Mode = "D" Then
        FlexGridHist.col = 3
        FlexGridHist.Sort = flexSortGenericDescending
    Else
        'do nothing
    End If
End Sub
Sub FlexSort(Mode As String)
    If Flexgrid1.MouseRow = 0 And Mode = "A" Then
        Flexgrid1.col = Flexgrid1.MouseCol
        If Flexgrid1.col = 10 Then
            Flexgrid1.Sort = flexSortGenericAscending
        Else
            Flexgrid1.Sort = flexSortStringAscending
        End If
    Else
        'do nothing
    End If
    If Flexgrid1.MouseRow = 0 And Mode = "D" Then
        Flexgrid1.col = Flexgrid1.MouseCol
        If Flexgrid1.col = 10 Then
            Flexgrid1.Sort = flexSortGenericDescending
        Else
            Flexgrid1.Sort = flexSortStringDescending
        End If
    Else
        'do nothing
    End If
End Sub
Public Sub FlexGridRowColor(FlexGrid As MSHFlexGrid, _
                            ByVal lngRow As Long, _
                            ByVal lngColor As Long)
    Dim lngPrevCol       As Long
    Dim lngPrevColSel    As Long
    Dim lngPrevRow       As Long
    Dim lngPrevRowSel    As Long
    Dim lngPrevFillStyle As Long
    If lngRow > FlexGrid.Rows - 1 Then
        Exit Sub
    End If
    lngPrevCol = FlexGrid.col
    lngPrevRow = FlexGrid.Row
    lngPrevColSel = FlexGrid.ColSel
    lngPrevRowSel = FlexGrid.RowSel
    lngPrevFillStyle = FlexGrid.FillStyle
    FlexGrid.col = FlexGrid.FixedCols
    FlexGrid.Row = lngRow
    FlexGrid.ColSel = FlexGrid.Cols - 1
    FlexGrid.RowSel = lngRow
    FlexGrid.FillStyle = flexFillRepeat
    FlexGrid.CellBackColor = lngColor
    FlexGrid.col = lngPrevCol
    FlexGrid.Row = lngPrevRow
    FlexGrid.ColSel = lngPrevColSel
    FlexGrid.RowSel = lngPrevRowSel
    FlexGrid.FillStyle = lngPrevFillStyle
End Sub
Public Sub PrintFlexGrid(FlexGrid As MSHFlexGrid)
    Dim sMsg As String
    Dim HWidth, HHeight As Integer
    Dim PrevX, PrevY As Integer
    On Error Resume Next
    Printer.ScaleMode = 1
    Printer.Orientation = vbPRORLandscape
    With Printer
        .ScaleMode = 1
        Printer.Print
        .FontSize = 20
        sMsg = strReportType
        HWidth = Printer.TextWidth(sMsg) / 2 ' Get one-half width.
        HHeight = Printer.TextHeight(sMsg) / 2 ' Get one-half height.
        Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
        'Printer.CurrentY = Printer.ScaleHeight - 2000 ' / 2 - HHeight
        Printer.Print sMsg
        Printer.FontSize = 8
        Printer.Print "      " & sAddlMsg
    End With
    Printer.FontSize = 7
    Printer.Print ""
    Printer.Print ""
    Printer.Print "    Report date: " & Date & " " & Time & "      Printed by: " & UCase$(Environ$("USERNAME"))
    Const GAP = 20
    Dim xmax, xmin As Single
    xmin = 200
    xmax = 9500 '10000
    Dim ymax, ymin As Single
    ymin = 2000
    ymax = 11000
    Dim X As Single
    Dim c As Integer
    Dim R As Integer
    With Printer.Font
        .Name = FlexGrid.Font.Name
        .Size = 7
    End With
    With FlexGrid
        ' See how wide the whole thing is.
        xmax = xmin + GAP
        For c = 0 To .Cols - 1
            xmax = xmax + .ColWidth(c) + 2 * GAP
        Next c
        ' Print each row.
        Printer.CurrentY = ymin
        For R = 0 To .Rows - 1
            ' Draw a line above this row.
            If R > 0 Then Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Print the entries on this row.
            X = xmin + GAP
            For c = 0 To .Cols - 1
                Printer.CurrentX = X
                Printer.Print BoundedText(Printer, .TextMatrix(R, c), .ColWidth(c));
                X = X + .ColWidth(c) + 2 * GAP
            Next c
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Move to the next line.
            Printer.Print
            '            PrevX = Printer.CurrentX
            '            PrevY = Printer.CurrentY
            '
            '            Printer.CurrentX = xmax
            '            Printer.CurrentY = ymax + 200
            '
            '            Printer.Print "Page " & Printer.Page
            '
            '            Printer.CurrentX = PrevX
            '            Printer.CurrentY = PrevY
            '
            ' if near end of page, start a new one
            If Printer.CurrentY >= ymax Then
                Printer.Line (xmin, ymin)-(xmax, Printer.CurrentY), , B
                X = xmin
                For c = 0 To .Cols - 2
                    X = X + .ColWidth(c) + 2 * GAP
                    Printer.Line (X, ymin)-(X, Printer.CurrentY)
                Next c
                Printer.NewPage
                Printer.CurrentX = xmax
                Printer.CurrentY = ymax + 200
                Printer.Print "Page " & Printer.Page
                Printer.CurrentX = xmin
                ymin = 400
                Printer.CurrentY = ymin
            End If
        Next R
        ymax = Printer.CurrentY
        ' Draw a box around everything.
        Printer.Line (xmin, ymin)-(xmax, ymax), , B
        ' Draw lines between the columns.
        X = xmin
        For c = 0 To .Cols - 2
            X = X + .ColWidth(c) + 2 * GAP
            Printer.Line (X, ymin)-(X, Printer.CurrentY)
        Next c
    End With
    Printer.EndDoc
End Sub
Private Sub PrintFlexGridColor(FlexGrid As MSHFlexGrid)
    On Error Resume Next
    FlexGrid.Redraw = False
    'Dim sMsg As String
    Dim intPadding    As Integer
    Dim PrevX         As Integer, PrevY As Integer, intMidStart As Integer, intMidLen As Integer, intTotLen As Integer
    Dim strSizedTxt   As String, strOrigTxt As String
    Dim bolLongLine   As Boolean
    Dim TwipPix       As Long
    Dim lngYTopOfGrid As Long
    bolLongLine = False
    Dim sMsg            As String
    Dim intCenterOffset As Long
    Dim lngStartY       As Long, lngStartX As Long, lngEndX As Long, lngEndY As Long
    Dim xmax            As Single, xmin As Single
    xmin = 300
    xmax = 14800
    Dim ymax As Single, ymin As Single
    ymin = 1500
    ymax = 10800
    Printer.Font.Underline = False
    Printer.ScaleMode = 1
    Printer.Orientation = vbPRORLandscape
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbSolid
    sMsg = strReportType
    Printer.Print sMsg
    Printer.FontSize = 8
    Printer.Print "      " & sAddlMsg
    With Printer
        .ScaleMode = 1
        Printer.Print
        .FontSize = 20
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(sMsg) / 2)
        Printer.Print sMsg
        Printer.FontSize = 8
    End With
    Printer.FontSize = 7
    ' Printer.Print "    " & strReportMsg
    'Printer.Print ""
    Printer.Print "    Report date: " & Date & " " & Time & "      Printed by: " & UCase$(Environ$("USERNAME"))
    Const GAP = 40
    With Printer.Font
        .Name = FlexGrid.Font.Name
        .Size = 9
    End With
    Printer.Print ""
    Printer.DrawStyle = vbDash
    Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
    Printer.DrawStyle = vbSolid
    Printer.Print ""
    PrevY = Printer.CurrentY
    '    Dim xBoxEnd As Single, lngCenterXStartPos As Long
    '    Printer.Font.Size = 7
    '    lngCenterXStartPos = (xmax / 2) - (Printer.TextWidth(strReportInfo) / 2)
    '    xBoxEnd = lngCenterXStartPos + Printer.TextWidth(strReportInfo)
    '    Printer.Line (lngCenterXStartPos, PrevY)-(xBoxEnd, Printer.CurrentY + (Printer.TextHeight(strReportInfo) * 3)), &H80000016, BF
    '    Printer.Font.Bold = True
    '    Printer.CurrentX = (xmax / 2) - (Printer.TextWidth("Attendance Stats") / 2)
    '    Printer.CurrentY = PrevY
    '    Printer.Print "Attendance Stats"
    '    Printer.Font.Bold = False
    '    Printer.CurrentX = lngCenterXStartPos
    '    Printer.Print strReportInfo
    '    Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(strReportEntryCount) / 2)
    '    Printer.Print strReportEntryCount
    '    Printer.Line (lngCenterXStartPos, PrevY)-(xBoxEnd, Printer.CurrentY), vbBlack, B
    '    Printer.Print ""
    Printer.Font.Size = 9
    Printer.DrawStyle = vbSolid
    Dim X As Single, XFirstColumn As Single
    Dim c As Integer, cc As Integer
    Dim R As Integer
    intMidStart = 1
    With FlexGrid
        PrevX = Printer.CurrentX
        PrevY = Printer.CurrentY
        Printer.CurrentX = xmax - 600
        Printer.CurrentY = ymax + 300
        Printer.ForeColor = vbBlack
        Printer.Font.Underline = False
        Printer.Print "Page " & Printer.Page
        Printer.CurrentX = PrevX
        Printer.CurrentY = PrevY
        intPadding = 150
        TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX
        XFirstColumn = xmin + TwipPix * GAP
        lngYTopOfGrid = Printer.CurrentY
        Printer.CurrentY = Printer.CurrentY + GAP
        'If FlexGrid.Header = True Then
        X = xmin + GAP
        For c = 1 To .Cols
            Printer.CurrentX = X
            TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX
            PrevY = Printer.CurrentY
            If c = .Cols Then
                lngStartY = Printer.CurrentY - GAP + 5
                lngStartX = Printer.CurrentX - GAP + 5
                lngEndX = xmax
                lngEndY = Printer.CurrentY + Printer.TextHeight(.ColHeader(c)) + GAP
                Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), &H80000016, BF
            Else
                lngStartY = Printer.CurrentY - GAP + 5
                lngStartX = Printer.CurrentX - GAP + 5
                lngEndX = Printer.CurrentX + TwipPix + GAP
                lngEndY = Printer.CurrentY + Printer.TextHeight(.ColHeader(c)) + GAP
                Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), &H80000016, BF
            End If
            Printer.CurrentX = lngStartX + GAP
            Printer.CurrentY = PrevY
            Printer.Print BoundedText(Printer, .ColHeader(c), TwipPix);
            TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX
            X = X + TwipPix + 2 * GAP
        Next c
        Printer.CurrentY = Printer.CurrentY + GAP
        Printer.Print
        ' End If
        For R = 1 To .Rows '- 1
            'If bolStop = True Then
            ' Printer.EndDoc
            ' bolStop = False
            ' frmpBar.Visible = False
            ' Exit Sub
            'End If
            ' Draw a line above this row.
            If R > 0 Then
                Printer.Line (XFirstColumn, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
            End If
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Print the entries on this row.
            X = xmin + GAP
            For c = 1 To .Cols ' - 1
                If frmPrinters.optCenterJust And c < .Cols Then
                    intCenterOffset = ((.ColWidth(c) * Screen.TwipsPerPixelX) / 2) - (Printer.TextWidth(.TextMatrix(R, c)) / 2)
                Else
                    intCenterOffset = 0
                End If
                Printer.CurrentX = X
                If .TextMatrix(R, c) <> "" And Printer.TextWidth(.TextMatrix(R, c)) + intPadding >= xmax - Printer.CurrentX Then           '.ColWidth(c)
                    lngStartY = Printer.CurrentY + Printer.TextHeight(.TextMatrix(R, c))
                    strOrigTxt = .TextMatrix(R, c)
                    Do Until intTotLen >= Len(strOrigTxt)
                        Do Until Printer.TextWidth(strSizedTxt) + intPadding >= xmax - Printer.CurrentX Or intTotLen >= Len(strOrigTxt)
                            intMidLen = intMidLen + 1
                            intTotLen = intTotLen + 1
                            strSizedTxt = Mid$(strOrigTxt, intMidStart, intMidLen)
                        Loop
                        intMidStart = intMidStart + intMidLen ' - 1
                        intMidLen = 1
                        'Printer.Font.Underline = .CellFontUnderline(R, c).Underline
                        '                        If .CellFontUnderline(R, c).Underline = True Then
                        '                            Printer.ForeColor = vbBlack
                        '                        Else
                        '                            Printer.ForeColor = &H404040
                        '                        End If
                        Printer.Print strSizedTxt 'Left$(.CellText (R, c), i)
                        lngEndY = Printer.CurrentY + GAP
                        PrevY = Printer.CurrentY
                        .Row = R
                        .ColSel = .Cols - 1
                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor, BF
                        Printer.CurrentY = PrevY + 5
                        If Printer.CurrentY >= ymax Then ' new page
                            Printer.Line (XFirstColumn, lngYTopOfGrid)-(xmax, Printer.CurrentY + GAP), vbBlack, B
                            X = xmin
                            For cc = 1 To .Cols - 1
                                TwipPix = .ColWidth(cc) * Screen.TwipsPerPixelX
                                X = X + TwipPix + 2 * GAP
                                Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack
                            Next cc
                            Printer.NewPage
                            Printer.CurrentX = xmax - 600
                            Printer.CurrentY = ymax + 300
                            Printer.ForeColor = vbBlack
                            Printer.Font.Underline = False
                            Printer.Print "Page " & Printer.Page
                            Printer.CurrentX = xmin
                            ymin = 400
                            lngYTopOfGrid = ymin
                            Printer.CurrentY = ymin
                            lngStartY = Printer.CurrentY '+ Printer.TextHeight(.CellText(R, c))
                        End If
                        Printer.CurrentX = X + GAP
                        strSizedTxt = ""
                    Loop
                    intMidStart = 1
                    intMidLen = 0
                    intTotLen = 0
                    strSizedTxt = ""
                    bolLongLine = True
                Else
                    'bolLongLine = False
                    PrevY = Printer.CurrentY - GAP ' + 10
                    '                    If c = 3 Then
                    '                        lngStartY = Printer.CurrentY - GAP + 5
                    '                        lngStartX = Printer.CurrentX - GAP + 5
                    '                        lngEndX = Printer.CurrentX + .ColWidth(c) * Screen.TwipsPerPixelX + GAP ' - 10
                    '                        lngEndY = Printer.CurrentY + Printer.TextHeight(.TextMatrix(R, c)) + GAP - 5
                    '                        .Row = R
                    '                        .ColSel = .Cols - 1
                    '                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor, BF
                    '                    End If
                    Printer.CurrentX = X + intCenterOffset
                    TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX
                    '                    Printer.Font.Underline = .CellFontUnderline(R, c)
                    '                    If .CellFontUnderline(R, c) = True Then
                    '                        Printer.ForeColor = vbBlack
                    '                    Else
                    '                        Printer.ForeColor = &H404040   '&H808080
                    '                    End If
                    Printer.CurrentX = X + intCenterOffset
                    Printer.CurrentY = PrevY + GAP
                    Printer.Print BoundedText(Printer, .TextMatrix(R, c), TwipPix);
                End If
                TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX
                X = X + TwipPix + 2 * GAP
            Next c
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Move to the next line.
            If bolLongLine = True Then
                bolLongLine = False
            Else
                Printer.Print
                bolLongLine = False
            End If
            ' if near end of page, start a new one
            If Printer.CurrentY >= ymax And R < .Rows Then
                Printer.Line (XFirstColumn, lngYTopOfGrid)-(xmax, Printer.CurrentY), vbBlack, B
                X = xmin
                For c = 1 To .Cols - 1 '3
                    TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX '+ GAP
                    X = X + TwipPix + 2 * GAP
                    Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack 'ymax
                Next c
                Printer.NewPage
                Printer.CurrentX = xmax - 600
                Printer.CurrentY = ymax + 300
                Printer.ForeColor = vbBlack
                Printer.Font.Underline = False
                Printer.Print "Page " & Printer.Page
                Printer.CurrentX = xmin
                ymin = 400
                lngYTopOfGrid = ymin
                Printer.CurrentY = ymin
            End If
        Next R
        ymax = Printer.CurrentY
        'Draw a box around everything.
        Printer.Line (XFirstColumn, lngYTopOfGrid)-(xmax, ymax), vbBlack, B
        X = xmin
        ' Draw lines between the columns.
        For c = 1 To .Cols - 1 '3
            TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX
            X = X + TwipPix + 2 * GAP
            'vbBlack
            Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack 'Printer.CurrentY
        Next c
    End With
End Sub
Private Function BoundedText(ByVal ptr As Object, _
                             ByVal txt As String, _
                             ByVal max_wid As Single) As String
    Do While Printer.TextWidth(txt) > max_wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    BoundedText = txt
End Function
Public Sub SizeTheSheet(TargetGrid As MSHFlexGrid)
    On Error Resume Next
    Dim z, Y As Integer
    z = 1
    Y = 600
    TargetGrid.ScrollBars = flexScrollBarNone
    Dim col(), i, b As Integer
    ReDim col(TargetGrid.Cols)
    For i = 0 To TargetGrid.Rows - 1
        For b = 0 To TargetGrid.Cols - 1
            If TextWidth(TargetGrid.TextMatrix(i, b)) > col(b) Then col(b) = TextWidth(TargetGrid.TextMatrix(i, b))
        Next b
    Next i
    For b = 0 To TargetGrid.Cols - 1
        If b = 4 Then
            TargetGrid.ColWidth(b) = (col(b) * z) + Y
        Else
            TargetGrid.ColWidth(b) = (col(b) * z) + Y
        End If
        TargetGrid.ColAlignment(b) = flexAlignLeftCenter
    Next b
    TargetGrid.ScrollBars = flexScrollBarBoth
    TargetGrid.ColWidth(0) = 0
End Sub
Public Sub RefreshHistory() 'Redraws History Grid
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim b       As Integer
    On Error GoTo errs
    If Me.ActiveControl.Name = "FlexGridHist" Then Exit Sub
    If txtJobNo.Text = "" Then Exit Sub
    If bolHasTicket = False Then Exit Sub
    ShowData
    strSQL1 = "SELECT * From ticketdatabase Where idTicketJobNum = '" & txtJobNo.Text & "' Order By ticketdatabase.idTicketDate Desc"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    FlexGridHist.Redraw = False
    FlexGridHist.Visible = False
    FlexGridHist.Clear
    FlexGridHist.Cols = 6
    FlexGridHist.Rows = 0
    'FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
    'FlexGridHist.col = 1
    With rs
        dtLatestHistDate = Format$(!idTicketDate, strDBDateTimeFormat)
    End With
    rs.MoveLast
    Do Until rs.BOF
        With rs
            Call FillFlexHist(!idTicketAction, !idTicketStatus, !idTicketComment, !idTicketDate, !idTicketCreator, !idTicketUserFrom, !idTicketUserTo, !idTicketUser, !idGUID)
            rs.MovePrevious
        End With
    Loop
    For b = 0 To FlexGridHist.Cols - 1
        FlexGridHist.ColAlignment(b) = flexAlignLeftCenter
    Next b
    FlexGridHist.ColWidth(0) = 1000
    FlexGridHist.ColWidth(1) = 8500
    FlexGridHist.ColWidth(3) = 0
    FlexGridHist.ColWidth(4) = 0
    FlexGridHist.RowHeight(0) = 0
    FlexGridHist.TopRow = FlexHistLastTopRow
    FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
    Call FlexFlipHist("D")
    DisplayArrows
    FlexBoldFirst FlexGridHist
    FlexGridRedrawHeight
    FlexGridHist.Redraw = True
    FlexGridHist.Visible = True
    rs.Close
    cn.Close
    HideData
    Exit Sub
errs:
    If Err.Number = -2147467259 Then
        CommsDown
    Else
        CommsUp
        MsgBox (Err.Number & " - " & Err.Description)
    End If
End Sub
Public Sub SetControls()
    If strTicketAction = "FILED" And strCurUser <> strLocalUser Then
        optReceive.Enabled = False
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = True
        optFile.Enabled = False
        bolCanEdit = False
        StatusBar1.Panels.Item(1).Text = "This packet has been Filed by " & strCurUser & ". Please re-open the packet if on hand"
        Exit Sub
    ElseIf strTicketAction = "FILED" And strCurUser = strLocalUser Then
        optReceive.Enabled = False
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = True
        optFile.Enabled = False
        StatusBar1.Panels.Item(1).Text = "This packet is currently Filed. Filing location should be noted in the latest comment."
        bolCanEdit = False
        Exit Sub
    End If
    If strTicketAction = "REOPENED" And strCurUser <> strLocalUser Then
        optReceive.Enabled = False
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = False
        bolCanEdit = False
        StatusBar1.Panels.Item(1).Text = strCurUser & " has reopened this packet and currently has it on hand."
    ElseIf strTicketAction = "REOPENED" And strCurUser = strLocalUser Then
        optReceive.Enabled = False
        optMove.Enabled = True
        optClose.Enabled = True
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = True
        StatusBar1.Panels.Item(1).Text = "You have Reopened this job packet and now have it on hand."
        bolCanEdit = True
    End If
    If strTicketAction = "INTRANSIT" And strUserTo <> strLocalUser Then
        optReceive.Enabled = False
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = False
        bolCanEdit = False
        StatusBar1.Panels.Item(1).Text = "This packet is in transit to " & strUserTo & "."
    ElseIf strTicketAction = "INTRANSIT" And strUserTo = strLocalUser Then
        optReceive.Enabled = True
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = False
        bolCanEdit = False
        StatusBar1.Panels.Item(1).Text = "Job packet ready to be Received!"
    End If
    If strTicketAction = "RECEIVED" And strCurUser <> strLocalUser Then
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optReceive.Enabled = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = False
        bolCanEdit = False
        optFile.Enabled = False
        StatusBar1.Panels.Item(1).Text = strCurUser & " currently has this packet onhand."
    ElseIf strTicketAction = "RECEIVED" And strCurUser = strLocalUser Then
        optMove.Enabled = True
        optReceive.Enabled = False
        optClose.Enabled = True
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = True
        bolCanEdit = True
        StatusBar1.Panels.Item(1).Text = "Job packet OK to be Sent, Filed or Closed."
    End If
    If strTicketStatus = "CLOSED" Then
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optReceive.Enabled = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = True
        optFile.Enabled = False
        bolCanEdit = False
        StatusBar1.Panels.Item(1).Text = "This job packet is closed and cannot be changed until it has been re-opened."
    End If
    If strTicketAction = "CREATED" And strCurUser <> strLocalUser Then
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optReceive.Enabled = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = False
        bolCanEdit = False
        StatusBar1.Panels.Item(1).Text = "The job packet creator, " & strCurUser & ", has not Sent this job packet to anyone."
    ElseIf strTicketAction = "CREATED" And strCurUser = strLocalUser Then
        optMove.Enabled = True
        optReceive.Enabled = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = True
        bolCanEdit = True
        StatusBar1.Panels.Item(1).Text = "Job packet ready to be Sent."
    End If
    If EditMode = True Then
        optMove.Value = 0
        optReceive.Value = 0
        optClose.Value = 0
        optCreate.Value = 0
        optReOpen.Value = 0
        optFile.Value = 0
        optMove.Enabled = False
        cmbUsers.Visible = False
        lblUser.Visible = False
        optReceive.Enabled = False
        optClose.Enabled = False
        optCreate.Enabled = False
        optReOpen.Enabled = False
        optFile.Enabled = False
        bolCanEdit = False
        StatusBar1.Panels.Item(1).Text = "Enter the new data and then click the green checkmark to update."
    End If
    If Not bolIsAdmin And bolCanEdit Then
        bolCanEdit = True
    Else
        bolCanEdit = bolIsAdmin
    End If
End Sub
Public Sub LiveSearch(ByVal strSearchString As String) '
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    On Error GoTo LeaveSub
    List1.Clear
    ShowData
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT idTicketJobNum From ticketdatabase Where idTicketJobNum Like '" & strSearchString & "%' AND idTicketIsActive = '1' Order By ticketdatabase.idTicketJobNum"
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    Do Until rs.EOF
        With rs
            List1.AddItem !idTicketJobNum, .AbsolutePosition - 1
            rs.MoveNext
        End With
    Loop
    If rs.RecordCount >= 1 Then
        List1.Visible = True
    ElseIf rs.RecordCount <= 0 Then
        List1.Visible = False
    End If
    rs.Close
    cn.Close
LeaveSub:
    HideData
End Sub
Public Sub SubmitFile()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim intBlah As Integer
    If Trim$(strTicketComment) = "" Then
        MsgBox ("Please enter a comment describing the filing location.")
        optFile.Value = True
        frmComments.Show (vbModal)
        Exit Sub
    End If
    ShowData
    strSQL1 = "select * from TicketDatabase WHERE idTicketJobNum = '" & txtJobNo.Text & "' Order By idTicketDate desc"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        !idTicketIsActive = 0
        rs.Update
        rs.AddNew
        !idTicketAction = "FILED"
        !idTicketUser = strLocalUser
        !idTicketCreateDate = txtCreateDate.Text
        !idTicketStatus = "OPEN"
        !idTicketCreator = txtCreator.Text
        !idTicketUserFrom = "NULL"
        !idTicketUserTo = "NULL"
        !idTicketComment = strTicketComment
        !idTicketJobNum = txtJobNo.Text
        !idTicketPartNum = txtPartNoRev.Text
        !idTicketDrawingNum = txtDrawNoRev.Text
        !idTicketCustPoNum = txtCustPoNo.Text
        !idTicketSalesNum = txtSalesNo.Text
        !idTicketDescription = txtTicketDescription.Text
        !idTicketPlant = strPlant
        !idTicketIsActive = 1
        rs.Update
        rs.Close
        cn.Close
        HideData
    End With
    DisableBoxes
    cmdSubmit.Enabled = False
    optReceive.Value = False
    optMove.Value = False
    optClose.Value = False
    optCreate.Value = False
    optReOpen.Value = False
    optFile.Value = False
    bolOptionClicked = False
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    RefreshAfterEdit
    If Err.Number = 0 Then
        ShowBanner colFiled, "Job Packet Filed Successfully."
    Else
    End If
    Exit Sub
errs:
    Dim blah
    blah = MsgBox("An error was detected!" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Yikes!")
    ClearFields
    Err.Clear
End Sub
Public Sub SubmitReOpen()
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim intBlah As Integer
    On Error GoTo errs
    ShowData
    strSQL1 = "select * from TicketDatabase WHERE idTicketJobNum = '" & txtJobNo.Text & "' Order By idTicketDate desc"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        !idTicketIsActive = 0
        .Update
        .AddNew
        !idTicketAction = "REOPENED"
        !idTicketUser = strLocalUser
        !idTicketCreateDate = txtCreateDate.Text
        !idTicketStatus = "OPEN"
        !idTicketCreator = txtCreator.Text
        !idTicketUserFrom = "NULL"
        !idTicketUserTo = "NULL"
        !idTicketComment = strTicketComment
        !idTicketJobNum = txtJobNo.Text
        !idTicketPartNum = txtPartNoRev.Text
        !idTicketDrawingNum = txtDrawNoRev.Text
        !idTicketCustPoNum = txtCustPoNo.Text
        !idTicketSalesNum = txtSalesNo.Text
        !idTicketDescription = txtTicketDescription.Text
        !idTicketPlant = strPlant
        !idTicketIsActive = 1
        .Update
        rs.Close
        cn.Close
        HideData
    End With
    RefreshAfterEdit
    If Err.Number = 0 Then
        ShowBanner colReopened, "Job Packet Reopened Successfully."
    Else
    End If
    cmdSubmit.Enabled = False
    optReceive.Value = False
    optMove.Value = False
    optClose.Value = False
    optCreate.Value = False
    optReOpen.Value = False
    bolOptionClicked = False
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    Exit Sub
errs:
    Dim blah
    blah = MsgBox("An error was detected!" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Yikes!")
    ClearFields
    Err.Clear
End Sub
Public Sub SubmitClose()
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim intBlah As Integer
    On Error GoTo errs
    If Trim$(strTicketComment) = "" Then
        MsgBox ("Please enter a comment describing the closed file location.")
        optClose.Value = True
        frmComments.Show (vbModal)
        Exit Sub
    End If
    ShowData
    strSQL1 = "select * from TicketDatabase WHERE idTicketJobNum = '" & txtJobNo.Text & "' Order By idTicketDate desc"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        !idTicketIsActive = 0
        .Update
        .AddNew
        !idTicketAction = "NULL"
        !idTicketUser = strLocalUser
        !idTicketCreateDate = txtCreateDate.Text
        !idTicketStatus = "CLOSED"
        !idTicketCreator = txtCreator.Text
        !idTicketUserFrom = "NULL"
        !idTicketUserTo = "NULL"
        !idTicketComment = strTicketComment
        !idTicketJobNum = txtJobNo.Text
        !idTicketPartNum = txtPartNoRev.Text
        !idTicketDrawingNum = txtDrawNoRev.Text
        !idTicketCustPoNum = txtCustPoNo.Text
        !idTicketSalesNum = txtSalesNo.Text
        !idTicketDescription = txtTicketDescription.Text
        !idTicketPlant = strPlant
        !idTicketIsActive = 1
        rs.Update
        rs.Close
        cn.Close
        HideData
        RefreshAfterEdit
        If Err.Number = 0 Then
            ShowBanner colClosed, "Job Packet Closed Successfully."
        Else
        End If
    End With
    DisableBoxes
    cmdSubmit.Enabled = False
    optReceive.Value = False
    optMove.Value = False
    optClose.Value = False
    optCreate.Value = False
    optReOpen.Value = False
    bolOptionClicked = False
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    optFile.Value = False
    Exit Sub
errs:
    Dim blah
    blah = MsgBox("An error was detected!" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Yikes!")
    ClearFields
    Err.Clear
End Sub
Public Sub SubmitReceive()
    Dim rs          As New ADODB.Recordset
    Dim cn          As New ADODB.Connection
    Dim strSQL1     As String
    Dim ConfirmText As String
    On Error GoTo errs
    ShowData
    strSQL1 = "select * from TicketDatabase WHERE idTicketJobNum = '" & txtJobNo.Text & "' Order By idTicketDate desc"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        !idTicketIsActive = 0
        .Update
        .AddNew
        !idTicketAction = "RECEIVED"
        !idTicketUser = strLocalUser
        !idTicketCreateDate = txtCreateDate.Text
        !idTicketStatus = "OPEN"
        !idTicketCreator = txtCreator.Text
        !idTicketUserFrom = strUserFrom
        ConfirmText = "Job Packet Received From " & !idTicketUserFrom
        !idTicketUserTo = "NULL"
        !idTicketComment = strTicketComment
        !idTicketJobNum = txtJobNo.Text
        !idTicketPartNum = txtPartNoRev.Text
        !idTicketDrawingNum = txtDrawNoRev.Text
        !idTicketCustPoNum = txtCustPoNo.Text
        !idTicketSalesNum = txtSalesNo.Text
        !idTicketDescription = txtTicketDescription.Text
        !idTicketPlant = strPlant
        !idTicketIsActive = 1
        .Update
    End With
    rs.Close
    cn.Close
    HideData
    DisableBoxes
    RefreshAfterEdit
    cmdSubmit.Enabled = False
    optReceive.Value = False
    optMove.Value = False
    optClose.Value = False
    optCreate.Value = False
    optReOpen.Value = False
    optFile.Value = False
    bolOptionClicked = False
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    If Err.Number = 0 Then
        ShowBanner colReceived, ConfirmText
    Else
    End If
    Exit Sub
errs:
    Dim blah
    blah = MsgBox("An error was detected!" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Yikes!")
    ClearFields
    Err.Clear
End Sub
Public Sub SubmitMove()
    Dim rs          As New ADODB.Recordset
    Dim cn          As New ADODB.Connection
    Dim strSQL1     As String
    Dim ConfirmText As String
    Dim Hits        As Integer
    On Error GoTo errs
    Hits = GetINIValue(strSelectUserTo)
    If Hits = 0 Then
        Call SetINIValue(strSelectUserTo, 1)
    ElseIf Hits >= 1 Then
        Call SetINIValue(strSelectUserTo, (Hits + 1))
    End If
    ShowData
    strSQL1 = "select * from TicketDatabase WHERE idTicketJobNum = '" & txtJobNo.Text & "' Order By idTicketDate desc"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        !idTicketIsActive = 0
        .Update
        .AddNew
        !idTicketAction = "INTRANSIT"
        !idTicketUserFrom = strLocalUser
        !idTicketCreateDate = txtCreateDate.Text
        !idTicketStatus = "OPEN"
        !idTicketCreator = txtCreator.Text
        !idTicketUser = strLocalUser
        !idTicketUserTo = strSelectUserTo
        ConfirmText = "Job Packet Sent To " & !idTicketUserTo
        !idTicketUserFrom = strLocalUser
        !idTicketComment = strTicketComment
        cmbUsers.Visible = False
        lblUser.Visible = False
        cmbUsers.ComboItems.Item(1).Selected = True
        !idTicketJobNum = txtJobNo.Text
        !idTicketPartNum = txtPartNoRev.Text
        !idTicketDrawingNum = txtDrawNoRev.Text
        !idTicketCustPoNum = txtCustPoNo.Text
        !idTicketSalesNum = txtSalesNo.Text
        !idTicketDescription = txtTicketDescription.Text
        !idTicketPlant = strPlant
        !idTicketIsActive = 1
        rs.Update
    End With
    rs.Close
    cn.Close
    HideData
    RefreshAfterEdit
    cmdSubmit.Enabled = False
    optMove.Value = False
    optReceive.Value = False
    optMove.Value = False
    optClose.Value = False
    optCreate.Value = False
    optReOpen.Value = False
    optFile.Value = False
    bolOptionClicked = False
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    GetTopHits
    UpdateUserList
errs:
    If Err.Number = 0 Then
        ShowBanner colInTransit, ConfirmText
    ElseIf Err.Number <> 0 Then
        Dim blah
        blah = MsgBox("An error was detected!" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Yikes!")
        ClearFields
        Err.Clear
    End If
End Sub
Public Sub SubmitCreate()
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim strSQL1, strSQL2, strJobNum As String
    Dim FormatDate, FormatTime As String
    strJobNum = txtJobNo.Text
    On Error GoTo errs
    ShowData
    FormatDate = Format$(Date, strDBDateFormat)
    FormatTime = Format$(Time, "hh:mm:ss")
    strSQL2 = "SELECT idTicketJobNum From ticketdatabase Where idTicketJobNum = '" & strJobNum & "' Order By ticketdatabase.idTicketDate Desc"
    strSQL1 = "INSERT INTO TicketDatabase (idTicketCreateDate,idTicketCreator,idTicketUser,idTicketAction,idTicketStatus,idTicketuserFrom,idTicketUserTo,idTicketComment,idTicketJobNum,idTicketPartNum,idTicketDrawingNum,idTicketCustPoNum,idTicketSalesNum,idTicketDescription,idTicketPlant,idTicketIsActive) VALUES ('" & FormatDate & " " & FormatTime & "','" & strLocalUser & "','" & strLocalUser & "','CREATED','OPEN','NULL','NULL','" & Replace$(strTicketComment, "'", "''") & "','" & Replace$(txtJobNo.Text, "'", "''") & "','" & Replace$(txtPartNoRev.Text, "'", "''") & "','" & Replace$(txtDrawNoRev.Text, "'", "''") & "','" & Replace$(txtCustPoNo.Text, "'", "''") & "','" & Replace$(txtSalesNo.Text, "'", "''") & "','" & Replace$(txtTicketDescription.Text, "'", "''") & "','" & cmbPlant.Text & "','1')"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL2, cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        ShowBanner &HC0C0FF, "A Job Packet with that Job Number already exists!", 500
        optCreate.Value = 1
        cmdSubmit.Enabled = False
        txtJobNo.SetFocus
        rs.Close
        cn.Close
        HideData
        Exit Sub
    Else
        With rs
            .Close
            rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
        End With
    End If
    cn.Close
    HideData
    bolHasTicket = True
    DisableBoxes
    RefreshAfterEdit
    RefreshHistory
    cmdSubmit.Enabled = False
    optReceive.Value = False
    optMove.Value = False
    optClose.Value = False
    optCreate.Value = False
    optReOpen.Value = False
    optFile.Value = False
    bolOptionClicked = False
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    '
errs:
    If Err.Number = 0 Then
        cmbPlant.BackColor = vbWindowBackground
        txtCustPoNo.BackColor = vbWindowBackground
        txtDrawNoRev.BackColor = vbWindowBackground
        txtSalesNo.BackColor = vbWindowBackground
        txtPartNoRev.BackColor = vbWindowBackground
        txtTicketDescription.BackColor = vbWindowBackground
        txtJobNo.BackColor = vbWindowBackground
        ShowBanner colCreate, "Job Packet Created Successfully."
    Else
        Dim blah
        blah = MsgBox("An error was detected!" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Yikes!")
        Err.Clear
    End If
End Sub
Public Sub HideOpts()
    optMove.Enabled = False
    cmbUsers.Visible = False
    lblUser.Visible = False
    optReceive.Enabled = False
    optClose.Enabled = False
End Sub
Public Sub EnableBoxes()
    txtPartNoRev.Locked = False
    txtDrawNoRev.Locked = False
    txtSalesNo.Locked = False
    txtCustPoNo.Locked = False
    txtTicketDescription.Locked = False
    cmbPlant.Enabled = True
    frmComments.txtComment.Locked = False
End Sub
Public Sub DisableBoxes()
    txtPartNoRev.Locked = True
    txtDrawNoRev.Locked = True
    txtSalesNo.Locked = True
    txtCustPoNo.Locked = True
    txtTicketDescription.Locked = True
    cmbPlant.Enabled = False
    lblChars.Visible = False
End Sub
Public Sub RefreshAfterEdit() 'Fills fields, refreshes MyPackets, does not refresh History Grid.
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim strSQL1, strSQL2 As String
    On Error GoTo errs
    If txtJobNo.Text = "" Or optCreate.Value = True Or bolHasTicket = False Then Exit Sub
    SetBoxesForEdit "All"
    ShowData
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From ticketdatabase Where idTicketIsActive = '1' AND idTicketJobNum = '" & txtJobNo.Text & "' Order By ticketdatabase.idTicketDate Desc"
    strSQL2 = "SELECT * FROM ticketdb.ticketdatabase ticketdatabase_0" & " WHERE (ticketdatabase_0.idTicketAction='CREATED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='RECEIVED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='REOPENED') AND (ticketdatabase_0.idTicketUser='" & strLocalUser & "') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') OR (ticketdatabase_0.idTicketAction='INTRANSIT') AND (ticketdatabase_0.idTicketIsActive='1') AND (ticketdatabase_0.idTicketStatus='OPEN') AND (ticketdatabase_0.idTicketUserTo='" & strLocalUser & "')" & " ORDER BY ticketdatabase_0.idTicketDate"
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    With rs
        txtPartNoRev.Text = !idTicketPartNum
        txtDrawNoRev.Text = !idTicketDrawingNum
        txtCustPoNo.Text = !idTicketCustPoNum
        txtSalesNo.Text = !idTicketSalesNum
        txtCreator.Text = !idTicketCreator
        txtCreateDate.Text = !idTicketCreateDate
        txtActionDate.Text = !idTicketDate
        strTicketAction = !idTicketAction
        strUserFrom = !idTicketUserFrom
        strUserTo = !idTicketUserTo
        strCurUser = !idTicketUser
        strTicketStatus = !idTicketStatus
        txtTicketAction.Text = !idTicketAction
        txtTicketOwner.Text = !idTicketUser
        txtTicketDescription.Text = !idTicketDescription
        txtTicketStatus.Text = !idTicketStatus
        strPlant = !idTicketPlant
        cmbPlant.Text = strPlant
        If txtJobNo.Text = "" Then
            DisableBoxes
            tmrRefresher.Enabled = False
        Else
            bolHasTicket = True
            tmrRefresher.Enabled = True
        End If
        If !idTicketComment <> "" Then
            TheX = pbScrollBox.ScaleWidth
            strCommentText = !idTicketComment
            tmrScroll.Enabled = True
        Else
            pbScrollBox.Cls
            strCommentText = ""
            tmrScroll.Enabled = False
        End If
    End With
    rs.Close
    Dim LineIN, LineOUT, Row As Integer
    FlexGridOUT.Clear
    FlexGridOUT.Redraw = False
    FlexGridOUT.Rows = 2
    FlexGridOUT.FixedCols = 1
    FlexGridOUT.FixedRows = 1
    FlexGridIN.Clear
    FlexGridIN.Redraw = False
    FlexGridIN.Rows = 2
    FlexGridIN.FixedCols = 1
    FlexGridIN.FixedRows = 1
    rs.Open strSQL2, cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <= 0 Then
        SSTab1.TabCaption(3) = "On-hand Packets (0)"
        SSTab1.TabCaption(2) = "Incoming Packets (0)"
        FlexGridOUT.Visible = False
        FlexGridOUT.Redraw = True
        FlexGridIN.Visible = False
        FlexGridIN.Redraw = True
        rs.Close
        cn.Close
        HideData
        FlexGridOUT.Clear
        FlexGridIN.Clear
        Exit Sub
    End If
    LineIN = 1
    LineOUT = 1
    Row = 0
    FlexGridOUT.Rows = rs.RecordCount + 1
    FlexGridOUT.Cols = 10
    FlexGridIN.Rows = rs.RecordCount + 1
    FlexGridIN.Cols = 10
    ' Create header row
    FlexGridOUT.TextMatrix(0, 1) = "Job Number"
    FlexGridOUT.TextMatrix(0, 2) = "Part Number"
    FlexGridOUT.TextMatrix(0, 3) = "Description"
    FlexGridOUT.TextMatrix(0, 4) = "Sales Number"
    FlexGridOUT.TextMatrix(0, 5) = "Customer/PO Number"
    FlexGridOUT.TextMatrix(0, 6) = "Created By"
    FlexGridOUT.TextMatrix(0, 7) = "Create Date"
    FlexGridOUT.TextMatrix(0, 8) = "Last Activity Date"
    FlexGridOUT.TextMatrix(0, 9) = "Last Activity"
    FlexGridIN.TextMatrix(0, 1) = "Job Number"
    FlexGridIN.TextMatrix(0, 2) = "Part Number"
    FlexGridIN.TextMatrix(0, 3) = "Description"
    FlexGridIN.TextMatrix(0, 4) = "Sales Number"
    FlexGridIN.TextMatrix(0, 5) = "Customer/PO Number"
    FlexGridIN.TextMatrix(0, 6) = "Created By"
    FlexGridIN.TextMatrix(0, 7) = "Create Date"
    FlexGridIN.TextMatrix(0, 8) = "Last Activity Date"
    FlexGridIN.TextMatrix(0, 9) = "Last Activity"
    Do Until rs.EOF
        With rs
            If !idTicketAction = "CREATED" And !idTicketUser = strLocalUser Or !idTicketAction = "RECEIVED" And !idTicketUser = strLocalUser Or !idTicketAction = "REOPENED" And !idTicketUser = strLocalUser Then
                Row = Row + 1
                FlexGridOUT.TextMatrix(LineOUT, 0) = LineOUT
                FlexGridOUT.TextMatrix(LineOUT, 1) = !idTicketJobNum
                FlexGridOUT.TextMatrix(LineOUT, 2) = !idTicketPartNum
                FlexGridOUT.TextMatrix(LineOUT, 3) = !idTicketDescription
                FlexGridOUT.TextMatrix(LineOUT, 4) = !idTicketSalesNum
                FlexGridOUT.TextMatrix(LineOUT, 5) = !idTicketCustPoNum
                FlexGridOUT.TextMatrix(LineOUT, 6) = !idTicketCreator
                FlexGridOUT.TextMatrix(LineOUT, 7) = !idTicketCreateDate
                FlexGridOUT.TextMatrix(LineOUT, 8) = !idTicketDate
                If !idTicketAction = "CREATED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &H80C0FF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = "Job packet was CREATED by " & !idTicketCreator
                ElseIf !idTicketAction = "RECEIVED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &H80FFFF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = !idTicketUser & " RECEIVED the job packet from " & !idTicketUserFrom
                ElseIf !idTicketAction = "REOPENED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &HFF80FF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = !idTicketUser & " REOPENED the job packet."
                End If
                LineOUT = LineOUT + 1
            ElseIf !idTicketAction = "INTRANSIT" And !idTicketUserTo = strLocalUser Then '**************************************
                Row = Row + 1
                FlexGridIN.TextMatrix(LineIN, 0) = LineIN
                FlexGridIN.TextMatrix(LineIN, 1) = !idTicketJobNum
                FlexGridIN.TextMatrix(LineIN, 2) = !idTicketPartNum
                FlexGridIN.TextMatrix(LineIN, 3) = !idTicketDescription
                FlexGridIN.TextMatrix(LineIN, 4) = !idTicketSalesNum
                FlexGridIN.TextMatrix(LineIN, 5) = !idTicketCustPoNum
                FlexGridIN.TextMatrix(LineIN, 6) = !idTicketCreator
                FlexGridIN.TextMatrix(LineIN, 7) = !idTicketCreateDate
                FlexGridIN.TextMatrix(LineIN, 8) = !idTicketDate
                Call FlexGridRowColor(FlexGridIN, LineIN, &H80FF80)
                FlexGridIN.TextMatrix(LineIN, 9) = !idTicketUserFrom & " SENT the job packet to " & !idTicketUserTo
                LineIN = LineIN + 1
            ElseIf !idTicketStatus = "CLOSED" Then
NextLoop:
            End If
            Row = Row + 1
            rs.MoveNext
        End With
    Loop
    FlexGridOUT.Rows = LineOUT
    FlexGridIN.Rows = LineIN
    rs.Close
    SizeTheSheet FlexGridOUT
    SizeTheSheet FlexGridIN
    FlexGridOUT.Redraw = True
    FlexGridIN.Redraw = True
    FlexGridIN.Visible = True
    FlexGridOUT.Visible = True
    If LineIN <= 1 Then FlexGridIN.Visible = False
    If LineOUT <= 1 Then FlexGridOUT.Visible = False
    SSTab1.TabCaption(3) = "On-hand Packets (" & FlexGridOUT.Rows - 1 & ")"
    SSTab1.TabCaption(2) = "Incoming Packets (" & FlexGridIN.Rows - 1 & ")"
    If SSTab1.Tab = 2 And ProgHasFocus = True Then
        If Me.ActiveControl.Name <> "SSTab1" Then
            cn.Close
            HideData
            Exit Sub
        ElseIf Me.ActiveControl.Name <> "FlexGridIN" Then
            cn.Close
            HideData
            Exit Sub
        End If
        FlexGridIN.col = FlexINLastSel(1)
        FlexGridIN.Row = FlexINLastSel(0)
        FlexGridIN.ColSel = FlexINLastSel(1)
        FlexGridIN.RowSel = FlexINLastSel(0)
        FlexGridIN.SetFocus
    ElseIf SSTab1.Tab = 3 And ProgHasFocus = True And Me.ActiveControl.Name = "SSTab2" Or Me.ActiveControl.Name = "FlexGridOUT" Then
        If Me.ActiveControl.Name <> "SSTab2" Then
            cn.Close
            HideData
            Exit Sub
        ElseIf Me.ActiveControl.Name <> "FlexGridOUT" Then
            cn.Close
            HideData
            Exit Sub
        End If
        FlexGridOUT.col = FlexOUTLastSel(1)
        FlexGridOUT.Row = FlexOUTLastSel(0)
        FlexGridOUT.ColSel = FlexOUTLastSel(1)
        FlexGridOUT.RowSel = FlexOUTLastSel(0)
        FlexGridOUT.SetFocus
    End If
    cn.Close
    HideData
    Exit Sub
errs:
    If Err.Number = -2147467259 Then
        CommsDown
    Else
        CommsUp
    End If
End Sub
Public Sub ClearAllButJobN()
    ' CloseBanner
    strLastJobNum = ""
    cmbPlant.BackColor = vbWindowBackground
    txtCustPoNo.BackColor = vbWindowBackground
    txtDrawNoRev.BackColor = vbWindowBackground
    txtSalesNo.BackColor = vbWindowBackground
    txtPartNoRev.BackColor = vbWindowBackground
    txtTicketDescription.BackColor = vbWindowBackground
    txtJobNo.BackColor = vbWindowBackground
    cmdSubmit.BackColor = vbButtonFace
    txtTicketDescription.Text = ""
    txtPartNoRev.Text = ""
    txtDrawNoRev.Text = ""
    txtCustPoNo.Text = ""
    txtSalesNo.Text = ""
    txtCreator.Text = ""
    txtTicketAction.Text = ""
    txtTicketStatus.Text = ""
    txtCreateDate.Text = ""
    txtTicketOwner.Text = ""
    txtActionDate.Text = ""
    cmbUsers.Visible = False
    cmbUsers.ComboItems.Item(1).Selected = True
    lblUser.Visible = False
    frmComments.txtComment.Text = ""
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    strUserTo = ""
    strUserFrom = ""
    strTicketAction = ""
    strTicketStatus = ""
    optMove.Value = False
    optReceive.Value = False
    optClose.Value = False
    optFile.Value = False
    optReOpen.Value = False
    optMove.Enabled = False
    cmbUsers.Visible = False
    lblUser.Visible = False
    optReceive.Enabled = False
    optClose.Enabled = False
    optFile.Enabled = False
    optReOpen.Enabled = False
    cmbPlant.ListIndex = 0
    optCreate.Value = False
    bolHasTicket = False
    HideOpts
    tmrScroll.Enabled = False
    strCommentText = ""
    pbScrollBox.Cls
    SetBoxesForEdit "All"
    EditMode = False
    lblChars.Visible = False
    StatusBar1.Panels.Item(1).Text = ""
    FlexGridHist.Visible = False
    picOlder.Visible = False
    bolCanEdit = False
    FlexHistLastTopRow = 0
End Sub
Public Sub ClearFields()
    ClearBanners
    Screen.MousePointer = vbDefault
    HideData
    strLastJobNum = ""
    cmbPlant.BackColor = vbWindowBackground
    txtCustPoNo.BackColor = vbWindowBackground
    txtDrawNoRev.BackColor = vbWindowBackground
    txtSalesNo.BackColor = vbWindowBackground
    txtPartNoRev.BackColor = vbWindowBackground
    txtTicketDescription.BackColor = vbWindowBackground
    txtJobNo.BackColor = vbWindowBackground
    cmdSubmit.BackColor = vbButtonFace
    txtTicketDescription.Text = ""
    txtJobNo.Text = ""
    txtPartNoRev.Text = ""
    txtDrawNoRev.Text = ""
    txtCustPoNo.Text = ""
    txtSalesNo.Text = ""
    txtCreator.Text = ""
    txtTicketAction.Text = ""
    txtTicketStatus.Text = ""
    txtCreateDate.Text = ""
    txtTicketOwner.Text = ""
    txtActionDate.Text = ""
    bolOptionClicked = False
    cmbUsers.Visible = False
    cmbUsers.ComboItems.Item(1).Selected = True
    lblUser.Visible = False
    frmComments.txtComment.Text = ""
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    strUserTo = ""
    strUserFrom = ""
    strTicketAction = ""
    strTicketStatus = ""
    optMove.Value = False
    optReceive.Value = False
    optClose.Value = False
    optFile.Value = False
    optReOpen.Value = False
    optMove.Enabled = False
    cmbUsers.Visible = False
    lblUser.Visible = False
    optReceive.Enabled = False
    optClose.Enabled = False
    optFile.Enabled = False
    optReOpen.Enabled = False
    cmbPlant.ListIndex = 0
    optCreate.Value = False
    bolHasTicket = False
    HideOpts
    FlexGridHist.Visible = False
    FlexGridHist.Clear
    FlexGridHist.Cols = 0
    FlexGridHist.Rows = 0
    lblChars.Visible = False
    tmrScroll.Enabled = False
    strCommentText = ""
    pbScrollBox.Cls
    picOlder.Visible = False
    cmdEdit.Visible = False
    cmdEdit.Picture = ButtonPics(1)
    cmdEdit.ToolTipText = "Edit Field"
    bolCanEdit = False
    EditMode = False
    FlexHistLastTopRow = 0
End Sub
Public Sub GetTopHits()
    Dim sGet          As String
    Dim sSections()   As String
    Dim iSectionCount As Long
    Dim sKeys()       As String
    Dim iKeycount     As Long
    Dim iSection      As Long
    Dim iKey          As Long
    Dim lSect         As Long
    Dim MaxKeys       As Long
    Dim i             As Integer
    Dim SortHits()    As Variant
    ReDim SortHits(1, 0)
    bolNoHits = True
    With m_cIni
        .EnumerateAllSections sSections(), iSectionCount
        If iSectionCount > 1 Then
            bolNoHits = False
            cmbUsers.ComboItems.Clear
            cmbUsers.ComboItems.Add , , ""
            For iSection = 2 To iSectionCount
                .Section = sSections(iSection)
                .EnumerateCurrentSection sKeys(), iKeycount
                For iKey = 1 To iKeycount
                    ReDim Preserve SortHits(1, UBound(SortHits, 2) + 1)
                    .Key = sKeys(iKey)
                    SortHits(1, UBound(SortHits, 2)) = .Key
                    SortHits(0, UBound(SortHits, 2)) = .Value
                Next iKey
            Next iSection
            Call MySort(SortHits)
            If UBound(SortHits, 2) - 1 < 4 Then
                MaxKeys = UBound(SortHits, 2) - 1
            ElseIf UBound(SortHits, 2) - 1 >= 4 Then
                MaxKeys = 3
            End If
            For i = 0 To MaxKeys
                cmbUsers.ComboItems.Add , SortHits(1, i), ReturnEmpInfo(SortHits(1, i)).FullName, 1
            Next i
            cmbUsers.ComboItems.Add , , "____________________________"
        End If
    End With
End Sub
Public Sub UpdateUserList()
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim i       As Integer
    On Error GoTo errs
    strSQL1 = "select * from users"
    '
    On Error Resume Next
    '    ShowData
    '
    '    Set rs = New ADODB.Recordset
    '    Set cn = New ADODB.Connection
    '
    '    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    '    cn.CursorLocation = adUseClient
    'rs.Open strSQL1, cn, adOpenKeyset
    If bolNoHits Then
        cmbUsers.ComboItems.Clear
        cmbUsers.ComboItems.Add 1, , ""
    End If
    frmReportFilter.cmbUsers.Clear
    frmReportFilter.cmbUsers.AddItem "", 0
    frmRedirect.cmbOwner.Clear
    frmRedirect.cmbUserTo.Clear
    frmRedirect.cmbUserFrom.Clear
    frmRedirect.cmbOwner.AddItem "", 0
    frmRedirect.cmbUserTo.AddItem "", 0
    frmRedirect.cmbUserFrom.AddItem "", 0
    frmUserSelect.cmbUsers.Clear
    frmUserSelect.cmbUsers.AddItem "", 0
    'i = 1
    'ReDim strUserIndex(1, rs.RecordCount)
    ' Do Until rs.EOF
    ' With rs
    ' strUserIndex(0, i) = UCase$(!idUsers)
    ' strUserIndex(1, i) = !idFullname
    For i = 1 To UBound(strUserIndex, 2)
        cmbUsers.ComboItems.Add , strUserIndex(0, i), strUserIndex(1, i)
        frmReportFilter.cmbUsers.AddItem strUserIndex(1, i), i
        frmRedirect.cmbOwner.AddItem strUserIndex(1, i), i
        frmRedirect.cmbUserTo.AddItem strUserIndex(1, i), i
        frmRedirect.cmbUserFrom.AddItem strUserIndex(1, i), i
        frmUserSelect.cmbUsers.AddItem strUserIndex(1, i), i
        'i = i + 1
    Next i
    'DoEvents
    ' rs.MoveNext
    'End With
    '
    ' Loop
    ' rs.Close
    ' cn.Close
    ' HideData
    frmReportFilter.cmbUsers.ListIndex = 0
    Err.Clear
    Exit Sub
errs:
    If Err.Number = -2147467259 Then
        If bolInitialLoad = True Then
            Dim blah
            blah = MsgBox("Could not connect to the server!", vbCritical + vbOKOnly, "No Data")
            Unload Me
        Else
            CommsDown
        End If
    Else
        CommsUp
    End If
End Sub
Private Sub GetUserIndex()
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim i       As Integer
    On Error GoTo errs
    strSQL1 = "select * from users"
    '
    '        cmbUsers.ComboItems.Clear
    '    cmbUsers.ComboItems.Add 1, , ""
    '
    ShowData
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset
    i = 1
    ReDim strUserIndex(1, rs.RecordCount)
    Do Until rs.EOF
        With rs
            strUserIndex(0, i) = UCase$(!idUsers)
            strUserIndex(1, i) = !idFullname
            i = i + 1
            'DoEvents
            rs.MoveNext
        End With
    Loop
    rs.Close
    cn.Close
    HideData
    ' frmReportFilter.cmbUsers.ListIndex = 0
    Exit Sub
errs:
    If Err.Number = -2147467259 Then
        If bolInitialLoad = True Then
            Dim blah
            blah = MsgBox("Could not connect to the server!", vbCritical + vbOKOnly, "No Data")
            Unload Me
        Else
            CommsDown
        End If
    Else
        CommsUp
    End If
End Sub
Private Sub cmbPlant_Click()
    cmbPlant.BackColor = vbWindowBackground
    If cmbPlant.Text <> "" And txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtSalesNo.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Or bolHasTicket = True And bolOptionClicked = True Then
        cmdSubmit.Enabled = True
    Else
        cmdSubmit.Enabled = False
    End If
End Sub
Private Sub cmbUsers_Click()
    strSelectUserTo = UCase$(cmbUsers.SelectedItem.Key)
    If strSelectUserTo = "" Then
        bolOptionClicked = False
        cmdSubmit.Enabled = False
    Else
        cmdSubmit.Enabled = True
        bolOptionClicked = True
    End If
End Sub
Private Sub cmbUsers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If strSelectUserTo = "" Then
            bolOptionClicked = False
            cmdSubmit.Enabled = False
        Else
            cmdSubmit.Enabled = True
            bolOptionClicked = True
            Call cmdSubmit_Click
        End If
    End If
End Sub
Private Sub cmdAllClosedReport_Click()
    If bolRunning = True Then 'if already running the ary, dont try to start another one. (Prevents server flooding is return key is held down)
        Exit Sub
    Else
        ClearBanners
        ShowAllClosed
    End If
End Sub
Private Sub ShowAllClosed()
    bolRunning = True
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim Line    As Integer
    Dim TotT    As Single
    On Error GoTo errs
    Screen.MousePointer = vbHourglass
    Flexgrid1.Clear
    Flexgrid1.Redraw = False
    Flexgrid1.Rows = 2
    Flexgrid1.FixedCols = 1
    Flexgrid1.FixedRows = 1
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    strReportType = "All Closed Job Packets"
    sAddlMsg = ""
    ShowData
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From ticketdatabase Where idTicketIsActive = '1' and idTicketStatus = 'CLOSED' Order By ticketdatabase.idTicketDate Desc"
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    pBar.Value = 0
    frmpBar.Visible = True
    If rs.RecordCount <= 0 Then
        rs.Close
        cn.Close
        bolRunning = False
        HideData
        Flexgrid1.Clear
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Line = 1
    Flexgrid1.Rows = rs.RecordCount + 1
    Flexgrid1.Cols = 10
    ' Create header row
    Flexgrid1.TextMatrix(0, 1) = "Job Number"
    Flexgrid1.TextMatrix(0, 2) = "Part Number"
    Flexgrid1.TextMatrix(0, 3) = "Description"
    Flexgrid1.TextMatrix(0, 4) = "Sales Number"
    Flexgrid1.TextMatrix(0, 5) = "Customer/PO Number"
    Flexgrid1.TextMatrix(0, 6) = "Created By"
    Flexgrid1.TextMatrix(0, 7) = "Create Date"
    Flexgrid1.TextMatrix(0, 8) = "Last Activity Date"
    Flexgrid1.TextMatrix(0, 9) = "Last Activity"
    ReDim strUsedJobNums(rs.RecordCount + 1)
    pBar.Max = rs.RecordCount
    Do Until rs.EOF
        With rs
            pBar.Value = .AbsolutePosition
            DoEvents
            Flexgrid1.TextMatrix(Line, 0) = Line
            Flexgrid1.TextMatrix(Line, 1) = !idTicketJobNum
            Flexgrid1.TextMatrix(Line, 2) = !idTicketPartNum
            Flexgrid1.TextMatrix(Line, 3) = !idTicketDescription
            Flexgrid1.TextMatrix(Line, 4) = !idTicketSalesNum
            Flexgrid1.TextMatrix(Line, 5) = !idTicketCustPoNum
            Flexgrid1.TextMatrix(Line, 6) = !idTicketCreator
            Flexgrid1.TextMatrix(Line, 7) = !idTicketCreateDate
            Flexgrid1.TextMatrix(Line, 8) = !idTicketDate
            If !idTicketAction = "CREATED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80C0FF)
                Flexgrid1.TextMatrix(Line, 9) = "Job packet was CREATED by " & !idTicketCreator
            ElseIf !idTicketAction = "INTRANSIT" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80FF80)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUserFrom & " SENT the job packet to " & !idTicketUserTo
            ElseIf !idTicketAction = "RECEIVED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80FFFF)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " RECEIVED the job packet from " & !idTicketUserFrom
            ElseIf !idTicketStatus = "CLOSED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H8080FF)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " CLOSED the job packet."
            ElseIf !idTicketStatus = "OPEN" And !idTicketAction = "FILED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &HFF8080)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " FILED the job packet."
            ElseIf !idTicketAction = "REOPENED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &HFF80FF)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " REOPENED the job packet."
            End If
            Line = Line + 1
            rs.MoveNext
        End With
    Loop
    Flexgrid1.Rows = Line
    rs.Close
    cn.Close
    bolRunning = False
    HideData
    SizeTheSheet Flexgrid1
    Flexgrid1.Visible = True
    Flexgrid1.Redraw = True
    pBar.Value = 0
    frmpBar.Visible = False
    Screen.MousePointer = vbDefault
    TotT = lngQryTimes(intQryIndex) * 0.001
    StatusBar1.Panels.Item(1).Text = Line - 1 & " results returned in " & TotT & " seconds"
    Exit Sub
errs:
    If Err.Number = -2147467259 Then
        CommsDown
        Screen.MousePointer = vbDefault
    Else
        CommsUp
        Screen.MousePointer = vbDefault
    End If
    bolRunning = False
End Sub
Private Sub ShowAllOpenHeatMap()
    bolRunning = True
    Dim rs          As New ADODB.Recordset
    Dim cn          As New ADODB.Connection
    Dim strSQL1     As String
    Dim Line        As Integer
    Dim TotT        As Single
    Dim Entries     As Integer
    Dim OrderYesNo  As Integer
    Const ColorsRGB As Integer = 255
    Dim CalcColor   As Integer
    On Error GoTo errs
    OrderYesNo = MsgBox("Select 'Yes' to order the results by the number of entries." & vbCrLf & vbCrLf & "Select 'No' to order by most recent activity. (Default)", vbQuestion + vbYesNo, "Display results ordered by number of entries?")
    DoEvents
    Screen.MousePointer = vbHourglass
    Flexgrid1.Redraw = False
    Flexgrid1.Clear
    Flexgrid1.Rows = 2
    Flexgrid1.FixedCols = 1
    Flexgrid1.FixedRows = 1
    strReportType = "All Open Job Packets"
    sAddlMsg = ""
    ShowData
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From ticketdatabase Where idTicketIsActive = '1' Order By ticketdatabase.idTicketDate Desc"
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <= 0 Then
        rs.Close
        cn.Close
        bolRunning = False
        HideData
        Flexgrid1.Clear
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    QryEntryNumbers
    Line = 1
    Flexgrid1.Rows = rs.RecordCount + 1
    Flexgrid1.Cols = 11
    pBar.Value = 0
    frmpBar.Visible = True
    pBar.Max = rs.RecordCount
    ' Create header row
    Flexgrid1.TextMatrix(0, 1) = "Job Number"
    Flexgrid1.TextMatrix(0, 2) = "Part Number"
    Flexgrid1.TextMatrix(0, 3) = "Description"
    Flexgrid1.TextMatrix(0, 4) = "Sales Number"
    Flexgrid1.TextMatrix(0, 5) = "Customer/PO Number"
    Flexgrid1.TextMatrix(0, 6) = "Created By"
    Flexgrid1.TextMatrix(0, 7) = "Create Date"
    Flexgrid1.TextMatrix(0, 8) = "Last Activity Date"
    Flexgrid1.TextMatrix(0, 9) = "Last Activity"
    Flexgrid1.TextMatrix(0, 10) = "Entries"
    ReDim strUsedJobNums(rs.RecordCount + 1)
    Do Until rs.EOF
        With rs
            pBar.Value = .AbsolutePosition
            DoEvents
            Entries = GetNumberOfEntries(!idTicketJobNum)
            CalcColor = ColorsRGB - (Entries * RGBMulti)
            If CalcColor <= 0 Then CalcColor = 0
            Flexgrid1.TextMatrix(Line, 0) = Line
            Flexgrid1.TextMatrix(Line, 1) = !idTicketJobNum
            Flexgrid1.TextMatrix(Line, 2) = !idTicketPartNum
            Flexgrid1.TextMatrix(Line, 3) = !idTicketDescription
            Flexgrid1.TextMatrix(Line, 4) = !idTicketSalesNum
            Flexgrid1.TextMatrix(Line, 5) = !idTicketCustPoNum
            Flexgrid1.TextMatrix(Line, 6) = !idTicketCreator
            Flexgrid1.TextMatrix(Line, 7) = !idTicketCreateDate
            Flexgrid1.TextMatrix(Line, 8) = !idTicketDate
            Flexgrid1.TextMatrix(Line, 10) = Entries
            If !idTicketAction = "CREATED" Then
                Call FlexGridRowColor(Flexgrid1, Line, RGB(255, CalcColor, CalcColor))
                Flexgrid1.TextMatrix(Line, 9) = "Job packet was CREATED by " & !idTicketCreator
            ElseIf !idTicketAction = "INTRANSIT" Then
                Call FlexGridRowColor(Flexgrid1, Line, RGB(255, CalcColor, CalcColor))
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUserFrom & " SENT the job packet to " & !idTicketUserTo
            ElseIf !idTicketAction = "RECEIVED" Then
                Call FlexGridRowColor(Flexgrid1, Line, RGB(255, CalcColor, CalcColor))
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " RECEIVED the job packet from " & !idTicketUserFrom
            ElseIf !idTicketStatus = "CLOSED" Then
                Call FlexGridRowColor(Flexgrid1, Line, RGB(255, CalcColor, CalcColor))
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " CLOSED the job packet."
            ElseIf !idTicketStatus = "OPEN" And !idTicketAction = "FILED" Then
                Call FlexGridRowColor(Flexgrid1, Line, RGB(255, CalcColor, CalcColor))
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " FILED the job packet."
            ElseIf !idTicketAction = "REOPENED" Then
                Call FlexGridRowColor(Flexgrid1, Line, RGB(255, CalcColor, CalcColor))
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " REOPENED the job packet."
            End If
            Line = Line + 1
            rs.MoveNext
        End With
    Loop
    Flexgrid1.Rows = Line
    rs.Close
    cn.Close
    bolRunning = False
    HideData
    SizeTheSheet Flexgrid1
    If OrderYesNo = vbYes Then
        Flexgrid1.col = 10
        Flexgrid1.Sort = flexSortGenericDescending
    Else
        'dont order
    End If
    Flexgrid1.Redraw = True
    Flexgrid1.Visible = True
    pBar.Value = 0
    frmpBar.Visible = False
    Screen.MousePointer = vbDefault
    TotT = lngQryTimes(intQryIndex) * 0.001
    StatusBar1.Panels.Item(1).Text = Line - 1 & " results returned in " & TotT & " seconds"
    Exit Sub
errs:
    Screen.MousePointer = vbDefault
    If Err.Number = -2147467259 Then
        CommsDown
    Else
        CommsUp
    End If
    bolRunning = False
End Sub
Private Sub ShowAllOpen()
    bolRunning = True
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim Line    As Integer
    Dim TotT    As Single
    On Error GoTo errs
    Screen.MousePointer = vbHourglass
    Flexgrid1.Redraw = False
    Flexgrid1.Clear
    Flexgrid1.Rows = 2
    Flexgrid1.FixedCols = 1
    Flexgrid1.FixedRows = 1
    strReportType = "All Open Job Packets"
    sAddlMsg = ""
    ShowData
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From ticketdatabase Where idTicketIsActive = '1' and idTicketStatus = 'OPEN' Order By ticketdatabase.idTicketDate Desc"
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <= 0 Then
        rs.Close
        cn.Close
        bolRunning = False
        HideData
        Flexgrid1.Clear
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Line = 1
    Flexgrid1.Rows = rs.RecordCount + 1
    Flexgrid1.Cols = 10
    pBar.Value = 0
    frmpBar.Visible = True
    pBar.Max = rs.RecordCount
    ' Create header row
    Flexgrid1.TextMatrix(0, 1) = "Job Number"
    Flexgrid1.TextMatrix(0, 2) = "Part Number"
    Flexgrid1.TextMatrix(0, 3) = "Description"
    Flexgrid1.TextMatrix(0, 4) = "Sales Number"
    Flexgrid1.TextMatrix(0, 5) = "Customer/PO Number"
    Flexgrid1.TextMatrix(0, 6) = "Created By"
    Flexgrid1.TextMatrix(0, 7) = "Create Date"
    Flexgrid1.TextMatrix(0, 8) = "Last Activity Date"
    Flexgrid1.TextMatrix(0, 9) = "Last Activity"
    ReDim strUsedJobNums(rs.RecordCount + 1)
    Do Until rs.EOF
        With rs
            pBar.Value = .AbsolutePosition
            DoEvents
            Flexgrid1.TextMatrix(Line, 0) = Line
            Flexgrid1.TextMatrix(Line, 1) = !idTicketJobNum
            Flexgrid1.TextMatrix(Line, 2) = !idTicketPartNum
            Flexgrid1.TextMatrix(Line, 3) = !idTicketDescription
            Flexgrid1.TextMatrix(Line, 4) = !idTicketSalesNum
            Flexgrid1.TextMatrix(Line, 5) = !idTicketCustPoNum
            Flexgrid1.TextMatrix(Line, 6) = !idTicketCreator
            Flexgrid1.TextMatrix(Line, 7) = !idTicketCreateDate
            Flexgrid1.TextMatrix(Line, 8) = !idTicketDate
            If !idTicketAction = "CREATED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80C0FF)
                Flexgrid1.TextMatrix(Line, 9) = "Job packet was CREATED by " & !idTicketCreator
            ElseIf !idTicketAction = "INTRANSIT" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80FF80)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUserFrom & " SENT the job packet to " & !idTicketUserTo
            ElseIf !idTicketAction = "RECEIVED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80FFFF)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " RECEIVED the job packet from " & !idTicketUserFrom
            ElseIf !idTicketStatus = "CLOSED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H8080FF)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " CLOSED the job packet."
            ElseIf !idTicketStatus = "OPEN" And !idTicketAction = "FILED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &HFF8080)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " FILED the job packet."
            ElseIf !idTicketAction = "REOPENED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &HFF80FF)
                Flexgrid1.TextMatrix(Line, 9) = !idTicketUser & " REOPENED the job packet."
            End If
            Line = Line + 1
            rs.MoveNext
        End With
    Loop
    Flexgrid1.Rows = Line
    rs.Close
    cn.Close
    bolRunning = False
    HideData
    SizeTheSheet Flexgrid1
    Flexgrid1.Redraw = True
    Flexgrid1.Visible = True
    pBar.Value = 0
    frmpBar.Visible = False
    Screen.MousePointer = vbDefault
    TotT = lngQryTimes(intQryIndex) * 0.001
    StatusBar1.Panels.Item(1).Text = Line - 1 & " results returned in " & TotT & " seconds"
    Exit Sub
errs:
    Screen.MousePointer = vbDefault
    If Err.Number = -2147467259 Then
        CommsDown
    Else
        CommsUp
    End If
    bolRunning = False
End Sub
Private Sub cmdAllOpenReport_Click()
    If bolRunning = True Then 'if already running the ary, dont try to start another one. (Prevents server flooding is return key is held down)
        Exit Sub
    Else
        ClearBanners
        ShowAllOpen
    End If
End Sub
Private Sub cmdClear_Click()
    ClearFields
    DisableBoxes
    optCreate.Value = False
    optCreate.Enabled = True
    StatusBar1.Panels.Item(1).Text = ""
    txtJobNo.SetFocus
End Sub
Private Sub cmdComment_Click()
    frmComments.txtComment.Text = strTicketComment
    frmComments.Show (vbModal)
End Sub
Private Sub cmdFilterReport_Click()
    frmReportFilter.chkAllTickets.Value = 1
    frmReportFilter.cmbPacketType.ListIndex = 0
    frmReportFilter.Show (vbModal)
End Sub
Private Sub cmdGetInBox_Click()
    GetMyPackets
End Sub
Private Sub cmdGetOutBox_Click()
    GetMyPackets
End Sub
Private Sub cmdPrintHistory_Click()
    If FlexGridHist.Rows > 1 Then
        frmPrinters.Show 1
        If bolCancelPrint = True Then
            bolCancelPrint = False
            Exit Sub
        End If
        strReportType = "Packet History"
        sAddlMsg = "Job Number: " & txtJobNo.Text & "   Job Description: " & txtTicketDescription.Text
        'PrintFlexGrid FlexGridHist
        PrintFlexGridColor FlexGridHist
    Else
        MsgBox "Nothing to print!"
    End If
End Sub
Private Sub cmdPrintReport_Click()
    If Flexgrid1.Rows < 1 Then
        MsgBox ("Nothing to print!")
        Exit Sub
    End If
    frmPrinters.Show 1
    If bolCancelPrint = True Then
        bolCancelPrint = False
        Exit Sub
    End If
    Flexgrid1.ColWidth(1) = 1005
    Flexgrid1.ColWidth(2) = 1005
    Flexgrid1.ColWidth(3) = 2715
    Flexgrid1.ColWidth(4) = 930
    Flexgrid1.ColWidth(5) = 1170
    Flexgrid1.ColWidth(6) = 885
    Flexgrid1.ColWidth(7) = 1335
    Flexgrid1.ColWidth(8) = 1290
    Flexgrid1.ColWidth(9) = 3525
    PrintFlexGrid Flexgrid1
    SizeTheSheet Flexgrid1
End Sub
Private Sub cmdRefresh_Click()
    'If bolRefreshRunning = True Then Exit Sub
    tmrRefresher.Enabled = False
    tmrRefresher.Enabled = True    '  Reset timer
    Screen.MousePointer = vbHourglass
    DoEvents
    bolRefreshRunning = True
    RefreshAll
    RefreshHistory
    UpdateUserList
    bolRefreshRunning = False
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSearch_Click()
    OpenPacket txtJobNo.Text
End Sub
Private Sub cmdRefreshHist_Click()
    RefreshHistory
End Sub
Private Sub cmdSubmit_Click()
    cmdSubmit.BackColor = vbButtonFace
    lblChars.Visible = False
    If optFile.Value = True Then
        optCreate.Value = False
        optMove.Value = False
        optReceive.Value = False
        optClose.Value = False
        optReOpen.Value = False
        optFile.Value = False
        DisableBoxes
        SubmitFile
        SetControls
        Exit Sub
    End If
    If optCreate.Value = True Then
        optCreate.Value = False
        optMove.Value = False
        optReceive.Value = False
        optClose.Value = False
        optReOpen.Value = False
        optFile.Value = False
        DisableBoxes
        SubmitCreate
        SetControls
        Exit Sub
    End If
    If optMove.Value = True Then
        optMove.Value = False
        optCreate.Value = False
        optReceive.Value = False
        optClose.Value = False
        optReOpen.Value = False
        optFile.Value = False
        SubmitMove
        SetControls
        Exit Sub
        'cmbUsers.Refresh
    End If
    If optReceive.Value = True Then
        optReceive.Value = False
        optMove.Value = False
        optCreate.Value = False
        optClose.Value = False
        optReOpen.Value = False
        optFile.Value = False
        SubmitReceive
        SetControls
        Exit Sub
    End If
    If optClose.Value = True Then
        optMove.Value = False
        optCreate.Value = False
        optReceive.Value = False
        optClose.Value = False
        optReOpen.Value = False
        optFile.Value = False
        SubmitClose
        SetControls
        Exit Sub
    End If
    If optReOpen.Value = True Then
        optMove.Value = False
        optCreate.Value = False
        optReceive.Value = False
        optClose.Value = False
        optReOpen.Value = False
        optFile.Value = False
        SubmitReOpen
        SetControls
        Exit Sub
    End If
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
End Sub
Private Sub cmdShowMore_Click()
    'If bolOpenForm = False Then
    intMovement = 200
    intMovementAccel = 50
    'ElseIf bolOpenForm = True Then
    ' intMovement = 0
    'intMovementAccel = 0
    'End If
    tmrReSizer.Enabled = True
End Sub
Private Sub FlexGrid1_DblClick()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    DoEvents
    'ClearFields
    OpenPacket Flexgrid1.TextMatrix(Flexgrid1.RowSel, 1)
    tmrRefresher.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
Private Sub FlexGrid1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        txtJobNo.Text = Flexgrid1.TextMatrix(Flexgrid1.RowSel, 1)
        Call cmdSearch_Click
    End If
End Sub
Private Sub FlexGridHist_Click()
    Set WhichGrid = FlexGridHist
End Sub
Private Sub FlexGridHist_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    On Error Resume Next
    If Button = 1 Then
        intRowSel = FlexGridHist.RowSel
        If FlexGridHist.TextMatrix(FlexGridHist.RowSel, 4) = "com" Then
            FlexGridHist.col = 0
            FlexGridHist.Row = intRowSel - 1
            FlexGridHist.ColSel = FlexGridHist.Cols - 1
            FlexGridHist.RowSel = intRowSel
        ElseIf (FlexGridHist.RowSel + 1) < FlexGridHist.Rows And FlexGridHist.TextMatrix((FlexGridHist.RowSel + 1), 4) = "com" Then
            intRowSel = FlexGridHist.RowSel
            FlexGridHist.Row = 0
            FlexGridHist.col = 0
            FlexGridHist.ColSel = 0
            FlexGridHist.RowSel = 0
            FlexGridHist.col = 0
            FlexGridHist.Row = intRowSel
            FlexGridHist.ColSel = FlexGridHist.Cols - 1
            FlexGridHist.RowSel = intRowSel + 1
    
        End If
    End If
    If Button = 2 Then PopupMenu mnuPopup, vbPopupMenuRightButton, SSTab1.Left + Frame1.Left + FlexGridHist.Left + FlexGridHist.ColWidth(0), (SSTab1.Top + Frame1.Top + FlexGridHist.Top + FlexGridHist.CellTop + FlexGridHist.CellHeight)
End Sub
Private Sub FlexGridHist_Scroll()
    FlexHistLastTopRow = FlexGridHist.TopRow
    DisplayArrows
End Sub
Private Sub FlexGridIN_Click()
    Set WhichGrid = FlexGridIN
    Erase FlexINLastSel
    FlexINLastSel(0) = FlexGridIN.RowSel
    FlexINLastSel(1) = FlexGridIN.ColSel
    If FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1) = strLastJobNum Then
        Exit Sub
    Else
        strLastJobNum = FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1)
        OpenPacket FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1)
    End If
End Sub
Private Sub FlexGridIN_DblClick()
    ClearFields
    If FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1) = strLastJobNum Then
        Exit Sub
    Else
        strLastJobNum = FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1)
        OpenPacket FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1)
    End If
End Sub
Private Sub FlexGridIN_EnterCell()
    FlexINLastSel(0) = FlexGridIN.RowSel
    FlexINLastSel(1) = FlexGridIN.ColSel
End Sub
Private Sub FlexGridIN_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        FlexINLastSel(0) = FlexGridIN.RowSel
        FlexINLastSel(1) = FlexGridIN.ColSel
        If FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1) = strLastJobNum Then
            Exit Sub
        Else
            strLastJobNum = FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1)
            OpenPacket FlexGridIN.TextMatrix(FlexGridIN.RowSel, 1)
        End If
    End If
End Sub
Private Sub FlexGridIN_Scroll()
    intFlexGridInLastRow = FlexGridIN.TopRow
End Sub
Private Sub FlexGridOUT_Click()
    Set WhichGrid = FlexGridOUT
    Erase FlexOUTLastSel
    FlexOUTLastSel(0) = FlexGridOUT.RowSel
    FlexOUTLastSel(1) = FlexGridOUT.ColSel
    If FlexGridOUT.TextMatrix(FlexGridOUT.RowSel, 1) = strLastJobNum Then
        Exit Sub
    Else
        strLastJobNum = FlexGridOUT.TextMatrix(FlexGridOUT.RowSel, 1)
        OpenPacket FlexGridOUT.TextMatrix(FlexGridOUT.RowSel, 1)
    End If
End Sub
Private Sub FlexGridOUT_DblClick()
    ClearFields
    strLastJobNum = FlexGridOUT.TextMatrix(FlexGridOUT.RowSel, 1)
    OpenPacket FlexGridOUT.TextMatrix(FlexGridOUT.RowSel, 1)
End Sub
Private Sub FlexGridOUT_EnterCell()
    FlexOUTLastSel(0) = FlexGridOUT.RowSel
    FlexOUTLastSel(1) = FlexGridOUT.ColSel
End Sub
Private Sub FlexGridOUT_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        FlexOUTLastSel(0) = FlexGridOUT.RowSel
        FlexOUTLastSel(1) = FlexGridOUT.ColSel
        If FlexGridOUT.TextMatrix(FlexGridOUT.RowSel, 1) = strLastJobNum Then
            Exit Sub
        Else
            strLastJobNum = FlexGridOUT.TextMatrix(FlexGridOUT.RowSel, 1)
            OpenPacket FlexGridOUT.TextMatrix(FlexGridOUT.RowSel, 1)
        End If
    End If
End Sub
Private Sub FlexGridOUT_Scroll()
    intFlexGridOutLastRow = FlexGridOUT.TopRow
End Sub
Private Sub Form_Initialize()
    frmSplash.Show
    DoEvents
End Sub
Private Sub Form_Load()
    
    Dim i          As Integer
    Dim Commands() As String
    Dim ErrToss    As Boolean
    ErrToss = False
    On Error GoTo errs
    strINILoc = Environ$("APPDATA") & "\JPT\HITS.INI"
    Call CreateINI
    With m_cIni
        .Path = strINILoc
        .Section = "HITS"
    End With
    bolInitialLoad = True
    FindMySQLDriver
    Debug.Print strSQLDriver
    mnuAdmin.Visible = False
    mnuPopup.Visible = False
    picOlder.Top = FlexGridHist.Top + FlexGridHist.Height - picOlder.Height
    bolHook = True ' change to false to disable mouse hook (change to false when run in dev mode)
    intQryIndex = 0
    If bolHook Then
        Hook Me.hwnd, True
        Call WheelHook(Form1)
    End If
    lblAppVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
    intFlexGridInLastRow = 0
    intFlexGridOutLastRow = 0
    intPrevInPackets = 0
    intShpTimerStartWidth = shpTimer.Width
    intCachedBanners = 0
    intCurrentCache = -1
    frmConfirm.Top = -frmConfirm.Height
    bolMessageDelivered = False
    GetFadeColor
    HelpClosed = False
    DrawDayLines = True
    ReDim HistoryIcons(1 To 7)
    strLocalUser = UCase$(Environ$("USERNAME"))
    txtLocalUser.Text = strLocalUser
    cmdPrintReport.UseMaskColor = True
    cmdPrintHistory.UseMaskColor = True
    strSortMode = "A"
    frmSplash.lblStatus.Caption = "Connecting to server..."
    DoEvents
    strServerAddress = "10.35.1.40" '"10.0.1.232"
    strUserName = "TicketApp"
    strPassword = "yb4w4"
    intFormHMax = 10500 '10620 '10500
    intFormHMin = 5535 '5535 '5025
    If CheckForAdmin(strLocalUser) Then
        SetupAdmin
        'do stuff to enable admin things
    End If
    intSearchWait = 2
    Form1.Height = intFormHMin
    bolOpenForm = True
    bolOpenConfirm = True
    bolPrinting = False
    txtLocalUser.Text = strLocalUser
    txtDateTime.Text = Date & " " & Time
    cmbPlant.AddItem ""
    cmbPlant.AddItem "STEEL FAB"
    cmbPlant.AddItem "INDUSTRIAL MACH"
    cmbPlant.AddItem "NUCLEAR"
    cmbPlant.AddItem "CONTROLS"
    cmbPlant.AddItem "ROCKY MT"
    cmbPlant.AddItem "WOOSTER"
    SSTab1.Tab = 0
    SetComboBoxHeight cmbUsers, 260
    frmSplash.lblStatus.Caption = "Getting incoming and on-hand packets..."
    DoEvents
    GetMyPackets
    ' Cache icons into memory for quick access
    frmSplash.lblStatus.Caption = "Caching images..."
    DoEvents
    Call ImgList.ListImages.Add(, , LoadPicture(App.Path & "\Images\star2.gif"))
    cmbUsers.ImageList = ImgList
    Set HistoryIcons(1) = LoadPicture(App.Path & "\Images\Created-trans.gif")
    Set HistoryIcons(2) = LoadPicture(App.Path & "\Images\Sent-trans.gif")
    Set HistoryIcons(3) = LoadPicture(App.Path & "\Images\Received-trans.gif")
    Set HistoryIcons(4) = LoadPicture(App.Path & "\Images\Filed-trans.gif")
    Set HistoryIcons(5) = LoadPicture(App.Path & "\Images\Closed-trans.gif")
    Set HistoryIcons(6) = LoadPicture(App.Path & "\Images\Comment-trans.gif")
    Set HistoryIcons(7) = LoadPicture(App.Path & "\Images\Reopened-trans.gif")
    'Cache Timeline Help images
    ReDim HelpPics(1 To 5)
    For i = 1 To 5
        Set HelpPics(i) = LoadPicture(App.Path & "\Images\Timeline\" & i & ".gif")
    Next i
    ReDim ButtonPics(1 To 4)
    Set ButtonPics(1) = LoadPicture(App.Path & "\Images\Edit.bmp")
    Set ButtonPics(2) = LoadPicture(App.Path & "\Images\YesCheck.bmp")
    Set ButtonPics(3) = LoadPicture(App.Path & "\Images\Comment-En.gif")
    Set ButtonPics(4) = LoadPicture(App.Path & "\Images\Comment-Dis.gif")
    imgComment.Picture = ButtonPics(4)
    imgComment.Enabled = False
    Set TabPics(0) = LoadPicture(App.Path & "\Images\history.gif")
    Set TabPics(1) = LoadPicture(App.Path & "\Images\search.gif")
    Set TabPics(2) = LoadPicture(App.Path & "\Images\incoming.gif")
    Set TabPics(3) = LoadPicture(App.Path & "\Images\onhand.gif")
    SSTab1.TabPicture(0) = TabPics(0)
    SSTab1.TabPicture(1) = TabPics(1)
    SSTab1.TabPicture(2) = TabPics(2)
    SSTab1.TabPicture(3) = TabPics(3)
    
    Set picDataPics(0) = LoadPicture(App.Path & "\Images\DataOff2Light.gif")
    Set picDataPics(1) = LoadPicture(App.Path & "\Images\NoData2.gif")
    Set picDataPics(2) = LoadPicture(App.Path & "\Images\Data2.gif")
    
    frmSplash.lblStatus.Caption = "Loading user lists..."
    DoEvents
    GetUserIndex
    GetTopHits
    UpdateUserList
    If ErrToss = True Then
        frmSplash.lblStatus.Caption = "ERRORS WHILE LOADING!"
        Wait 5000
    End If
    frmSplash.Hide
    'Command line arguments
    Commands() = Split(Command$, " ")
    For i = 0 To UBound(Commands)
        If Commands(i) = "-m" Then ' Start expanded
            Form1.Height = intFormHMax
            bolOpenForm = False
            cmdShowMore.Caption = "Show Less"
            Label17.Caption = "á"
            SSTab1.ToolTipText = ""
        End If
        If Commands(i) = "-autorefreshoff" Then ' start with auto refresh off
            chkAutoRefresh.Value = 0
            tmrRefresher.Enabled = False
        End If
    Next
    TheX = pbScrollBox.ScaleWidth
    picOlder.Left = FlexGridHist.Left + FlexGridHist.Width / 2 - (picOlder.Width / 2) - 120
    bolInitialLoad = False
    lblQryTime.Caption = "0 ms"
    frmpBar.Top = SSTab1.Top + SSTab1.Height / 2 - frmpBar.Height / 2 + SSTab1.TabHeight - 250
    frmpBar.Left = 4440
    Exit Sub
errs:
    ErrToss = True
    frmSplash.Print Err.Description
    Err.Clear
    DoEvents
    Resume Next
End Sub
Private Sub SetupAdmin()
    bolIsAdmin = True
    FlexGridHist.HighLight = flexHighlightAlways
    mnuAdmin.Visible = True
    intFormHMin = intFormHMin + 300
    intFormHMax = intFormHMax + 300
    'FlexGridHist.FocusRect = flexFocusLight
    'FlexGridHist.SelectionMode = flexSelectionByRow
    'FlexGridHist.HighLight = flexHighlightWithFocus
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call WheelUnHook
    Unload Me
    End
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Hook Me.hwnd, False
End Sub
Private Sub Frame1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    Dim i As Integer
    For i = 0 To frmKey.UBound
        frmKey(i).Visible = False
    Next
End Sub
Private Sub Frame3_Click()
    List1.Visible = False
End Sub
Private Sub Frame4_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    Dim i As Integer
    For i = 0 To frmKey.UBound
        frmKey(i).Visible = False
    Next
End Sub
Private Sub Frame5_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    Dim i As Integer
    For i = 0 To frmKey.UBound
        frmKey(i).Visible = False
    Next
End Sub
Private Sub Frame6_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    Dim i As Integer
    For i = 0 To frmKey.UBound
        frmKey(i).Visible = False
    Next
End Sub
Private Sub frmConfirm_Click()
    BannerClick strConfirmClickCase
End Sub
Private Sub frmKey_MouseMove(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    Dim i As Integer
    For i = 0 To frmKey.UBound
        frmKey(i).Visible = False
    Next
End Sub
Private Sub imgComment_Click()
    frmComments.txtComment.Text = strTicketComment
    frmComments.Show (vbModal)
End Sub
Private Sub imgNewWindow_Click()
    If Flexgrid1.Visible = True Then
        Unload frmGrid
        CopyGrid Flexgrid1, frmGrid.FlexGrid
        Set WhichGrid = frmGrid.FlexGrid
        frmGrid.Caption = strReportType
        frmGrid.Show
    End If
End Sub
Private Sub imgNewWindowHist_Click()
    If FlexGridHist.Visible = True Then
        Unload frmGrid
        CopyGridHistory FlexGridHist, frmGrid.FlexGrid
        Set WhichGrid = frmGrid.FlexGrid
        frmGrid.Caption = "Packet History"
        frmGrid.Show
    End If
End Sub
Private Sub imgNewWindowIn_Click()
    If FlexGridIN.Visible = True Then
        Unload frmGrid
        CopyGrid FlexGridIN, frmGrid.FlexGrid
        Set WhichGrid = frmGrid.FlexGrid
        frmGrid.Caption = "Incoming Packets"
        frmGrid.Show
    End If
End Sub
Private Sub imgNewWindowOut_Click()
    If FlexGridOUT.Visible = True Then
        Unload frmGrid
        CopyGrid FlexGridOUT, frmGrid.FlexGrid
        Set WhichGrid = frmGrid.FlexGrid
        frmGrid.Caption = "On Hand Packets"
        frmGrid.Show
    End If
End Sub
Private Sub Label12_Click()
    Dim Huh
    Clicks = Clicks + 1
    If Clicks >= 3 Then
        Clicks = 0
        Randomize Timer
        Huh = Int(Rnd * 5) + 1
        If Huh = 1 Then
            MsgBox "Stop it! That tickles!"
        ElseIf Huh = 2 Then
            MsgBox "Knock it off!"
        ElseIf Huh = 3 Then
            MsgBox "That's my name, don't wear it out."
        ElseIf Huh = 4 Then
            MsgBox "Surprise!"
        ElseIf Huh = 5 Then
            MsgBox "No TV and No Beer Make Homer something something..."
            MsgBox "Go crazy?"
            MsgBox "Don't mind if I do!"
        End If
    End If
End Sub
Private Sub lblColorKey_MouseMove(Index As Integer, _
                                  Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    frmKey(Index).Top = lblColorKey(Index).Top - frmKey(Index).Height
    frmKey(Index).Left = lblColorKey(Index).Left + (lblColorKey(Index).Width / 2) - (frmKey(Index).Width / 2)
    frmKey(Index).Visible = True
End Sub
Private Sub lblConfirm_Click()
    BannerClick strConfirmClickCase
End Sub
Private Sub List1_GotFocus()
    tmrLiveSearch.Enabled = False
    intSearchWaitTicks = 0
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtJobNo.Text = List1.Text
        Call cmdSearch_Click
        List1.Visible = False
        List1.Clear
    End If
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtJobNo.Text = List1.Text
    Call cmdSearch_Click
    List1.Visible = False
    List1.Clear
End Sub
Private Sub mnuDelete_Click()
    Dim blah
    If bolHasTicket Then
        blah = MsgBox("Are you sure you want to delete this packet?" & vbCrLf & vbCrLf & "Job#: " & txtJobNo.Text & vbCrLf & "Desc: " & txtTicketDescription.Text & vbCrLf, vbYesNo + vbExclamation, "Delete Packet")
        If blah = vbYes Then
            DeletePacket txtJobNo.Text
        Else
        End If
    Else
        blah = MsgBox("Please open a packet first!", vbOKOnly + vbExclamation, "No packet open")
    End If
End Sub
Private Sub mnuDeleteEntry_Click()
    With FlexGridHist
        Call DeleteEntry(FlexGridHist.TextMatrix(FlexGridHist.RowSel, 5), FlexGridHist.TextMatrix(FlexGridHist.RowSel, 1))
    End With
End Sub
Private Sub mnuFauxUser_Click()
    If Not mnuFauxUser.Checked Then
        frmUserSelect.Show
    Else
        strLocalUser = UCase$(Environ$("USERNAME"))
        Form1.txtLocalUser.Enabled = True
        Form1.txtLocalUser.Locked = True
        frmUserSelect.cmbUsers.ListIndex = 0
        Form1.txtLocalUser.BackColor = vbWhite
        Form1.txtLocalUser.Text = strLocalUser
        'Form1.lblFauxUser.Visible = True
        Form1.GetMyPackets
        Form1.SetControls
        Form1.mnuFauxUser.Checked = False
        ShowBanner colInTransit, "Faux user disabled."
    End If
End Sub
Private Sub mnuRedirect_Click()
    If bolHasTicket Then
        frmRedirect.Show
        frmRedirect.GetPacket
    Else
        Dim blah
        blah = MsgBox("Please open a packet first!", vbOKOnly + vbExclamation, "No packet open")
    End If
End Sub
Private Sub mnuStats_Click()
    Dim blah
    Dim TotalPackets As Long
    Dim TotalEntries As Long
    Dim i            As Long, b As Long
    Dim strStats     As String
    Dim DBStats()    As Stats
    Dim rs           As New ADODB.Recordset
    Dim cn           As New ADODB.Connection
    Dim strSQL1      As String
    Dim strSQL2      As String
    Dim strSQL3      As String
    Dim strSQL4      As String
    strSQL1 = "SELECT Count(Distinct ticketdatabase_0.idTicketJobNum) FROM ticketdb.ticketdatabase ticketdatabase_0"
    strSQL2 = "SELECT COUNT(*) FROM ticketdb.ticketdatabase"
    strSQL3 = "SHOW TABLE STATUS FROM ticketdb LIKE 'ticketdatabase'"
    cn.Open "uid=" & strUserName & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    TotalPackets = rs.Fields(0)
    rs.Close
    rs.Open strSQL2, cn, adOpenForwardOnly, adLockReadOnly
    TotalEntries = rs.Fields(0)
    rs.Close
    rs.Open strSQL3, cn, adOpenForwardOnly, adLockReadOnly
    ReDim DBStats(rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        DBStats(i).Name = rs.Fields(i).Name
        DBStats(i).Value = IIf(IsNull(rs.Fields(i).Value), "NULL", rs.Fields(i).Value)
    Next i
    For i = 0 To UBound(DBStats)
        strStats = strStats + DBStats(i).Name & " - " & DBStats(i).Value & vbCrLf
    Next i
    rs.Close
    blah = MsgBox("DB Stats:" & vbCrLf & vbCrLf & "Total Packets: " & TotalPackets & vbCrLf & "Total Entries: " & TotalEntries & vbCrLf & vbCrLf & "Raw DB Info: " & vbCrLf & vbCrLf & strStats)
End Sub
Private Sub optClose_Click()
    cmdSubmit.Enabled = True
    SetBoxesForEdit "All"
    bolOptionClicked = True
    cmbUsers.Visible = False
    lblUser.Visible = False
    imgComment.Picture = ButtonPics(3)
    imgComment.Enabled = True
    frmComments.txtComment.Text = ""
    frmComments.txtComment.Locked = False
End Sub
Private Sub optCreate_Click()
    SetBoxesForEdit "All"
    EnableBoxes
    'ClearFields
    bolOptionClicked = True
    cmbUsers.Visible = False
    lblUser.Visible = False
    imgComment.Picture = ButtonPics(3)
    imgComment.Enabled = True
    frmComments.txtComment.Text = ""
    frmComments.txtComment.Locked = False
    txtTicketDescription.BackColor = &HC0FFC0
    cmbPlant.BackColor = &HC0FFC0
    txtCustPoNo.BackColor = &HC0FFC0
    txtDrawNoRev.BackColor = &HC0FFC0
    txtSalesNo.BackColor = &HC0FFC0
    txtPartNoRev.BackColor = &HC0FFC0
    txtJobNo.BackColor = &HC0FFC0
    txtJobNo.SetFocus
    If txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtSalesNo.Text <> "" And txtTicketDescription.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Then
        cmdSubmit.Enabled = True
    End If
End Sub
Private Sub optFile_Click()
    SetBoxesForEdit "All"
    bolOptionClicked = True
    cmbUsers.Visible = False
    lblUser.Visible = False
    imgComment.Picture = ButtonPics(3)
    imgComment.Enabled = True
    frmComments.txtComment.Text = ""
    frmComments.txtComment.Locked = False
    cmdSubmit.Enabled = True
    SetBoxesForEdit "All"
End Sub
Private Sub optMove_Click()
    SetBoxesForEdit "All"
    imgComment.Picture = ButtonPics(3)
    imgComment.Enabled = True
    frmComments.txtComment.Text = ""
    frmComments.txtComment.Locked = False
    cmdSubmit.Enabled = False
    cmbUsers.Visible = True
    cmbUsers.SetFocus
    lblUser.Visible = True
    SendMessage cmbUsers.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End Sub
Private Sub optReceive_Click()
    SetBoxesForEdit "All"
    imgComment.Picture = ButtonPics(3)
    imgComment.Enabled = True
    bolOptionClicked = True
    cmdSubmit.Enabled = True
    cmbUsers.Visible = False
    lblUser.Visible = False
    frmComments.txtComment.Text = ""
    frmComments.txtComment.Locked = False
End Sub
Private Sub optReOpen_Click()
    SetBoxesForEdit "All"
    cmdSubmit.Enabled = True
    bolOptionClicked = True
    cmbUsers.Visible = False
    lblUser.Visible = False
    frmComments.txtComment.Text = ""
    imgComment.Picture = ButtonPics(3)
    imgComment.Enabled = True
    frmComments.txtComment.Locked = False
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then Set WhichGrid = FlexGridHist
    If SSTab1.Tab = 1 Then Set WhichGrid = Flexgrid1
    If SSTab1.Tab = 2 Then Set WhichGrid = FlexGridIN
    If SSTab1.Tab = 3 Then Set WhichGrid = FlexGridOUT
End Sub
Private Sub SSTab1_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    If bolOpenForm = True Then
        Call cmdShowMore_Click
    End If
End Sub
Private Sub tmrBannerWait_Timer()
    On Error Resume Next
WaitforBannerClose:
    If bolBannerOpen = False Then
        If bolBannerCleared = True Then Exit Sub
        intCurrentCache = intCurrentCache + 1
        OpenCloseBanner BannerColor(intCurrentCache), BannerText(intCurrentCache), BannerTicks(intCurrentCache), BannerCase(intCurrentCache), BannerFontColor(intCurrentCache)
        If intCurrentCache + 1 >= intCachedBanners Then
            intCachedBanners = 0
            intCurrentCache = -1
            tmrBannerWait.Enabled = False
        End If
    Else
        Wait 20
        DoEvents
        If bolBannerCleared = True Then Exit Sub
        GoTo WaitforBannerClose
    End If
End Sub
Function GetFade(steps, step)
    GetFade = RGB(r1 + (r2 - r1) / steps * step, g1 + (g2 - g1) / steps * step, b1 + (b2 - b1) / steps * step)
End Function
Private Sub tmrButtonFlasher_Timer()
    Dim iSteps    As Integer
    Dim FadeColor As Long
    iSteps = 255
    If cmdSubmit.Enabled = True Then
        If shpButtonFlash.Visible = False Then shpButtonFlash.Visible = True
        If iStep <= 0 Then iStep = iSteps
        FadeColor = RGB(r1 + (r2 - r1) / iSteps * iStep, g1 + (g2 - g1) / iSteps * iStep, b1 + (b2 - b1) / iSteps * iStep)
        shpButtonFlash.BackColor = FadeColor
        shpButtonFlash.BorderColor = FadeColor
        iStep = iStep - 8
    Else
        iStep = 0
        shpButtonFlash.Visible = False
    End If
End Sub
Private Sub tmrConfirmSlider_Timer()
    On Error Resume Next
    Dim intSliderMax, intSliderMin As Integer
    intSliderMax = 0
    intSliderMin = -frmConfirm.Height
    If bolWaitToClose = False Then
        If bolOpenConfirm = True Then   ' Open
            bolBannerOpen = True
            frmConfirm.Top = frmConfirm.Top + intConfirmMovement
            ' DoEvents
            If frmConfirm.Top >= intSliderMax Then
                'tmrConfirmSlider.Enabled = False
                frmConfirm.Top = intSliderMax
                bolWaitToClose = True
                bolOpenConfirm = False
                intConfirmMovement = 0
                'Exit Sub
            End If
        ElseIf bolOpenConfirm = False Then 'Close
            frmConfirm.Top = frmConfirm.Top - intConfirmMovement
            'DoEvents
            If frmConfirm.Top <= intSliderMin Then
                tmrConfirmSlider.Enabled = False
                bolWaitToClose = False
                frmConfirm.Visible = False
                frmConfirm.Top = intSliderMin
                bolOpenConfirm = True
                bolBannerOpen = False
                Exit Sub
            End If
        End If
        intConfirmMovement = intConfirmMovement + 5
    Else
        If intTicksWaited >= intTicksToWait Then
            bolWaitToClose = False
            intTicksWaited = 0
        Else
            bolWaitToClose = True
            intTicksWaited = intTicksWaited + 1
            If sngShapeResize > shpTimer.Width Then
                shpTimer.Width = 0
            Else
                shpTimer.Width = shpTimer.Width - sngShapeResize
            End If
            shpTimer.Left = frmConfirm.Width / 2 - shpTimer.Width / 2
        End If
    End If
End Sub
Private Sub tmrDateTime_Timer()
    txtDateTime.Text = Date & " " & Time
    'Me.Refresh
End Sub
Private Sub tmrLiveSearch_Timer()
    On Error Resume Next
    intSearchWaitTicks = intSearchWaitTicks + 1
    If bolHasTicket = True Then
        tmrLiveSearch.Enabled = False
        intSearchWaitTicks = 0
    End If
    If intSearchWaitTicks >= intSearchWait Then
        LiveSearch (txtJobNo.Text)
        intSearchWaitTicks = 0
        tmrLiveSearch.Enabled = False
    End If
End Sub
Private Sub tmrRefresher_Timer()
    On Error Resume Next
    If GetActiveWindow() <> Form1.hwnd Then
        'Do form's lost-focus routines here.
        ProgHasFocus = False
    Else
        ProgHasFocus = True
    End If
    If chkAutoRefresh.Value = 0 Then Exit Sub
    If EditMode = True Then Exit Sub
    RefreshAll
    txtDateTime.Text = Date & " " & Time
End Sub
Private Sub tmrReSizer_Timer()
    On Error Resume Next
    cmdShowMore.Enabled = False
    If bolOpenForm = True Then   ' Open
        Form1.Height = Form1.Height + intMovement
        ' DoEvents
        If Form1.Height + intMovement >= intFormHMax Then
            tmrReSizer.Enabled = False
            Form1.Height = intFormHMax
            bolOpenForm = False
            cmdShowMore.Caption = "Hide Tabs"
            Label17.Caption = "á"
            SSTab1.ToolTipText = ""
            If Me.Top + Me.Height > Screen.Height - 200 Then Me.Top = Screen.Height - Me.Height - 600
            cmdShowMore.Enabled = True
            Form1.Refresh
            Exit Sub
        End If
    End If
    If bolOpenForm = False Then  'Close
        Form1.Height = Form1.Height - intMovement
        'DoEvents
        If Form1.Height - intMovement <= intFormHMin Then
            tmrReSizer.Enabled = False
            Form1.Height = intFormHMin
            bolOpenForm = True
            cmdShowMore.Caption = "Show Tabs"
            Label17.Caption = "â"
            SSTab1.ToolTipText = "Click to expand"
            If Me.Top + Me.Height > Screen.Height - 200 Then Me.Top = Screen.Height - Me.Height - 600
            cmdShowMore.Enabled = True
            Form1.Refresh
            Exit Sub
        End If
    End If
    intMovement = intMovement + intMovementAccel
    intMovementAccel = intMovementAccel + 2
End Sub
Private Sub tmrScroll_Timer()
    On Error Resume Next
    pbScrollBox.Cls ' so we don't get text trails
    ' Scroll from right to left
    If TheX <= 0 - pbScrollBox.TextWidth(strCommentText) Then
        TheX = pbScrollBox.ScaleWidth
    Else
        TheX = TheX - 15 ' larger number means faster scrolling
    End If
    pbScrollBox.CurrentX = TheX
    pbScrollBox.CurrentY = 22 'TheY
    pbScrollBox.Print strCommentText
End Sub
Private Sub tmrWindowFlasher_Timer()
    FlashWindow Me.hwnd, Invert
End Sub
Private Sub txtCustPoNo_Change()
    PositionMaxChar txtCustPoNo
    txtCustPoNo.BackColor = vbWindowBackground
    If cmbPlant.Text <> "" And txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtSalesNo.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Then
        cmdSubmit.Enabled = True
    Else
        cmdSubmit.Enabled = False
    End If
End Sub
Private Sub txtCustPoNo_Click()
    PositionEdit Me.ActiveControl
End Sub
Private Sub txtCustPoNo_GotFocus()
    With txtCustPoNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtCustPoNo_LostFocus()
    txtCustPoNo.Text = Trim$(UCase$(txtCustPoNo.Text))
End Sub
Private Sub txtDrawNoRev_Change()
    PositionMaxChar txtDrawNoRev
    txtDrawNoRev.BackColor = vbWindowBackground
    If cmbPlant.Text <> "" And txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtTicketDescription.Text <> "" And txtSalesNo.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Then
        cmdSubmit.Enabled = True
    Else
        cmdSubmit.Enabled = False
    End If
End Sub
Private Sub txtDrawNoRev_Click()
    PositionEdit Me.ActiveControl
End Sub
Private Sub txtDrawNoRev_GotFocus()
    With txtDrawNoRev
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtDrawNoRev_LostFocus()
    txtDrawNoRev.Text = Trim$(UCase$(txtDrawNoRev.Text))
End Sub
Private Sub txtJobNo_Change()
    CloseBanner
    txtJobNo.BackColor = vbWindowBackground
    If cmbPlant.Text <> "" And txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtTicketDescription.Text <> "" And txtSalesNo.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Then
        cmdSubmit.Enabled = True
    Else
        cmdSubmit.Enabled = False
    End If
    If bolHasTicket = False Then
    Else
        ClearAllButJobN
        cmbUsers.Visible = False
        lblUser.Visible = False
        cmbUsers.ComboItems.Item(1).Selected = True
        bolOptionClicked = False
        cmdSubmit.Enabled = False
        imgComment.Picture = ButtonPics(4)
        imgComment.Enabled = False
        bolHasTicket = False
    End If
End Sub
Private Sub txtJobNo_GotFocus()
    With txtJobNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtJobNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
        List1.Visible = False
        tmrLiveSearch.Enabled = False
        intSearchWaitTicks = 0
        txtTicketDescription.SetFocus
    End If
End Sub
Private Sub txtJobNo_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If List1.ListCount <= 0 Then List1.Visible = False
    If Len(txtJobNo.Text) >= 3 Then
        tmrLiveSearch.Enabled = True
        intSearchWaitTicks = 0
        'LiveSearch (txtJobNo.Text)
    Else
        List1.Visible = False
        tmrLiveSearch.Enabled = False
        intSearchWaitTicks = 0
    End If
    If KeyCode = vbKeyDown Then
        List1.SetFocus
        List1.Selected(0) = True
    End If
End Sub
Private Sub txtJobNo_LostFocus()
    txtJobNo.Text = Trim$(UCase$(txtJobNo.Text))
    'List1.Visible = False
    If GetTabState And optCreate.Value = 0 Then Call cmdSearch_Click
End Sub
Private Sub txtPartNoRev_Change()
    PositionMaxChar txtPartNoRev
    txtPartNoRev.BackColor = vbWindowBackground
    If cmbPlant.Text <> "" And txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtTicketDescription.Text <> "" And txtSalesNo.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Then
        cmdSubmit.Enabled = True
    Else
        cmdSubmit.Enabled = False
    End If
End Sub
Private Sub txtPartNoRev_Click()
    PositionEdit Me.ActiveControl
End Sub
Private Sub txtPartNoRev_GotFocus()
    List1.Visible = False
    With txtPartNoRev
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtPartNoRev_LostFocus()
    txtPartNoRev.Text = Trim$(UCase$(txtPartNoRev.Text))
End Sub
Private Sub txtSalesNo_Change()
    PositionMaxChar txtSalesNo
    txtSalesNo.BackColor = vbWindowBackground
    If cmbPlant.Text <> "" And txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtTicketDescription.Text <> "" And txtSalesNo.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Then
        cmdSubmit.Enabled = True
    Else
        cmdSubmit.Enabled = False
    End If
End Sub
Private Sub txtSalesNo_Click()
    PositionEdit Me.ActiveControl
End Sub
Private Sub txtSalesNo_GotFocus()
    List1.Visible = False
    With txtSalesNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtSalesNo_LostFocus()
    txtSalesNo.Text = Trim$(UCase$(txtSalesNo.Text))
End Sub
Private Sub txtTicketAction_Change()
    If txtTicketAction.Text = "INTRANSIT" Then
        txtTicketAction.ForeColor = &H8000&
        txtTicketAction.Text = "IN-TRANSIT to " & strUserTo
    End If
    If txtTicketAction.Text = "RECEIVED" Then
        txtTicketAction.Text = "RECEIVED by " & strCurUser
        txtTicketAction.ForeColor = &HC0C0&
    End If
    If txtTicketAction.Text = "FILED" Then
        txtTicketAction.Text = "FILED by " & strCurUser
        txtTicketAction.ForeColor = vbBlue
    End If
    If txtTicketAction.Text = "REOPENED" Then
        txtTicketAction.Text = "REOPENED by " & strCurUser
        txtTicketAction.ForeColor = &HFF00FF
    End If
    If txtTicketAction.Text = "NULL" Then
        txtTicketAction.Text = "CLOSED by " & strCurUser
        txtTicketAction.ForeColor = vbRed
    End If
    If txtTicketAction.Text = "CREATED" Then
        txtTicketAction.Text = "CREATED by " & strCurUser
        txtTicketAction.ForeColor = &H80FF&
    End If
End Sub
Private Sub txtTicketAction_GotFocus()
    List1.Visible = False
End Sub
Private Sub PositionEdit(WhatControl As TextBox)
    If EditMode = True Then Exit Sub
    If bolCanEdit = True Then
        cmdEdit.Visible = False
        cmdEdit.Left = WhatControl.Left + WhatControl.Width + 105
        cmdEdit.Top = WhatControl.Top + 120
        cmdEdit.Visible = True
        ActiveText = WhatControl.Name
    Else
        cmdEdit.Visible = False
    End If
End Sub
Private Sub txtTicketDescription_Change()
    PositionMaxChar txtTicketDescription
    txtTicketDescription.BackColor = vbWindowBackground
    If cmbPlant.Text <> "" And txtJobNo.Text <> "" And txtPartNoRev.Text <> "" And txtSalesNo.Text <> "" And txtDrawNoRev.Text <> "" And txtCustPoNo.Text <> "" And optCreate.Value = True Then
        cmdSubmit.Enabled = True
    Else
        cmdSubmit.Enabled = False
    End If
End Sub
Private Sub txtTicketDescription_Click()
    PositionEdit Me.ActiveControl
End Sub
Private Sub txtTicketDescription_GotFocus()
    List1.Visible = False
    With txtTicketDescription
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtTicketStatus_Change()
    On Error Resume Next
    If txtTicketStatus.Text = "CLOSED" Then txtTicketStatus.ForeColor = &HFF&
    If txtTicketStatus.Text = "OPEN" Then txtTicketStatus.ForeColor = &H8000&
End Sub
