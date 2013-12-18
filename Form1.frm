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
   Begin VB.PictureBox frmConfirm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   3300
      ScaleHeight     =   1155
      ScaleWidth      =   5595
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   0
      Width           =   5595
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4800
         MouseIcon       =   "Form1.frx":08CA
         MousePointer    =   2  'Cross
         TabIndex        =   126
         Top             =   120
         Width           =   255
      End
      Begin VB.Shape shpTimer 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   1320
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Shape Border 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         Height          =   855
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Visible         =   0   'False
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
         TabIndex        =   125
         Top             =   420
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.Frame frmTimers 
      Caption         =   "Timers"
      Height          =   5055
      Left            =   10755
      TabIndex        =   91
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
      TabIndex        =   84
      Top             =   6000
      Visible         =   0   'False
      Width           =   5355
      Begin ComctlLib.ProgressBar pBar 
         Height          =   405
         Left            =   120
         TabIndex        =   85
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
         TabIndex        =   86
         Top             =   360
         Width           =   5190
      End
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   370
      Left            =   7080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":0D5C
      Style           =   1  'Graphical
      TabIndex        =   43
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
      TabIndex        =   53
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
            TextSave        =   ""
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
      TabIndex        =   23
      ToolTipText     =   "Click to expand"
      Top             =   4320
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
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
      TabPicture(0)   =   "Form1.frx":12E0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "Form1.frx":1810
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Incoming Packets"
      TabPicture(2)   =   "Form1.frx":1C82
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "On Hand Packets"
      TabPicture(3)   =   "Form1.frx":1E1C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   74
         Top             =   480
         Width           =   11775
         Begin VB.Frame frmKey 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   3
            Left            =   10920
            TabIndex        =   117
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
               TabIndex        =   123
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
               TabIndex        =   122
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
               TabIndex        =   121
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
               TabIndex        =   120
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
               TabIndex        =   119
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
               TabIndex        =   118
               Top             =   1200
               Width           =   765
            End
         End
         Begin VB.CommandButton cmdPrintOnPack 
            Caption         =   "&Print"
            Height          =   840
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Form1.frx":250E
            Style           =   1  'Graphical
            TabIndex        =   37
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
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Maually refresh my packets"
            Top             =   480
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGridOUT 
            Height          =   4215
            Left            =   1560
            TabIndex        =   34
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
            Picture         =   "Form1.frx":40A2
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
            TabIndex        =   124
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "On-hand Packets"
            Height          =   195
            Left            =   6120
            TabIndex        =   82
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
         TabIndex        =   73
         Top             =   480
         Width           =   11775
         Begin VB.Frame frmKey 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   2
            Left            =   10920
            TabIndex        =   109
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
               TabIndex        =   115
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
               TabIndex        =   114
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
               TabIndex        =   113
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
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   110
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.CommandButton cmdPrintInPack 
            Caption         =   "&Print"
            Height          =   840
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "Form1.frx":4197
            Style           =   1  'Graphical
            TabIndex        =   36
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
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Maually refresh my packets"
            Top             =   480
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGridIN 
            Height          =   4215
            Left            =   1560
            TabIndex        =   33
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
            Picture         =   "Form1.frx":5D2B
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
            TabIndex        =   116
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
            TabIndex        =   81
            Top             =   2160
            Width           =   1245
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   70
         Top             =   480
         Width           =   11775
         Begin VB.Frame frmKey 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   1
            Left            =   10920
            TabIndex        =   101
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
               TabIndex        =   107
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
               TabIndex        =   106
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
               TabIndex        =   105
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
               TabIndex        =   104
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
               TabIndex        =   103
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
               TabIndex        =   102
               Top             =   1200
               Width           =   765
            End
         End
         Begin VB.CommandButton cmdAllOpenReport 
            Caption         =   "All Opened"
            Height          =   360
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Display all currently open packets"
            Top             =   795
            Width           =   1335
         End
         Begin VB.CommandButton cmdAllClosedReport 
            Caption         =   "All Closed"
            Height          =   360
            Left            =   120
            TabIndex        =   29
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
            TabIndex        =   28
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
            Picture         =   "Form1.frx":5E20
            Style           =   1  'Graphical
            TabIndex        =   27
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
            TabIndex        =   47
            ToolTipText     =   "Shows heat map of packet entries. (More entries = hotter)"
            Top             =   4200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flexgrid1 
            Height          =   4215
            Left            =   1560
            TabIndex        =   26
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
            TabIndex        =   108
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Image imgNewWindow 
            Appearance      =   0  'Flat
            Height          =   450
            Left            =   600
            Picture         =   "Form1.frx":79B4
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
            TabIndex        =   89
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
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   11775
         Begin VB.Frame frmKey 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   0
            Left            =   10920
            TabIndex        =   93
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
               TabIndex        =   99
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
               TabIndex        =   98
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
               TabIndex        =   97
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
               TabIndex        =   96
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
               TabIndex        =   95
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
               TabIndex        =   94
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
            Picture         =   "Form1.frx":7AA9
            ScaleHeight     =   300
            ScaleWidth      =   9810
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   4140
            Visible         =   0   'False
            Width           =   9810
         End
         Begin VB.CommandButton cmdTimeLine 
            Caption         =   "View Timeline"
            Height          =   480
            Left            =   120
            TabIndex        =   40
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
            Picture         =   "Form1.frx":945C
            Style           =   1  'Graphical
            TabIndex        =   25
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
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Manually refresh history data"
            Top             =   360
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGridHist 
            Height          =   4215
            Left            =   1560
            TabIndex        =   38
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
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "Form1.frx":AFF0
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
            TabIndex        =   100
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Image imgNewWindowHist 
            Appearance      =   0  'Flat
            Height          =   450
            Left            =   600
            Picture         =   "Form1.frx":AFF8
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
            TabIndex        =   72
            Top             =   2160
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tracking Info."
      Height          =   3975
      Left            =   7380
      TabIndex        =   60
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   2490
         TabIndex        =   83
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
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   360
            Width           =   765
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   360
            Left            =   120
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   92
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
         TabIndex        =   42
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
         TabIndex        =   22
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
         TabIndex        =   35
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
         TabIndex        =   17
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
         TabIndex        =   21
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
         TabIndex        =   7
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   76
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
         TabIndex        =   66
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action Date"
         Height          =   195
         Left            =   2520
         TabIndex        =   75
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Status"
         Height          =   195
         Left            =   240
         TabIndex        =   65
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latest Action"
         Height          =   195
         Left            =   240
         TabIndex        =   64
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create Date"
         Height          =   195
         Left            =   2520
         TabIndex        =   63
         Top             =   2160
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packet Creator"
         Height          =   195
         Left            =   240
         TabIndex        =   62
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Packet Owner"
         Height          =   195
         Left            =   2520
         TabIndex        =   61
         Top             =   1560
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Packet Info."
      Height          =   3975
      Left            =   120
      TabIndex        =   54
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox pbSubmitBox 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   2040
         ScaleHeight     =   855
         ScaleWidth      =   2295
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3060
         Width           =   2295
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
            Left            =   240
            MaskColor       =   &H000000FF&
            TabIndex        =   52
            ToolTipText     =   "Submit update"
            Top             =   180
            Width           =   1815
         End
      End
      Begin MSComctlLib.ImageCombo cmbUsers 
         Height          =   330
         Left            =   1680
         TabIndex        =   48
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
         TabIndex        =   8
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
         TabIndex        =   14
         Top             =   3510
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
         TabIndex        =   10
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   11
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
         TabIndex        =   9
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   87
         Top             =   3205
         Width           =   2175
         Begin VB.CommandButton cmdShowMore 
            Caption         =   "Show Tabs"
            Height          =   360
            Left            =   480
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Show additional features"
            Top             =   240
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
            TabIndex        =   88
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Image imgComment 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   4080
         Picture         =   "Form1.frx":B0ED
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
         TabIndex        =   78
         ToolTipText     =   "Current Chars / Max Chars"
         Top             =   1800
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2160
         TabIndex        =   77
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
         TabIndex        =   71
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
         TabIndex        =   68
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
      TabIndex        =   90
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
      TabIndex        =   67
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
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_TAB = &H9
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
Private bolNoHits      As Boolean
Private intRowSel      As Integer
Private strCommentText As String
Private intMovement As Integer, intMovementAccel As Integer, intMovementAccelRate As Integer
Public Sub GetUserIndex()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim i       As Integer
    On Error GoTo errs
    strSQL1 = "select * from users"
 ShowData
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset
    i = 1
    ReDim strUserIndex(2, rs.RecordCount)
    Do Until rs.EOF
        With rs
            strUserIndex(0, i) = UCase$(!idUsers)
            strUserIndex(1, i) = !idFullname
            strUserIndex(2, i) = !idEmail
            i = i + 1
            rs.MoveNext
        End With
    Loop
   HideData
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
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colCreate)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | Job packet was created by " & strCreator
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(1)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colCreate)
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
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colInTransit)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUserFrom & " sent the job packet to " & strUserTo
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(2)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colInTransit)
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
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colReceived)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUser & " received the job packet from " & strUserFrom
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(3)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colReceived)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    ElseIf strAction = "NULL" Then
        If strComment <> "" And bolPrinting = False Then
            FlexGridHist.Rows = FlexGridHist.Rows + 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 4) = "com"
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = "    " & Chr$(34) & strComment & Chr$(34)
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 1
            FlexGridHist.CellFontSize = lngFontSize
            FlexGridHist.CellFontItalic = True
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colClosed)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUser & " closed the job packet."
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(5)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colClosed)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    ElseIf strAction = "FILED" Then
        If strComment <> "" And bolPrinting = False Then
            FlexGridHist.Rows = FlexGridHist.Rows + 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 4) = "com"
            FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = "    " & Chr$(34) & strComment & Chr$(34)
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 1
            FlexGridHist.CellFontSize = lngFontSize
            FlexGridHist.CellFontItalic = True
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colFiled)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUser & " filed the job packet."
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(4)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colFiled)
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
            FlexGridHist.Row = FlexGridHist.Rows - 1
            FlexGridHist.col = 0
            Set FlexGridHist.CellPicture = HistoryIcons(6)
            FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
            Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colReopened)
            FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
        End If
        FlexGridHist.Rows = FlexGridHist.Rows + 1 'Add new row per entry
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 5) = strGUID
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 3) = FlexGridHist.Rows - 1
        FlexGridHist.TextMatrix(FlexGridHist.Rows - 1, 1) = strDate & " | " & strUser & " reopened the job packet."
        FlexGridHist.Row = FlexGridHist.Rows - 1
        FlexGridHist.col = 0
        Set FlexGridHist.CellPicture = HistoryIcons(7)
        FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
        Call FlexGridRowColor(FlexGridHist, FlexGridHist.Rows - 1, colReopened)
        FlexGridHist.RowHeight(FlexGridHist.Rows - 1) = intRowH
    End If
End Sub
Public Sub OpenPacket(JobNum As String) 'Opens Packet - Fills HistoryGrid, Fills Fields, Does not refresh MyPackets
    Dim rs      As New ADODB.Recordset
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
    ShowData
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.packetentrydb LEFT JOIN (ticketdb.packetlist) ON (packetlist.idJobNum=packetentrydb.idJobNum) WHERE packetlist.idJobNum = '" & JobNum & "' ORDER BY packetentrydb.idDate DESC"
    Set rs = cn_global.Execute(strSQL1)
    List1.Clear
    If rs.RecordCount <= 0 Then Err.Raise vbObjectError + 513, "ADO Open", "Zero Records Returned For Query"
    With rs
        dtLatestHistDate = Format$(!idDate, strDBDateTimeFormat)
        txtPartNoRev.Text = !idPartNum
        txtDrawNoRev.Text = !idDrawingNum
        txtCustPoNo.Text = !idCustPoNum
        txtSalesNo.Text = !idSalesNum
        txtCreator.Text = GetFullName(!idCreator)
        txtCreateDate.Text = !idCreateDate
        txtTicketOwner.Text = GetFullName(!idUser)
        txtActionDate.Text = !idDate
        strTicketAction = !idAction
        strUserFrom = !idUserFrom
        strUserTo = !idUserTo
        strTicketStatus = !idStatus
        txtTicketStatus.Text = UCase$(!idStatus)
        strCurUser = !idUser
        txtTicketAction.Text = !idAction
        txtTicketDescription.Text = !idDescription
        strLatestComment = !idComment
        frmComments.txtComment.Text = strLatestComment
        frmComments.txtComment.Locked = True
        strPlant = !idPlant
        cmbPlant.Text = strPlant
        If !idComment <> "" Then
            TheX = pbScrollBox.ScaleWidth
            strCommentText = !idComment
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
            Call FillFlexHist(!idAction, !idStatus, !idComment, !idDate, GetFullName(!idCreator), GetFullName(!idUserFrom), GetFullName(!idUserTo), GetFullName(!idUser), !idGUIDEntry)
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
    FlexBoldFirst FlexGridHist
    FlexGridRedrawHeight
    FlexGridHist.Redraw = True
    FlexGridHist.Visible = True
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
    If Hex$(Err.Number) = 80040201 Then
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
    Dim strSQL1 As String
    Dim dtTicketDate1, dtTicketDate2 As Date
    On Error Resume Next
    ShowData
    strSQL1 = "SELECT * FROM ticketdb.packetentrydb LEFT JOIN (ticketdb.packetlist) ON (packetlist.idJobNum=packetentrydb.idJobNum) WHERE packetlist.idJobNum = '" & txtJobNo.Text & "' ORDER BY packetentrydb.idDate"
    cn_global.CursorLocation = adUseClient
    Set rs = cn_global.Execute(strSQL1)
    Entry = 0
    With rs
        ReDim strTimelineComments(.RecordCount)
        dtTicketDate1 = !idDate
        .MoveLast
        dtTicketDate2 = !idDate
        TotalTime = DateDiff("n", dtTicketDate1, dtTicketDate2)
        .MoveFirst
    End With
    Do Until rs.EOF
        With rs
            If !idComment <> "" Then strTimelineComments(Entry) = Chr$(34) & !idComment & Chr$(34)
            dtTicketDate1 = !idDate
            .MoveNext
            If .EOF Then
                .MovePrevious
                dtTicketDate1 = !idDate
                dtTicketDate2 = Date & " " & Time
                TicketHours(Entry) = DateDiff("n", dtTicketDate1, dtTicketDate2)
                .MoveNext
            Else
                dtTicketDate2 = !idDate
                TicketHours(Entry) = DateDiff("n", dtTicketDate1, dtTicketDate2)
            End If
            .MovePrevious
            If !idAction = "CREATED" Then
                TicketActionText(Entry) = " Job packet was CREATED by " & !idCreator & " | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days ", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idAction = "INTRANSIT" Then
                TicketActionText(Entry) = " " & !idUserFrom & " SENT the job packet to " & !idUserTo & " | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days ", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idAction = "RECEIVED" Then
                TicketActionText(Entry) = " " & !idUser & " RECEIVED the job packet from " & !idUserFrom & " | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days ", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idAction = "NULL" Then
                TicketActionText(Entry) = " " & !idUser & " CLOSED the job packet. | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idAction = "FILED" Then
                TicketActionText(Entry) = " " & !idUser & " FILED the job packet. | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            ElseIf !idAction = "REOPENED" Then
                TicketActionText(Entry) = " " & !idUser & " REOPENED the job packet. | " & (IIf(TicketHours(Entry) > 1440, Round(TicketHours(Entry) / 1440, 1) & " days", Round(TicketHours(Entry) / 60, 1) & " hrs "))
            End If
            TicketDate(Entry) = !idDate
            TicketAction(Entry) = !idAction
            .MoveNext
            Entry = Entry + 1
        End With
    Loop
    If TotalTime / 1440 > 60 Then
        DrawDayLines = False
        frmTimeLine.chkDayLines.Value = 0
    End If
    TicketActionText(Entry - 1) = TicketActionText(Entry - 1) + " (Ongoing)"
    HideData
End Sub

Public Sub GetMyPackets(Optional Verbose As Boolean = True)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim LineIN, LineOUT, Row As Integer
    On Error GoTo errs
    strSQL1 = "SELECT * FROM ticketdb.packetentrydb LEFT JOIN (ticketdb.packetlist) ON (packetlist.idJobNum=packetentrydb.idJobNum) WHERE" & " ticketdb.packetentrydb.idDate=(SELECT MAX(s2.idDate) FROM ticketdb.packetentrydb s2 WHERE ticketdb.packetentrydb.idJobNum = s2.idJobNum" & " AND packetlist.idMailbox='" & strLocalUser & "') ORDER BY idDate DESC"
    cn_global.CursorLocation = adUseClient
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
    Set rs = cn_global.Execute(strSQL1)
    If rs.RecordCount <= 0 Then
        intPrevInPackets = 0
        SSTab1.TabCaption(3) = "On-hand Packets (0)"
        SSTab1.TabCaption(2) = "Incoming Packets (0)"
        FlexGridOUT.Visible = False
        FlexGridOUT.Redraw = True
        FlexGridIN.Visible = False
        FlexGridIN.Redraw = True
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
            If !idAction = "CREATED" And !idUser = strLocalUser Or !idAction = "RECEIVED" And !idUser = strLocalUser Or !idAction = "REOPENED" And !idUser = strLocalUser Then
                Row = Row + 1
                FlexGridOUT.TextMatrix(LineOUT, 0) = LineOUT
                FlexGridOUT.TextMatrix(LineOUT, 1) = !idJobNum
                FlexGridOUT.TextMatrix(LineOUT, 2) = !idPartNum
                FlexGridOUT.TextMatrix(LineOUT, 3) = !idDescription
                FlexGridOUT.TextMatrix(LineOUT, 4) = !idSalesNum
                FlexGridOUT.TextMatrix(LineOUT, 5) = !idCustPoNum
                FlexGridOUT.TextMatrix(LineOUT, 6) = !idCreator
                FlexGridOUT.TextMatrix(LineOUT, 7) = !idCreateDate
                FlexGridOUT.TextMatrix(LineOUT, 8) = !idDate
                If !idAction = "CREATED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &H80C0FF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = "Job packet was CREATED by " & !idCreator
                ElseIf !idAction = "RECEIVED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &H80FFFF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = !idUser & " RECEIVED the job packet from " & !idUserFrom
                ElseIf !idAction = "REOPENED" Then
                    Call FlexGridRowColor(FlexGridOUT, LineOUT, &HFF80FF)
                    FlexGridOUT.TextMatrix(LineOUT, 9) = !idUser & " REOPENED the job packet."
                End If
                LineOUT = LineOUT + 1
            ElseIf !idAction = "INTRANSIT" And !idUserTo = strLocalUser Then '**************************************
                Row = Row + 1
                FlexGridIN.TextMatrix(LineIN, 0) = LineIN
                FlexGridIN.TextMatrix(LineIN, 1) = !idJobNum
                FlexGridIN.TextMatrix(LineIN, 2) = !idPartNum
                FlexGridIN.TextMatrix(LineIN, 3) = !idDescription
                FlexGridIN.TextMatrix(LineIN, 4) = !idSalesNum
                FlexGridIN.TextMatrix(LineIN, 5) = !idCustPoNum
                FlexGridIN.TextMatrix(LineIN, 6) = !idCreator
                FlexGridIN.TextMatrix(LineIN, 7) = !idCreateDate
                FlexGridIN.TextMatrix(LineIN, 8) = !idDate
                Call FlexGridRowColor(FlexGridIN, LineIN, &H80FF80)
                FlexGridIN.TextMatrix(LineIN, 9) = !idUserFrom & " SENT the job packet to " & !idUserTo
                LineIN = LineIN + 1
            ElseIf !idStatus = "CLOSED" Then
NextLoop:
            End If
            Row = Row + 1
            rs.MoveNext
        End With
    Loop
    FlexGridOUT.Rows = LineOUT
    FlexGridIN.Rows = LineIN
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
    If FlexGridIN.Rows - 1 > intPrevInPackets And Verbose Then
        ShowBanner vbCyan, "You have incoming Job Packets. Click to view.", 500, "VIEWPACK" '&HC0C0C0
        intPrevInPackets = FlexGridIN.Rows - 1
    Else
        intPrevInPackets = FlexGridIN.Rows - 1
    End If
    ' intPrevInPackets = FlexGridIN.Rows - 1
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
Public Sub ShowData()
    Set pbData.Picture = picDataPics(2)
    DoEvents
    StartTimer
End Sub
Public Sub HideData()
    Dim lngCurQry As Double, lngAddQry As Double, lngAvgQry As Double
    Dim i         As Integer
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
Public Sub RefreshAll()
    Dim ConcurrentStatus As Integer
    ConcurrentStatus = DBConcurrent
    If bolHasTicket Then
        If ConcurrentStatus = 2 Then
            ClearFields
            GetMyPackets
            SetControls
            optCreate.Value = False
            optCreate.Enabled = True
            StatusBar1.Panels.Item(1).Text = ""
            txtJobNo.SetFocus
            ShowBanner vbYellow, "The packet has been deleted.  Current status updated.", 350
            Exit Sub
        End If
        If ConcurrentStatus = 0 Then
            ClearOptBoxes
            RefreshFields
            SetControls
        End If
        RefreshHistory
    End If
    GetMyPackets
End Sub
Public Sub CommsDown()
    On Error Resume Next
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
        cn_global.Close
        cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
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
        Dim strSQL1 As String
        strSQL1 = "SELECT * From packetlist Where idJobNum = '" & txtJobNo.Text & "'"
        cn_global.CursorLocation = adUseClient
        rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
        Do Until rs.EOF
            With rs
                If ActiveText = "txtPartNoRev" And txtPartNoRev <> PrevPartNum Then !idPartNum = UCase$(txtPartNoRev.Text)
                If ActiveText = "txtDrawNoRev" And txtDrawNoRev <> PrevDrawNoRev Then !idDrawingNum = UCase$(txtDrawNoRev.Text)
                If ActiveText = "txtCustPoNo" And txtCustPoNo <> PrevCustPoNo Then !idCustPoNum = UCase$(txtCustPoNo.Text)
                If ActiveText = "txtSalesNo" And txtSalesNo <> PrevSalesNo Then !idSalesNum = UCase$(txtSalesNo.Text)
                If ActiveText = "txtTicketDescription" And txtTicketDescription <> PrevDescription Then !idDescription = txtTicketDescription.Text
                rs.Update
                rs.MoveNext
            End With
        Loop
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
    frmTimeLine.DrawTimeLine
    frmTimeLine.Show
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
    FadeColor = GetRealColor(&HFF00&)
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
    On Error Resume Next
    intCellHeight = 600
    FlexGrid.Row = 0
    FlexGrid.col = 1
    FlexGrid.CellFontSize = 10
    FlexGrid.CellFontBold = True
    If FlexGrid.TextMatrix(1, 4) = "com" Then
        FlexGrid.Row = 1
        FlexGrid.col = 1
        FlexGrid.CellFontBold = True
        FlexGrid.CellFontSize = 10.75
        FlexGrid.RowHeight(1) = intCellHeight - 200
    End If
    Exit Sub
errs:
    If Err.Number = 381 Then FlexGrid.RowHeight(0) = intCellHeight 'if Subscript out of range, it most likely means the grid only has one row. Therefor, no comment, it should fail and finish setting grid height
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
        TwipPix = .ColWidth(c) ' * Screen.TwipsPerPixelX
        XFirstColumn = xmin ' + TwipPix + 2 ' * GAP
        lngYTopOfGrid = Printer.CurrentY
        Printer.CurrentY = Printer.CurrentY + GAP
        'If FlexGrid.Header = True Then
        X = xmin + GAP
        '        For c = 1 To .Cols
        '            Printer.CurrentX = x
        '            TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX
        '            PrevY = Printer.CurrentY
        '            If c = .Cols Then
        '                lngStartY = Printer.CurrentY - GAP + 5
        '                lngStartX = Printer.CurrentX - GAP + 5
        '                lngEndX = xmax
        '
        '                lngEndY = Printer.CurrentY + Printer.TextHeight(.ColHeader(c)) + GAP
        '                Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), &H80000016, BF
        '            Else
        '                lngStartY = Printer.CurrentY - GAP + 5
        '                lngStartX = Printer.CurrentX - GAP + 5
        '                lngEndX = Printer.CurrentX + TwipPix + GAP
        '                lngEndY = Printer.CurrentY + Printer.TextHeight(.ColHeader(c)) + GAP
        '                Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), &H80000016, BF
        '            End If
        '            Printer.CurrentX = lngStartX + GAP
        '            Printer.CurrentY = PrevY
        '
        '            Printer.Print BoundedText(Printer, .ColHeader(c), TwipPix);
        '            TwipPix = .ColWidth(c) * Screen.TwipsPerPixelX
        '            x = x + TwipPix + 2 * GAP
        '        Next c
        'Printer.CurrentY = Printer.CurrentY + GAP
        'Printer.Print
        ' End If
        For R = 0 To .Rows - 1
            'If bolStop = True Then
            ' Printer.EndDoc
            ' bolStop = False
            ' frmpBar.Visible = False
            ' Exit Sub
            'End If
            ' Draw a line above this row.
            If R > 0 Then
                Printer.Line (XFirstColumn, Printer.CurrentY)-(XFirstColumn + .ColWidth(1) + GAP, Printer.CurrentY), vbBlack
            End If
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Print the entries on this row.
            X = xmin + GAP
            For c = 1 To 1 '.Cols - 1
                If frmPrinters.optCenterJust And c < .Cols Then
                    intCenterOffset = ((.ColWidth(c)) / 2) - (Printer.TextWidth(.TextMatrix(R, c)) / 2)
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
                            '                            For cc = 1 To 1 '.Cols - 1
                            '                                TwipPix = .ColWidth(cc) ' * Screen.TwipsPerPixelX
                            '                                X = X + TwipPix + 2 * GAP
                            '                                Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack
                            '                            Next cc
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
                    lngStartY = Printer.CurrentY - GAP + 5
                    lngStartX = Printer.CurrentX - GAP + 5
                    lngEndX = Printer.CurrentX + .ColWidth(c) + GAP  ' - 10
                    lngEndY = Printer.CurrentY + Printer.TextHeight(.TextMatrix(R, c)) + GAP - 5
                    .Row = R
                    .ColSel = .Cols - 1
                    Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor, BF
                    '                    End If
                    Printer.CurrentX = X + intCenterOffset
                    TwipPix = .ColWidth(c) ' * Screen.TwipsPerPixelX
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
                TwipPix = .ColWidth(c) ' * Screen.TwipsPerPixelX
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
                ' Draw lines between the columns.
                '                For c = 1 To 1 '.Cols - 1 '3
                '                    TwipPix = .ColWidth(c) ' * Screen.TwipsPerPixelX '+ GAP
                '                    X = X + TwipPix + 2 * GAP
                '                    Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack 'ymax
                '                Next c
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
        Printer.Line (XFirstColumn, lngYTopOfGrid)-(lngEndX, ymax), vbBlack, B
        X = xmin
        ' Draw lines between the columns.
        '        For c = 1 To 1 '.Cols - 1 '3
        '            TwipPix = .ColWidth(c) ' * Screen.TwipsPerPixelX
        '            X = X + TwipPix + 2 * GAP
        '
        '            Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack 'Printer.CurrentY
        '        Next c
    End With
    Printer.EndDoc
End Sub
Private Function BoundedText(ByVal ptr As Object, _
                             ByVal txt As String, _
                             ByVal max_wid As Single) As String
    On Error GoTo EF
    Do While Printer.TextWidth(txt) > max_wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    BoundedText = txt
EF:
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
    Dim strSQL1 As String
    Dim b       As Integer
    On Error GoTo errs
    If Me.ActiveControl.Name = "FlexGridHist" Then Exit Sub
    If txtJobNo.Text = "" Then Exit Sub
    If bolHasTicket = False Then Exit Sub
    ShowData
    strSQL1 = "SELECT * FROM packetentrydb e LEFT JOIN packetlist l ON l.idJobNum=e.idJobNum WHERE l.idJobNum = '" & txtJobNo.Text & "' ORDER BY idDate DESC"
    Set rs = New ADODB.Recordset
    cn_global.CursorLocation = adUseClient
    Set rs = cn_global.Execute(strSQL1)
    FlexGridHist.Redraw = False
    FlexGridHist.Visible = False
    FlexGridHist.ScrollBars = flexScrollBarNone
    FlexGridHist.Clear
    FlexGridHist.Cols = 6
    FlexGridHist.Rows = 0
    With rs
        dtLatestHistDate = Format$(!idDate, strDBDateTimeFormat)
    End With
    rs.MoveLast
    Do Until rs.BOF
        With rs
            Call FillFlexHist(!idAction, !idStatus, !idComment, !idDate, GetFullName(!idCreator), GetFullName(!idUserFrom), GetFullName(!idUserTo), GetFullName(!idUser), !idGUIDEntry)
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
    Call FlexFlipHist("D")
    FlexBoldFirst FlexGridHist
    FlexGridRedrawHeight
    FlexGridHist.ScrollBars = flexScrollBarVertical
    FlexGridHist.Visible = True
    FlexGridHist.Redraw = True
    FlexGridHist.TopRow = FlexHistLastTopRow
    FlexGridHist.CellPictureAlignment = flexAlignCenterCenter
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
        StatusBar1.Panels.Item(1).Text = "This packet has been Filed by " & GetFullName(strCurUser) & ". Please re-open the packet if on hand"
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
        StatusBar1.Panels.Item(1).Text = GetFullName(strCurUser) & " has reopened this packet and currently has it on hand."
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
        StatusBar1.Panels.Item(1).Text = "This packet is in transit to " & GetFullName(strUserTo) & "."
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
        StatusBar1.Panels.Item(1).Text = GetFullName(strCurUser) & " currently has this packet onhand."
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
        StatusBar1.Panels.Item(1).Text = "The job packet creator, " & GetFullName(strCurUser) & ", has not yet Sent this job packet to anyone."
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
    Dim strSQL1 As String
    On Error GoTo LeaveSub
    List1.Clear
    ShowData
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT idJobNum FROM packetlist Where idJobNum Like '" & strSearchString & "%' ORDER BY idJobNum"
    Set rs = cn_global.Execute(strSQL1)
    Do Until rs.EOF
        With rs
            List1.AddItem !idJobNum, .AbsolutePosition - 1
            rs.MoveNext
        End With
    Loop
    If rs.RecordCount >= 1 Then
        List1.Visible = True
    ElseIf rs.RecordCount <= 0 Then
        List1.Visible = False
    End If
LeaveSub:
    HideData
End Sub
Public Sub SubmitFile()
    On Error GoTo errs
    If Not DBConcurrent = 1 Then
        ShowBanner vbYellow, "The packet status has changed since last refresh.  Current status updated.", 350
        ClearOptBoxes
        RefreshAll
        SetControls
        Exit Sub
    End If
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim intBlah As Integer
    If Trim$(strTicketComment) = "" Then
        Dim Msg
        Msg = MsgBox("Please enter a comment describing the filing location.", vbOKOnly + vbExclamation, "Note Required")
        optFile.Value = True
        frmComments.Show (vbModal)
        Exit Sub
    End If
    ShowData
    strSQL1 = "select * from packetentrydb WHERE idJobNum = '" & txtJobNo.Text & "'"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        rs.AddNew
        !idAction = "FILED"
        !idUser = strLocalUser
        !idUserFrom = "NULL"
        !idUserTo = "NULL"
        !idComment = strTicketComment
        !idJobNum = txtJobNo.Text
        .Update
        .Close
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

    
    If DBConcurrent = 1 And Err.Number = 0 Then
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
    Dim strSQL1 As String, strSQL2 As String
    Dim intBlah As Integer
    On Error GoTo errs
    If Not DBConcurrent = 1 Then
        ShowBanner vbYellow, "The packet status has changed since last refresh.  Current status updated.", 350
        ClearOptBoxes
        RefreshAll
        SetControls
        Exit Sub
    End If
    ShowData
    strSQL1 = "select * from packetentrydb WHERE idJobNum = '" & txtJobNo.Text & "'"
    strSQL2 = "select * from packetlist WHERE idJobNum = '" & txtJobNo.Text & "'"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .AddNew
        !idAction = "REOPENED"
        !idUser = strLocalUser
        !idUserFrom = "NULL"
        !idUserTo = "NULL"
        !idComment = strTicketComment
        !idJobNum = txtJobNo.Text
        .Update
        .Close
        rs.Open strSQL2, cn_global, adOpenKeyset, adLockOptimistic
        If !idStatus = "CLOSED" Then
            !idStatus = "OPEN"
            !idCreateDate = !idCreateDate
            .Update
            .Close
        Else
            .Close
        End If
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
    Dim strSQL1 As String, strSQL2 As String
    Dim intBlah As Integer
    On Error GoTo errs
    If Not DBConcurrent = 1 Then
        ShowBanner vbYellow, "The packet status has changed since last refresh.  Current status updated.", 350
        ClearOptBoxes
        RefreshAll
        SetControls
        Exit Sub
    End If
    If Trim$(strTicketComment) = "" Then
        Dim Msg
        Msg = MsgBox("Please enter a comment describing the filing location.", vbOKOnly + vbExclamation, "Note Required")
        optClose.Value = True
        frmComments.Show (vbModal)
        Exit Sub
    End If
    ShowData
    strSQL1 = "select * from packetentrydb WHERE idJobNum = '" & txtJobNo.Text & "'"
    strSQL2 = "select * from packetlist WHERE idJobNum = '" & txtJobNo.Text & "'"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .AddNew
        !idAction = "NULL"
        !idUser = strLocalUser
        !idUserFrom = "NULL"
        !idUserTo = "NULL"
        !idComment = strTicketComment
        !idJobNum = txtJobNo.Text
        .Update
        .Close
        .Open strSQL2, cn_global, adOpenKeyset, adLockOptimistic
        !idStatus = "CLOSED"
        !idCreateDate = !idCreateDate
        .Update
        .Close
    End With
    HideData
    RefreshAfterEdit
    If Err.Number = 0 Then
        ShowBanner colClosed, "Job Packet Closed Successfully."
    Else
    End If
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
    Dim strSQL1     As String
    Dim ConfirmText As String
    On Error GoTo errs
    If Not DBConcurrent = 1 Then
        ShowBanner vbYellow, "The packet status has changed since last refresh.  Current status updated.", 350
        ClearOptBoxes
        RefreshAll
        SetControls
        Exit Sub
    End If
    ShowData
    strSQL1 = "select * from packetentrydb WHERE idJobNum = '" & txtJobNo.Text & "'"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .AddNew
        !idAction = "RECEIVED"
        !idUser = strLocalUser
        !idUserFrom = strUserFrom
        ConfirmText = "Job Packet Received From " & !idUserFrom
        !idUserTo = "NULL"
        !idComment = strTicketComment
        !idJobNum = txtJobNo.Text
        .Update
        .Close
    End With
    HideData
    'cmdSubmit.Enabled = False
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
    SendEmailToQueue "REC", strLocalUser, strUserFrom, txtJobNo.Text, strTicketComment
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
    Dim strSQL1     As String
    Dim ConfirmText As String
    Dim Hits        As Integer
    On Error GoTo errs
    If Not DBConcurrent = 1 Then
        ShowBanner vbYellow, "The packet status has changed since last refresh.  Current status updated.", 350
        ClearOptBoxes
        RefreshAll
        SetControls
        Exit Sub
    End If
    Hits = GetINIValue(strSelectUserTo)
    If Hits = 0 Then
        Call SetINIValue(strSelectUserTo, 1)
    ElseIf Hits >= 1 Then
        Call SetINIValue(strSelectUserTo, (Hits + 1))
    End If
    ShowData
    strSQL1 = "Select * from packetentrydb WHERE idJobNum = '" & txtJobNo.Text & "'"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .AddNew
        !idAction = "INTRANSIT"
        !idUserFrom = strLocalUser
        !idUser = strLocalUser
        !idUserTo = strSelectUserTo
        ConfirmText = "Job Packet sent to " & GetFullName(!idUserTo)
        !idUserFrom = strLocalUser
        !idComment = strTicketComment
        cmbUsers.Visible = False
        lblUser.Visible = False
        cmbUsers.ComboItems.Item(1).Selected = True
        !idJobNum = txtJobNo.Text
        .Update
        .Close
    End With
    HideData
    cmdSubmit.Enabled = False
    SendEmailToQueue "SEND", strLocalUser, strSelectUserTo, txtJobNo.Text, strTicketComment
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
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String, strJobNum As String, strSQL3 As String
    Dim FormatDate, FormatTime As String
    strJobNum = txtJobNo.Text
    On Error GoTo errs
    ShowData
    FormatDate = Format$(Date, strDBDateFormat)
    FormatTime = Format$(Time, "hh:mm:ss")
    strSQL2 = "SELECT idJobNum From packetlist Where idJobNum = '" & strJobNum & "'"
    strSQL1 = "INSERT INTO packetlist (idJobNum,idPartNum,idDrawingNum,idCustPONum,idSalesNum,idStatus,idCreator,idDescription,idPlant,idMailbox)" & " VALUES ('" & Replace$(txtJobNo.Text, "'", "''") & "','" & Replace$(txtPartNoRev.Text, "'", "''") & "','" & Replace$(txtDrawNoRev.Text, "'", "''") & "','" & Replace$(txtCustPoNo.Text, "'", "''") & "','" & Replace$(txtSalesNo.Text, "'", "''") & "','" & "OPEN','" & strLocalUser & "','" & Replace$(txtTicketDescription.Text, "'", "''") & "','" & cmbPlant.Text & "','" & strLocalUser & "')"
    strSQL3 = "INSERT INTO packetentrydb (idJobNum,idAction,idUser,idUserFrom,idUserTo,idComment) VALUES ('" & Replace$(txtJobNo.Text, "'", "''") & "','CREATED','" & strLocalUser & "','NULL','NULL','" & Replace$(strTicketComment, "'", "''") & "')"
    Set rs = New ADODB.Recordset
    cn_global.CursorLocation = adUseClient
    Set rs = cn_global.Execute(strSQL2)
    If rs.RecordCount > 0 Then
        ShowBanner &HC0C0FF, "A Job Packet with that Job Number already exists!", 500
        optCreate.Value = 1
        cmdSubmit.Enabled = False
        txtJobNo.SetFocus
        HideData
        Exit Sub
    Else
        With rs
            Set rs = cn_global.Execute(strSQL1)
            Set rs = cn_global.Execute(strSQL3)
        End With
    End If
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

Public Sub RefreshFields() 'Fills fields, refreshes MyPackets, does not refresh History Grid.
    Dim rs As New ADODB.Recordset
    Dim strSQL1, strSQL2 As String
    On Error GoTo errs
    If txtJobNo.Text = "" Or optCreate.Value = True Or bolHasTicket = False Then Exit Sub
    SetBoxesForEdit "All"
    ShowData
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.packetentrydb LEFT JOIN (ticketdb.packetlist) ON (packetlist.idJobNum=packetentrydb.idJobNum) WHERE packetlist.idJobNum = '" & txtJobNo.Text & "' ORDER BY packetentrydb.idDate DESC"
    Set rs = cn_global.Execute(strSQL1)
    With rs
        txtPartNoRev.Text = !idPartNum
        txtDrawNoRev.Text = !idDrawingNum
        txtCustPoNo.Text = !idCustPoNum
        txtSalesNo.Text = !idSalesNum
        txtCreator.Text = GetFullName(!idCreator)
        txtCreateDate.Text = !idCreateDate
        txtActionDate.Text = !idDate
        strTicketAction = !idAction
        strUserFrom = !idUserFrom
        strUserTo = !idUserTo
        strCurUser = !idUser
        strTicketStatus = !idStatus
        txtTicketAction.Text = !idAction
        txtTicketOwner.Text = GetFullName(!idUser)
        txtTicketDescription.Text = !idDescription
        txtTicketStatus.Text = !idStatus
        strPlant = !idPlant
        cmbPlant.Text = strPlant
        If txtJobNo.Text = "" Then
            DisableBoxes
            tmrRefresher.Enabled = False
        Else
            bolHasTicket = True
            tmrRefresher.Enabled = True
        End If
        If !idComment <> "" Then
            TheX = pbScrollBox.ScaleWidth
            strCommentText = !idComment
            tmrScroll.Enabled = True
        Else
            pbScrollBox.Cls
            strCommentText = ""
            tmrScroll.Enabled = False
        End If
    End With
    'GetMyPackets
    HideData
    Exit Sub
errs:
    If Err.Number = -2147467259 Then
        CommsDown
    Else
        CommsUp
    End If
End Sub
Public Sub RefreshAfterEdit()
    RefreshFields
    GetMyPackets
End Sub
Public Sub ClearAllButJobN()
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
                    SortHits(0, UBound(SortHits, 2)) = Int(.Value)
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
    Dim strSQL1 As String
    Dim i       As Integer
    On Error GoTo errs
    strSQL1 = "select * from users"
    On Error Resume Next
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
    For i = 1 To UBound(strUserIndex, 2)
        cmbUsers.ComboItems.Add , strUserIndex(0, i), strUserIndex(1, i)
        frmReportFilter.cmbUsers.AddItem strUserIndex(1, i), i
        frmRedirect.cmbOwner.AddItem strUserIndex(1, i), i
        frmRedirect.cmbUserTo.AddItem strUserIndex(1, i), i
        frmRedirect.cmbUserFrom.AddItem strUserIndex(1, i), i
        frmUserSelect.cmbUsers.AddItem strUserIndex(1, i), i
        'i = i + 1
    Next i
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
    strReportType = "All Closed Job Packets"
    sAddlMsg = ""
    ShowData
    strSQL1 = "SELECT * FROM packetlist d LEFT JOIN packetentrydb c ON c.idJobNum = d.idJobNum WHERE" & " idDate = (SELECT MAX(idDate) FROM packetentrydb c2 Where c2.idJobNum = d.idJobNum) AND idStatus='CLOSED' ORDER BY idDate DESC"
    cn_global.CursorLocation = adUseClient
    Set rs = cn_global.Execute(strSQL1)
    pBar.Value = 0
    frmpBar.Visible = True
    If rs.RecordCount <= 0 Then
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
            Flexgrid1.TextMatrix(Line, 1) = !idJobNum
            Flexgrid1.TextMatrix(Line, 2) = !idPartNum
            Flexgrid1.TextMatrix(Line, 3) = !idDescription
            Flexgrid1.TextMatrix(Line, 4) = !idSalesNum
            Flexgrid1.TextMatrix(Line, 5) = !idCustPoNum
            Flexgrid1.TextMatrix(Line, 6) = !idCreator
            Flexgrid1.TextMatrix(Line, 7) = !idCreateDate
            If !idAction = "CREATED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80C0FF)
                Flexgrid1.TextMatrix(Line, 9) = "Job packet was CREATED by " & !idCreator
            ElseIf !idAction = "INTRANSIT" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80FF80)
                Flexgrid1.TextMatrix(Line, 9) = !idUserFrom & " SENT the job packet to " & !idUserTo
            ElseIf !idAction = "RECEIVED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80FFFF)
                Flexgrid1.TextMatrix(Line, 9) = !idUser & " RECEIVED the job packet from " & !idUserFrom
            ElseIf !idStatus = "CLOSED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H8080FF)
                Flexgrid1.TextMatrix(Line, 9) = !idUser & " CLOSED the job packet."
            ElseIf !idStatus = "OPEN" And !idAction = "FILED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &HFF8080)
                Flexgrid1.TextMatrix(Line, 9) = !idUser & " FILED the job packet."
            ElseIf !idAction = "REOPENED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &HFF80FF)
                Flexgrid1.TextMatrix(Line, 9) = !idUser & " REOPENED the job packet."
            End If
            Line = Line + 1
            rs.MoveNext
        End With
    Loop
    Flexgrid1.Rows = Line
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

Private Sub ShowAllOpen()
    bolRunning = True
    Dim rs      As New ADODB.Recordset
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
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM packetlist d LEFT JOIN packetentrydb c ON c.idJobNum = d.idJobNum WHERE" & " idDate = (SELECT MAX(idDate) FROM packetentrydb c2 Where c2.idJobNum = d.idJobNum) AND idStatus='OPEN' ORDER BY idDate DESC"
    Set rs = cn_global.Execute(strSQL1)
    If rs.RecordCount <= 0 Then
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
            Flexgrid1.TextMatrix(Line, 1) = !idJobNum
            Flexgrid1.TextMatrix(Line, 2) = !idPartNum
            Flexgrid1.TextMatrix(Line, 3) = !idDescription
            Flexgrid1.TextMatrix(Line, 4) = !idSalesNum
            Flexgrid1.TextMatrix(Line, 5) = !idCustPoNum
            Flexgrid1.TextMatrix(Line, 6) = !idCreator
            Flexgrid1.TextMatrix(Line, 7) = !idCreateDate
            If !idAction = "CREATED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80C0FF)
                Flexgrid1.TextMatrix(Line, 9) = "Job packet was CREATED by " & !idCreator
            ElseIf !idAction = "INTRANSIT" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80FF80)
                Flexgrid1.TextMatrix(Line, 9) = !idUserFrom & " SENT the job packet to " & !idUserTo
            ElseIf !idAction = "RECEIVED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H80FFFF)
                Flexgrid1.TextMatrix(Line, 9) = !idUser & " RECEIVED the job packet from " & !idUserFrom
            ElseIf !idStatus = "CLOSED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &H8080FF)
                Flexgrid1.TextMatrix(Line, 9) = !idUser & " CLOSED the job packet."
            ElseIf !idStatus = "OPEN" And !idAction = "FILED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &HFF8080)
                Flexgrid1.TextMatrix(Line, 9) = !idUser & " FILED the job packet."
            ElseIf !idAction = "REOPENED" Then
                Call FlexGridRowColor(Flexgrid1, Line, &HFF80FF)
                Flexgrid1.TextMatrix(Line, 9) = !idUser & " REOPENED the job packet."
            End If
            Line = Line + 1
            rs.MoveNext
        End With
    Loop
    Flexgrid1.Rows = Line
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
        PrintFlexGrid FlexGridHist
        'PrintFlexGridColor FlexGridHist
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
    'PrintFlexGrid Flexgrid1
    PrintFlexGridColor Flexgrid1
    SizeTheSheet Flexgrid1
End Sub
Private Sub cmdRefresh_Click()
    tmrRefresher.Enabled = False
    tmrRefresher.Enabled = True    '  Reset timer
    Screen.MousePointer = vbHourglass
    DoEvents
    bolRefreshRunning = True
    RefreshAll
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
Public Sub ClearOptBoxes()
    optMove.Value = False
    optReceive.Value = False
    optClose.Value = False
    optFile.Value = False
    optReOpen.Value = False
    optCreate.Value = False
    cmdSubmit.Enabled = False
    bolOptionClicked = False
End Sub
Private Sub cmdSubmit_Click()
    'If Not DBConcurrent Then
    '
    '
    'ShowBanner vbYellow, "The packet status has changed since last refresh.  Current status updated.", 350
    'ClearOptBoxes
    'RefreshAll
    'SetControls
    'Exit Sub
    'End If
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
    intMovement = 0
    intMovementAccel = 1
    intMovementAccelRate = 5
    tmrReSizer.Enabled = True
End Sub
Private Sub FlexGrid1_DblClick()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    DoEvents
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
    mnuAdmin.Visible = False
    mnuPopup.Visible = False
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
    intShpTimerStartWidth = 3000
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
    strServerAddress = "10.35.1.40"
    strUsername = "TicketApp"
    strPassword = "yb4w4"
    cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    intFormHMax = 10500
    intFormHMin = 5535
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
    GetMyPackets False
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call WheelUnHook
    cn_global.Close
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

Private Sub lblClose_Click()
    CloseBanner
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
        If cmdShowMore.Enabled = True Then Call cmdShowMore_Click
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
Private Sub tmrButtonFlasher_Timer()
    Dim iSteps    As Integer
    Dim FadeColor As Long
    iSteps = 255
    If cmdSubmit.Enabled = True Then
        If iStep <= 0 Then iStep = iSteps
        FadeColor = RGB(r1 + (r2 - r1) / iSteps * iStep, g1 + (g2 - g1) / iSteps * iStep, b1 + (b2 - b1) / iSteps * iStep)
        pbSubmitBox.FillColor = FadeColor
        pbSubmitBox.ForeColor = FadeColor
        RoundRect pbSubmitBox.hdc, 7, 5, 145, 50, 10, 10
        iStep = iStep - 8
    Else
        iStep = 0
        pbSubmitBox.FillColor = pbSubmitBox.BackColor
        pbSubmitBox.ForeColor = pbSubmitBox.BackColor
        RoundRect pbSubmitBox.hdc, 7, 5, 145, 50, 10, 10
        Cls
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
                frmConfirm.Top = intSliderMax
                bolWaitToClose = True
                bolOpenConfirm = False
                intConfirmMovement = 0
                'Exit Sub
            End If
        ElseIf bolOpenConfirm = False Then 'Close
            frmConfirm.Top = frmConfirm.Top - intConfirmMovement
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
            If sngShapeResize > dTimer.Width Then
                dTimer.Width = 0
                frmConfirm.Cls
                RoundRect frmConfirm.hdc, (Border.Left / Screen.TwipsPerPixelY), (Border.Top / Screen.TwipsPerPixelY), ((Border.Left / Screen.TwipsPerPixelY) + (Border.Width / Screen.TwipsPerPixelY)), ((Border.Top / Screen.TwipsPerPixelY) + (Border.Height / Screen.TwipsPerPixelY)), 10, 10
                frmConfirm.CurrentX = lblConfirm.Left
                frmConfirm.CurrentY = lblConfirm.Top
                frmConfirm.ForeColor = lblConfirm.ForeColor
                frmConfirm.DrawStyle = 0
                frmConfirm.Font.Name = "Tahoma"
                frmConfirm.Font.Size = 11
                frmConfirm.FontTransparent = True
                frmConfirm.Print lblConfirm.Caption
            Else
                frmConfirm.Cls
                dTimer.Width = dTimer.Width - sngShapeResize
                frmConfirm.Line (dTimer.Left, dTimer.Top)-(dTimer.Left + dTimer.Width, dTimer.Top + 70), dTimer.Color, BF
                RoundRect frmConfirm.hdc, (Border.Left / Screen.TwipsPerPixelY), (Border.Top / Screen.TwipsPerPixelY), ((Border.Left / Screen.TwipsPerPixelY) + (Border.Width / Screen.TwipsPerPixelY)), ((Border.Top / Screen.TwipsPerPixelY) + (Border.Height / Screen.TwipsPerPixelY)), 10, 10
                frmConfirm.CurrentX = lblConfirm.Left
                frmConfirm.CurrentY = lblConfirm.Top
                frmConfirm.ForeColor = lblConfirm.ForeColor
                frmConfirm.DrawStyle = 0
                frmConfirm.Font.Name = "Tahoma"
                frmConfirm.Font.Size = 11
                frmConfirm.FontTransparent = True
                frmConfirm.Print lblConfirm.Caption
            End If
            dTimer.Left = frmConfirm.Width / 2 - dTimer.Width / 2
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
    intMovementAccel = intMovementAccel + intMovementAccelRate
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
