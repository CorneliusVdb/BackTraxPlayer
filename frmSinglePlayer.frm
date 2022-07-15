VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{7053654B-A6C9-4C60-B4AA-CB8D1BCFC2C0}#1.0#0"; "cpvslider.ocx"
Begin VB.Form frmSinglePlayer 
   BackColor       =   &H007C4E49&
   Caption         =   "LilacPro Music Player"
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "frmSinglePlayer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmSinglePlayer.frx":0442
   ScaleHeight     =   10665
   ScaleWidth      =   5880
   Begin VB.PictureBox picBg 
      BackColor       =   &H00EB7D58&
      Height          =   615
      Left            =   6585
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   15
      Top             =   9165
      Width           =   855
   End
   Begin VB.Timer tmrLogo 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6900
      Top             =   2835
   End
   Begin Threed.SSPanel sspBottom 
      Height          =   510
      Left            =   60
      TabIndex        =   10
      Top             =   10080
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   900
      _Version        =   131074
      BackStyle       =   1
      BevelOuter      =   0
      Begin Threed.SSCheck chkContinue 
         Height          =   255
         Left            =   3195
         TabIndex        =   12
         Top             =   15
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   12640511
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Continues Play"
         Value           =   1
      End
      Begin VB.Image cmdAdd 
         Height          =   270
         Left            =   0
         Picture         =   "frmSinglePlayer.frx":1CCC94
         Top             =   135
         Width           =   960
      End
      Begin VB.Image cmdSetupDevice 
         Height          =   435
         Left            =   4905
         Picture         =   "frmSinglePlayer.frx":1CD544
         Stretch         =   -1  'True
         ToolTipText     =   "Setup Device..."
         Top             =   0
         Width           =   450
      End
      Begin VB.Image cmdOpenPlaylist 
         Height          =   270
         Left            =   960
         Picture         =   "frmSinglePlayer.frx":1CFF96
         Top             =   135
         Width           =   960
      End
      Begin VB.Image cmdSavePlaylist 
         Height          =   270
         Left            =   1920
         Picture         =   "frmSinglePlayer.frx":1D08A7
         Top             =   135
         Width           =   960
      End
      Begin Threed.SSCheck chk5SecMix 
         Height          =   255
         Left            =   3195
         TabIndex        =   11
         Top             =   255
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   12640511
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mix in 5 sec"
         Value           =   1
      End
   End
   Begin MSComDlg.CommonDialog cmdFiles 
      Left            =   8265
      Top             =   1425
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer TimerPlay 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6900
      Top             =   2280
   End
   Begin VB.Timer tmrSlider 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6900
      Top             =   1350
   End
   Begin VB.Timer TimerP1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6900
      Top             =   1800
   End
   Begin VB.Timer TimerP1Level 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6900
      Top             =   810
   End
   Begin VB.Timer TimerF1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6900
      Top             =   345
   End
   Begin MSComctlLib.ListView lvFiles 
      DragIcon        =   "frmSinglePlayer.frx":1D11D8
      Height          =   7740
      Left            =   30
      TabIndex        =   2
      Top             =   2325
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   13653
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   1789711
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2055
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   3625
      _Version        =   131074
      BackColor       =   262722
      PictureFrames   =   1
      Picture         =   "frmSinglePlayer.frx":1FD18A
      BevelWidth      =   2
      BevelOuter      =   1
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   165
         TabIndex        =   4
         Top             =   240
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   767
         _Version        =   131074
         BackColor       =   8146505
         BackStyle       =   1
         BevelOuter      =   0
         Begin VB.Label lblBitRate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   225
            Left            =   1320
            TabIndex        =   29
            Top             =   195
            Width           =   555
         End
         Begin VB.Label lblType 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   225
            Left            =   1320
            TabIndex        =   28
            Top             =   45
            Width           =   555
         End
         Begin VB.Label lblTimePlayed 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-00:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   435
            Left            =   -360
            TabIndex        =   5
            Top             =   -45
            Width           =   1560
         End
      End
      Begin Threed.SSPanel sspSongTitle 
         Height          =   300
         Left            =   225
         TabIndex        =   1
         Top             =   765
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         _Version        =   131074
         ForeColor       =   65535
         BackColor       =   32768
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "You can win if you want New Ve"
         BorderWidth     =   0
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         FloodColor      =   0
      End
      Begin Slider2.cpvSlider cpvVolume 
         Height          =   240
         Left            =   3765
         Top             =   1605
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   423
         BackColor       =   3092271
         SliderIcon      =   "frmSinglePlayer.frx":22B9AE
         Orientation     =   0
         RailPicture     =   "frmSinglePlayer.frx":22BB88
         ShowValueTip    =   0   'False
         Max             =   10000
         Value           =   10000
      End
      Begin Slider2.cpvSlider cpvSlider1 
         Height          =   240
         Left            =   285
         Top             =   1170
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   423
         BackColor       =   3092271
         SliderIcon      =   "frmSinglePlayer.frx":22BBA4
         Orientation     =   0
         RailPicture     =   "frmSinglePlayer.frx":22BD7E
         ShowValueTip    =   0   'False
      End
      Begin Threed.SSPanel sspDevice 
         Height          =   225
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   397
         _Version        =   131074
         ForeColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
         BevelOuter      =   0
         Alignment       =   3
      End
      Begin Threed.SSPanel sspArtist 
         Height          =   240
         Left            =   3885
         TabIndex        =   13
         Top             =   810
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   423
         _Version        =   131074
         ForeColor       =   65280
         BackColor       =   32768
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Don Williams and"
         BorderWidth     =   0
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
         FloodColor      =   0
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   285
         Left            =   2115
         TabIndex        =   21
         Top             =   330
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         _Version        =   131074
         BackColor       =   0
         BevelOuter      =   1
         Begin Threed.SSPanel SSPanel5 
            Height          =   225
            Left            =   45
            TabIndex        =   26
            Top             =   45
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   397
            _Version        =   131074
            BackStyle       =   1
            BevelOuter      =   0
         End
         Begin LilacProBackTraxPlayer.ucSlider sspVolL 
            Height          =   105
            Left            =   120
            TabIndex        =   22
            Top             =   30
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   185
            DropDownCtrl    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   0
            BarColor        =   4
         End
         Begin LilacProBackTraxPlayer.ucSlider sspVolR 
            Height          =   105
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   185
            DropDownCtrl    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   0
            BarColor        =   4
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   5.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   180
            Left            =   60
            TabIndex        =   25
            Top             =   135
            Width           =   135
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "L"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   5.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   180
            Left            =   60
            TabIndex        =   24
            Top             =   15
            Width           =   135
         End
      End
      Begin VB.Image cmdRepeat 
         Height          =   255
         Left            =   4800
         Picture         =   "frmSinglePlayer.frx":22BD9A
         ToolTipText     =   "Repeat song"
         Top             =   1170
         Width           =   480
      End
      Begin VB.Image cmdReset 
         Height          =   255
         Left            =   3780
         Picture         =   "frmSinglePlayer.frx":22C22D
         ToolTipText     =   "Sort Filename"
         Top             =   1170
         Width           =   480
      End
      Begin VB.Image cmdShuffle 
         Height          =   255
         Left            =   4290
         Picture         =   "frmSinglePlayer.frx":22C83B
         ToolTipText     =   "Shuffle Playlist"
         Top             =   1170
         Width           =   480
      End
      Begin VB.Image cmdSortAZ 
         Height          =   255
         Left            =   3255
         Picture         =   "frmSinglePlayer.frx":22CCC4
         ToolTipText     =   "Sort A-Z"
         Top             =   1170
         Width           =   480
      End
      Begin VB.Image cmdNext 
         Height          =   480
         Left            =   2295
         Picture         =   "frmSinglePlayer.frx":22D164
         Top             =   1470
         Width           =   480
      End
      Begin VB.Image cmdStop 
         Height          =   480
         Left            =   1800
         Picture         =   "frmSinglePlayer.frx":22D6E7
         Top             =   1470
         Width           =   480
      End
      Begin VB.Image cmdPause 
         Height          =   480
         Left            =   1290
         Picture         =   "frmSinglePlayer.frx":22DC3E
         Top             =   1470
         Width           =   480
      End
      Begin VB.Image cmdPlay 
         Height          =   480
         Left            =   795
         Picture         =   "frmSinglePlayer.frx":22E1A8
         Top             =   1470
         Width           =   480
      End
      Begin VB.Image cmdPrevious 
         Height          =   480
         Left            =   300
         Picture         =   "frmSinglePlayer.frx":22E713
         Top             =   1470
         Width           =   480
      End
      Begin VB.Image imgSound 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3480
         Picture         =   "frmSinglePlayer.frx":22ECB3
         Top             =   1590
         Width           =   300
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   285
      Left            =   8895
      TabIndex        =   16
      Top             =   2430
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      _Version        =   131074
      BackStyle       =   1
      BevelOuter      =   1
      Begin MSComctlLib.ProgressBar pgLeft 
         Height          =   90
         Left            =   195
         TabIndex        =   17
         Top             =   45
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   159
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ProgressBar pgRight 
         Height          =   90
         Left            =   195
         TabIndex        =   18
         Top             =   150
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   159
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   60
         TabIndex        =   20
         Top             =   135
         Width           =   75
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   60
         TabIndex        =   19
         Top             =   15
         Width           =   75
      End
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Index           =   6
      Left            =   14055
      Picture         =   "frmSinglePlayer.frx":22F0E5
      Top             =   1260
      Width           =   480
   End
   Begin VB.Image cmdRepeatO 
      Height          =   255
      Left            =   7785
      Picture         =   "frmSinglePlayer.frx":22F9AF
      Top             =   8700
      Width           =   480
   End
   Begin VB.Image cmdRepeatF 
      Height          =   255
      Left            =   8625
      Picture         =   "frmSinglePlayer.frx":22FE56
      Top             =   8700
      Width           =   480
   End
   Begin VB.Label lblTimeLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "/00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   390
      Left            =   9225
      TabIndex        =   27
      Top             =   810
      Width           =   1005
   End
   Begin VB.Label lblWork 
      AutoSize        =   -1  'True
      Caption         =   "WORK..."
      Height          =   195
      Left            =   8370
      TabIndex        =   14
      Top             =   510
      Width           =   645
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   13365
      Picture         =   "frmSinglePlayer.frx":2302E9
      Top             =   1290
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Index           =   4
      Left            =   13530
      Picture         =   "frmSinglePlayer.frx":230BB3
      Top             =   765
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Index           =   3
      Left            =   12825
      Picture         =   "frmSinglePlayer.frx":23147D
      Top             =   765
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Index           =   2
      Left            =   12060
      Picture         =   "frmSinglePlayer.frx":231D47
      Top             =   765
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Index           =   1
      Left            =   11280
      Picture         =   "frmSinglePlayer.frx":232611
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Index           =   0
      Left            =   10545
      Picture         =   "frmSinglePlayer.frx":232A53
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTotTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   4830
      TabIndex        =   9
      Top             =   2100
      Width           =   600
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Left            =   4065
      TabIndex        =   8
      Top             =   2100
      Width           =   795
   End
   Begin VB.Image cmdResetF 
      Height          =   255
      Left            =   8595
      Picture         =   "frmSinglePlayer.frx":232E95
      Top             =   8055
      Width           =   480
   End
   Begin VB.Image cmdResetO 
      Height          =   255
      Left            =   7755
      Picture         =   "frmSinglePlayer.frx":2334A3
      Top             =   8055
      Width           =   480
   End
   Begin VB.Image cmdSetupDeviceF 
      Height          =   435
      Left            =   8760
      Picture         =   "frmSinglePlayer.frx":233AFF
      Stretch         =   -1  'True
      ToolTipText     =   "Setup Device..."
      Top             =   7125
      Width           =   450
   End
   Begin VB.Image cmdSetupDeviceO 
      Height          =   435
      Left            =   7695
      Picture         =   "frmSinglePlayer.frx":236551
      Stretch         =   -1  'True
      ToolTipText     =   "Setup Device..."
      Top             =   7050
      Width           =   450
   End
   Begin VB.Image cmdOpenF 
      Height          =   270
      Left            =   8595
      Picture         =   "frmSinglePlayer.frx":238C27
      Top             =   6270
      Width           =   960
   End
   Begin VB.Image cmdSaveF 
      Height          =   270
      Left            =   8595
      Picture         =   "frmSinglePlayer.frx":239538
      Top             =   6705
      Width           =   960
   End
   Begin VB.Image cmdOpenO 
      Height          =   270
      Left            =   7605
      Picture         =   "frmSinglePlayer.frx":239E69
      Top             =   6240
      Width           =   960
   End
   Begin VB.Image cmdSaveO 
      Height          =   270
      Left            =   7605
      Picture         =   "frmSinglePlayer.frx":23A764
      Top             =   6705
      Width           =   960
   End
   Begin VB.Image cmdAddF 
      Height          =   270
      Left            =   8595
      Picture         =   "frmSinglePlayer.frx":23B073
      Top             =   5805
      Width           =   960
   End
   Begin VB.Image cmdAddO 
      Height          =   270
      Left            =   7605
      Picture         =   "frmSinglePlayer.frx":23B923
      Top             =   5805
      Width           =   960
   End
   Begin VB.Image cmdSortAZF 
      Height          =   255
      Left            =   8595
      Picture         =   "frmSinglePlayer.frx":23C1C4
      Top             =   7725
      Width           =   480
   End
   Begin VB.Image cmdShuffleF 
      Height          =   255
      Left            =   8595
      Picture         =   "frmSinglePlayer.frx":23C664
      Top             =   8355
      Width           =   480
   End
   Begin VB.Image cmdSortAZO 
      Height          =   255
      Left            =   7740
      Picture         =   "frmSinglePlayer.frx":23CAED
      Top             =   7740
      Width           =   480
   End
   Begin VB.Image cmdShuffleO 
      Height          =   255
      Left            =   7755
      Picture         =   "frmSinglePlayer.frx":23D117
      Top             =   8355
      Width           =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Playlist"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   210
      Left            =   45
      TabIndex        =   6
      Top             =   2100
      Width           =   645
   End
   Begin VB.Image cmdPreviousI 
      Height          =   480
      Left            =   7530
      Picture         =   "frmSinglePlayer.frx":23D5B6
      Top             =   3750
      Width           =   480
   End
   Begin VB.Image cmdPlayI 
      Height          =   480
      Left            =   8085
      Picture         =   "frmSinglePlayer.frx":23DB56
      Top             =   3750
      Width           =   480
   End
   Begin VB.Image cmdPauseI 
      Height          =   480
      Left            =   8640
      Picture         =   "frmSinglePlayer.frx":23E0C1
      Top             =   3750
      Width           =   480
   End
   Begin VB.Image cmdStopI 
      Height          =   480
      Left            =   9195
      Picture         =   "frmSinglePlayer.frx":23E62B
      Top             =   3735
      Width           =   480
   End
   Begin VB.Image cmdNextI 
      Height          =   480
      Left            =   9750
      Picture         =   "frmSinglePlayer.frx":23EB82
      Top             =   3750
      Width           =   480
   End
   Begin VB.Image cmdPreviousP 
      Height          =   480
      Left            =   7530
      Picture         =   "frmSinglePlayer.frx":23F105
      Top             =   4305
      Width           =   480
   End
   Begin VB.Image cmdPlayP 
      Height          =   480
      Left            =   8085
      Picture         =   "frmSinglePlayer.frx":23F679
      Top             =   4305
      Width           =   480
   End
   Begin VB.Image cmdPauseP 
      Height          =   480
      Left            =   8640
      Picture         =   "frmSinglePlayer.frx":23FBB0
      Top             =   4305
      Width           =   480
   End
   Begin VB.Image cmdStopP 
      Height          =   480
      Left            =   9195
      Picture         =   "frmSinglePlayer.frx":2400D2
      Top             =   4290
      Width           =   480
   End
   Begin VB.Image cmdNextP 
      Height          =   480
      Left            =   9750
      Picture         =   "frmSinglePlayer.frx":2405F2
      Top             =   4305
      Width           =   480
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6900
      TabIndex        =   3
      Top             =   240
      Width           =   825
   End
   Begin VB.Menu PopMenu 
      Caption         =   "PopMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPop 
         Caption         =   "Play Item"
         Index           =   1
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Jump to file"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPop 
         Caption         =   "View File Info"
         Index           =   4
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Remove Items"
         Index           =   6
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Exit"
         Index           =   8
      End
   End
End
Attribute VB_Name = "frmSinglePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iVol4 As Single
Dim PlayStreamHandle As Long
Dim ShuffleState As Boolean
Dim RepeatState As Boolean
Dim TempPlaylist As String
Dim bPlaylistChanged As Boolean
Dim bPlaylistSaved As Boolean
Dim LastPlaylistSaved As String
Dim Arr() As Integer
Dim totTime As Long
Dim totHH As Long
Dim totMM As Long
Dim totSS As Long
Dim LastPlayed As Integer
Dim NextToPlay As Integer
Dim CurrentlyPlaying As Integer
Dim PlayingListNo As Integer
Dim bShowTimePlayed As Boolean
Dim iShowLevelColor As Integer
Dim iLogo As Integer
Const ScreenWidth As Integer = 5640
Const ScreenHeight As Integer = 11070


Sub ChangeLevelColor()
Dim lColor As Long

iShowLevelColor = iShowLevelColor + 1
If iShowLevelColor > 7 Then iShowLevelColor = 0

'Select Case iShowLevelColor
'  Case 1
'    lColor = vbGreen
'  Case 2
'    lColor = vbYellow
'  Case 3
'    lColor = vbRed
'  Case 4
'    lColor = vbCyan
'  Case 5
'    lColor = vbMagenta
'  Case 6
'    lColor = &H80FF&
'  Case 7
'    lColor = &H647F00
'  Case 8
'    lColor = &HFE76BA
'  Case 9
'    lColor = vbWhite
'End Select

'If iShowLevelColor - 1 > 7 Then
  sspVolL.BarColor = iShowLevelColor
  sspVolR.BarColor = iShowLevelColor
'Else
'  sspVolL.BarColor = iShowLevelColor - 1
'  sspVolR.BarColor = iShowLevelColor - 1
'End If

'Change_pb_ForeColor pgLeft.hWnd, lColor   '&HFF&
'Change_pb_Color pgLeft.hWnd, vbBlack
'Change_pb_ForeColor pgRight.hWnd, lColor  '&HFF&
'Change_pb_Color pgRight.hWnd, vbBlack


End Sub

Private Sub cmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.Picture = cmdAddO.Picture
DoEvents

End Sub

Private Sub cmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

cmdFiles.Filename = ""
cmdFiles.FileTitle = ""
cmdFiles.DefaultExt = ""
Err.Clear
cmdFiles.Filter = Filter

cmdFiles.ShowOpen
If Err.Number <> 0 Then
  Err.Clear
  cmdAdd.Picture = cmdAddF.Picture
  Exit Sub
End If

AddSongToList cmdFiles.Filename
bPlaylistChanged = True
bPlaylistSaved = False
 
cmdAdd.Picture = cmdAddF.Picture
DoEvents

End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.sspSongTitle.Caption = "" Then Exit Sub
cmdNext.Picture = cmdNextP.Picture
DoEvents

End Sub

Private Sub cmdNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TimerPlay.Enabled = False
If Me.sspSongTitle.Caption = "" Then Exit Sub

If Setstate("Next") Then TimerPlay.Enabled = True

End Sub

Private Sub cmdOpenPlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOpenPlaylist.Picture = cmdOpenO.Picture
DoEvents
End Sub

Private Sub cmdOpenPlaylist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

OpenPlaylist
cmdOpenPlaylist.Picture = cmdOpenF.Picture
DoEvents

End Sub

Private Sub cmdPause_Click()

If Me.sspSongTitle.Caption = "" Then Exit Sub

Setstate "Pause"

End Sub

Private Sub cmdPlay_Click()

'if sspSongTitle.TagVariant = "" then exit sub
If lvFiles.ListItems.Count = 0 Then Exit Sub

If sspSongTitle.Caption = "" Then
  If Setstate("Next") Then TimerPlay.Enabled = True: Exit Sub
End If

Setstate "Play"

End Sub

Sub ResetControls()

cmdPlay.Picture = cmdPlayI.Picture
cmdPause.Picture = cmdPauseI.Picture
cmdStop.Picture = cmdStopI.Picture
cmdPrevious.Picture = cmdPreviousI.Picture
cmdNext.Picture = cmdNextI.Picture

End Sub

Function Setstate(Options As String) As Boolean
Dim iVol As Single
Dim iSel As Integer
Dim iL As Integer
Dim iRow As Integer
Dim iTmx As ListItem
Dim iNext As Integer

On Error GoTo err1

ResetControls

Select Case Options
  Case "Play"
    Playsong PlayStreamHandle, True
        
  Case "Pause"
    PauseSong PlayStreamHandle
      
  Case "Stop"
    StopSong PlayStreamHandle
    ResetPlayer
    lblStatus.Caption = "Ready"
    cmdStop.Picture = cmdStopP.Picture
    
  Case "Next"
    Setstate = False
    If lvFiles.ListItems.Count < 1 Then iNext = -1: GoTo EndNext
    'Find an entry where the playingflag is set to "IsPlaying"
    Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
    If Not iTmx Is Nothing Then
      iNext = iTmx.Index + 1
      If iNext > lvFiles.ListItems.Count Then
        If chkContinue.value = ssCBChecked Then
          iNext = 1
        Else
          iNext = -1
        End If
      End If
    Else
      iNext = 1
    End If
    
    If iNext > 0 Then
      'Stop song if its playing
      If sspSongTitle.TagVariant <> "" Then StopSong CLng(sspSongTitle.TagVariant)
      
      PlayStreamHandle = CLng(lvFiles.ListItems(iNext).SubItems(3))
      LoadSong lvFiles.ListItems(iNext), lvFiles.ListItems(iNext).SubItems(1), PlayStreamHandle, lvFiles.ListItems(iNext).SubItems(2), iNext, lvFiles.ListItems(iNext).SubItems(9), lvFiles.ListItems(iNext).SubItems(10)
    End If
    
EndNext:
    cmdNext.Picture = cmdNextI.Picture
    DoEvents
    If iNext = -1 Then Exit Function
    
  Case "Previous"
    Setstate = False
    If lvFiles.ListItems.Count < 1 Then iNext = -1: GoTo EndPrevious
    'Find an entry where the playingflag is set to "IsPlaying"
    Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
    If Not iTmx Is Nothing Then
      iNext = iTmx.Index - 1
      If iNext < 1 Then
        If chkContinue.value = ssCBChecked Then
          iNext = 1
        Else
          iNext = -1
        End If
      End If
    Else
      iNext = 1
    End If
    
    If iNext > 0 Then
      'Stop song if its playing
      If sspSongTitle.TagVariant <> "" Then StopSong CLng(sspSongTitle.TagVariant)
      
      PlayStreamHandle = CLng(lvFiles.ListItems(iNext).SubItems(3))
      LoadSong lvFiles.ListItems(iNext), lvFiles.ListItems(iNext).SubItems(1), PlayStreamHandle, lvFiles.ListItems(iNext).SubItems(2), iNext, lvFiles.ListItems(iNext).SubItems(9), lvFiles.ListItems(iNext).SubItems(10)
    End If
    
EndPrevious:
    cmdNext.Picture = cmdNextI.Picture
    cmdPrevious.Picture = cmdPreviousI.Picture
    DoEvents
    If iNext = -1 Then Exit Function

  Case "First"
  
    lvFiles.ListItems(1).Selected = True
    
    ColorListView lvFiles, vbPlaylistSelColor, lvFiles.SelectedItem.Index, True, vbWhite, True, False
    
    lvFiles.SelectedItem.EnsureVisible
    DoEvents
    cmdNext.Picture = cmdNextI.Picture
    DoEvents
    lvFiles.SelectedItem.EnsureVisible
    DoEvents
    cmdPrevious.Picture = cmdPreviousI.Picture
    DoEvents
    
  Case "FF"
  
  Case "RR"
  
 
End Select

Setstate = True

Exit Function

err1:

MsgBox Err.Description
  
End Function

Sub PauseSong(StreamHandle As Long)

  If lblStatus.Caption = "Playing" Then
    tmrLogo.Enabled = False
    TimerP1Level.Enabled = False
    TimerP1.Enabled = False
    lblStatus.Caption = "Pause"
    cmdPause.Picture = cmdPauseP.Picture
    Call BASS_ChannelPause(StreamHandle)
    sspVolL.value = 0
    sspVolR.value = 0
   ' Me.Icon = imgLogo(6).Picture
    'Call BASS_ChannelPause(chan(4))
  ElseIf lblStatus.Caption = "Pause" Then
    iLogo = -1
    'Me.Icon = imgLogo(4).Picture
    tmrLogo.Enabled = True
    TimerP1Level.Enabled = True
    TimerP1.Enabled = True
    lblStatus.Caption = "Playing"
    cmdPlay.Picture = cmdPlayP.Picture
    Call BASS_ChannelPlay(StreamHandle, BASSFALSE)
    
    'Call BASS_ChannelPlay(chan(4), BASSFALSE)
  End If
        
End Sub

Sub Playsong(StreamHandle As Long, restart As Boolean)
Dim DataLength As Long

'Set the volume according to the slider, as per set by user
iVol = cpvVolume.value / 10
iVol4 = iVol
Call BASS_ChannelSetAttribute(StreamHandle, BASS_ATTRIB_VOL, iVol)

'Play new stream
Call BASS_ChannelPlay(StreamHandle, restart)
'Determine the duration of song
Duration(4) = Format(bassTime.GetDuration(StreamHandle), "0")

DataLength = FileLen(sspSongTitle.Tag)

lblBitRate.Caption = bassTime.GetBitsPerSecond(StreamHandle, DataLength) & " Kbp/s"

'Set Slider max value
cpvSlider1.max = Duration(4)
tmrSlider.Enabled = True
TimerP1.Enabled = True
TimerP1Level.Enabled = True
lblStatus.Caption = "Playing"
'Set the Control pictures
cmdPlay.Picture = cmdPlayP.Picture
iLogo = -1
'Me.Icon = imgLogo(4).Picture
tmrLogo.Enabled = True

'Saves the last song that is playong
SaveSetting regMainKey, regSubKey, "LastPlayEntry", CurrentlyPlaying
'Color the song playing
ColorListView lvFiles, vbPlaylistSelColor, lvFiles.ListItems(CurrentlyPlaying).Index, True, vbWhite, True, False  'vbPlaylistSelColor

Me.Caption = Replace(Trim(sspArtist.Caption & " " & sspSongTitle.Caption), "&&", "&")


DoEvents
   
End Sub

Function GetNextSongToPlay(CurrPlaying As Integer) As Integer
Dim iNext As Integer

If CurrPlaying = lvFiles.ListItems.Count Then
  iNext = lvFiles.ListItems(1).SubItems(4)
Else
  iNext = lvFiles.ListItems(CurrPlaying + 1).SubItems(4)
End If

GetNextSongToPlay = iNext

End Function

Sub LoadSong(SongTitle As String, FileToLoad As String, StreamHandle As Long, sTime As String, ToPlayIndex As Integer, sTitle As String, sArtist As String)
Dim iRow As Integer
Dim iTmx As ListItem
Dim iPos As Integer
Dim iLen As Integer
Dim iSongWidth As Integer
Dim sSTitle As String
Dim sSArtist As String
Dim iArtistWidth As Integer
Dim lArtistsLeft As Integer
Dim FileExt As String
Dim DataLength As Long


On Error Resume Next

lblType.Caption = ""

sspArtist.Caption = ""
sspSongTitle.Caption = ""
lblWork.Caption = ""
lblWork.Font = sspSongTitle.Font
lblWork.FontSize = sspSongTitle.FontSize
lblWork.FontBold = sspSongTitle.FontBold

sTitle = Trim(Replace(sTitle, "&", "&&"))
sArtist = Trim(Replace(sArtist, "&", "&&"))

'Load character for caharacter and test the width. If width > 4800, stop the load and add 3 ...
For iLen = 1 To Len(sTitle)
  lblWork.Caption = lblWork.Caption & Mid(sTitle, iLen, 1)
  If lblWork.Width > 3500 Then  '4600 Then
    sSTitle = lblWork.Caption '& "..."
    Exit For
  End If
Next iLen
If Len(sSTitle) = 0 Then sSTitle = lblWork.Caption

iSongWidth = lblWork.Width

lblWork.Caption = ""
lblWork.Font = sspArtist.Font
lblWork.FontSize = sspArtist.FontSize
lblWork.FontBold = sspArtist.FontBold
lblWork.FontItalic = sspArtist.FontItalic

sspSongTitle.Tag = FileToLoad
sspSongTitle.TagVariant = CStr(StreamHandle)

CurrentlyPlaying = ToPlayIndex

'i = GetSetting(regMainKey, regSubKey, "ShowTimePlayed")
If GetSetting(regMainKey, regSubKey, "ShowTimePlayed") Then
  lblTimePlayed.Caption = "0.00"
Else
  lblTimePlayed.Caption = Replace(sTime, "/", "-")
End If

'Find an entry where the playingflag is set to "IsPlaying"
Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
If Not iTmx Is Nothing Then
  lvFiles.ListItems(iTmx.Index).SubItems(8) = ""
  lvFiles.ListItems(iTmx.Index).Selected = False
End If
''Reset the Selected item except the one crrently playing
'Set itmx = lvFiles.FindItem(CStr(StreamHandle), lvwSubItem, , lvwPartial)
'For iRow = 1 To lvFiles.ListItems.Count
'  lvFiles.ListItems(iRow).Selected = False
'Next iRow


'lvFiles.ListItems(ToPlayIndex).Selected = True
lvFiles.ListItems(ToPlayIndex).SubItems(8) = "IsPlaying"

'Make sure we can see the file playing list in the view...
lvFiles.ListItems(ToPlayIndex).EnsureVisible
  
lvFiles.Enabled = True
Screen.MousePointer = Default
DoEvents

'Load character for caharacter and test the width. If width > 4800, stop the load and add 3 ...
sspSongTitle.Caption = sSTitle

lArtistsLeft = sspSongTitle.Left + iSongWidth + 100

'Also load the artist, if any...
For iLen = 1 To Len(sArtist)
  lblWork.Caption = lblWork.Caption & Mid(sArtist, iLen, 1)
  If lArtistsLeft + lblWork.Width > 4800 Then
    sSArtist = lblWork.Caption '& "..."
    Exit For
'  Else
'  If lblWork.Width > 1320 And lArtistsLeft + lblWork.Width > 5235 Then  '3000
'    sSArtist = lblWork.Caption & "..."
'    Exit For
  End If
Next iLen
If Len(sSArtist) = 0 Then sSArtist = lblWork.Caption

lblWork.Caption = ""


'Also load the artist, if any...
If Len(sSArtist) > 0 Then
  sspArtist.Caption = "(" & sSArtist & ")"
End If

sspArtist.Left = lArtistsLeft
iPos = InStr(Len(FileToLoad) - 6, FileToLoad, ".")

FileExt = UCase(Trim(Mid(FileToLoad, iPos + 1)))
lblType.Caption = FileExt
DataLength = FileLen(sspSongTitle.Tag)
lblBitRate.Caption = bassTime.GetBitsPerSecond(StreamHandle, DataLength) & " Kbp/s"


DoEvents

End Sub

Function CreateFileStreamHandle(sFile As String) As Long
Dim StreamHandle As Long
CreateFileStreamHandle = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, 0)
    
End Function

''''Sub CreateStream(channel As Integer)
''''
''''   Call BASS_StreamFree(chan(channel))
''''   Call BASS_SetDevice(lDeviceSingle)  ' set the device to create stream on
''''   chan(channel) = BASS_StreamCreateFile(BASSFALSE, StrPtr(sspSongTitle.Tag), 0, 0, 0)
''''   'Set the volume according to the slider, as per set by user
''''   iVol = cpvVolume.value / 10000
''''   iVol4 = iVol
''''   Call BASS_ChannelSetAttribute(chan(channel), BASS_ATTRIB_VOL, iVol)
''''
''''End Sub

Sub StopSong(StreamHandle As Long)
On Error Resume Next

Call BASS_ChannelStop(StreamHandle)
Call BASS_ChannelSetPosition(StreamHandle, 0, 0)
tmrLogo.Enabled = False
'Me.Icon = imgLogo(5).Picture
'Call BASS_ChannelStop(chan(4))
'chan(4) = 0

End Sub

Sub ResetPlayer()
Dim ssTime As String

On Error Resume Next

tmrSlider.Enabled = False
TimerP1.Enabled = False
TimerP1Level.Enabled = False
If GetSetting(regMainKey, regSubKey, "ShowTimePlayed") Then
  lblTimePlayed.Caption = "0.00"
Else
 ' If PlayStreamHandle <> 0 Then
    ssTime = Format(bassTime.GetDuration(PlayStreamHandle), "0")
    ssTime = (CInt(ssTime) \ 60) & ":" & Format(CInt(ssTime) Mod 60, "00")
    lblTimePlayed.Caption = "-" & ssTime
 ' Else
 '   lblTimePlayed.Caption = "0.00"
 ' End If
End If
'lblTimePlayed.Caption = "0:00"
'lblTimeLeft.Caption = "/00:00"
cpvSlider1.value = 0
'pgLeft.Value = 0
'pgRight.Value = 0
sspVolL.value = 0
sspVolR.value = 0

End Sub

Sub ClearPlayer()
sspSongTitle.TagVariant = ""
sspSongTitle.Tag = ""
sspSongTitle.Caption = ""
sspArtist.Caption = ""

lblType.Caption = ""
lblBitRate.Caption = ""


lblTimePlayed.Caption = "0:00"
lblTimeLeft.Caption = "/00:00"

End Sub

Private Sub cmdPrevious_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.sspSongTitle.Caption = "" Then Exit Sub

cmdPrevious.Picture = cmdPreviousP.Picture
DoEvents
    
End Sub

Private Sub cmdPrevious_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TimerPlay.Enabled = False
If Me.sspSongTitle.Caption = "" Then Exit Sub

If Setstate("Previous") Then TimerPlay.Enabled = True
End Sub

Sub SavePlaylist(Optional SaveTemp As Variant)
Dim FD
Dim FileToOpen As String
Dim sStr As String

On Error Resume Next

If IsMissing(SaveTemp) Then 'Make sure we do not show the dialog if the form is exited without saving the playlist
  cmdFiles.Filename = ""
  cmdFiles.FileTitle = ""
  Err.Clear
  cmdFiles.Filter = "Playlist (*.pls)|*.pls"
  cmdFiles.DefaultExt = "pls"
  
  cmdFiles.ShowSave
  If Err.Number <> 0 Then
    Err.Clear
    Exit Sub
  End If
  'Set the variable to use...
  FileToOpen = Replace(cmdFiles.Filename, ".pls", "") & ".pls"
Else
  FileToOpen = SaveTemp
End If

FD = FreeFile

'Remember the last filename when saving...
LastPlaylistSaved = FileToOpen

Open FileToOpen For Output As FD
'Print the headings
Print #FD, "[Playlist]"
Print #FD, "NumberOfEntries=" & lvFiles.ListItems.Count
For i = 1 To lvFiles.ListItems.Count
  'Print #FD, "Path" & i & "=" & lvFiles.ListItems(i).SubItems(1)
  Print #FD, "File" & i & "=" & lvFiles.ListItems(i).SubItems(1)
Next i
Close FD

bPlaylistChanged = True
bPlaylistSaved = True

End Sub

Sub OpenPlaylist(Optional SFileToOpen As Variant)
Dim FD
Dim FileToOpen As String
Dim sStr As String
Dim iCnt As Integer

On Error Resume Next

If IsMissing(SFileToOpen) Then
  cmdFiles.Filename = ""
  cmdFiles.FileTitle = ""
  cmdFiles.Filter = "Playlist (*.pls;*.m3u)|*.pls;*.m3u"
  cmdFiles.DefaultExt = "pls;m3u"
  Err.Clear
  cmdFiles.ShowOpen
  
  If Err.Number <> 0 Then
    Err.Clear
    Exit Sub
  End If
  
  FileToOpen = cmdFiles.Filename
  
Else
  FileToOpen = SFileToOpen
End If

lvFiles.ListItems.Clear

If Right(FileToOpen, 3) = "pls" Then
  For iCnt = 1 To Val(ReadIni("Playlist", "NumberOfEntries", FileToOpen))
    AddSongToList ReadIni("Playlist", "File" & iCnt, FileToOpen)
    bPlaylistChanged = True
    bPlaylistSaved = False
  Next iCnt
ElseIf Right(FileToOpen, 3) = "m3u" Then
  FD = FreeFile
  Open FileToOpen For Input As FD
  Do Until (EOF(FD) = True)
    Line Input #FD, sStr
    If Not Left(sStr, 1) = "#" Then
      AddSongToList sStr
      bPlaylistChanged = True
      bPlaylistSaved = False
    End If
    
  Loop
  Close FD
End If

End Sub


Private Sub cmdRepeat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


'''SortList 11
'''
''''CurrentlyPlaying
'''Dim iTmx As ListItem
''''Find an entry where the playingflag is set to "IsPlaying"
'''Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
'''If Not iTmx Is Nothing Then
'''  CurrentlyPlaying = iTmx.index
'''End If
'''
'''lvFiles.ListItems(CurrentlyPlaying).EnsureVisible

If RepeatSong = True Then
  RepeatSong = False
  cmdRepeat.Picture = cmdRepeatF.Picture
Else
  RepeatSong = True
  cmdRepeat.Picture = cmdRepeatO.Picture
End If

lvFiles.Sorted = False
DoEvents


End Sub

Private Sub cmdReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdReset.Picture = cmdResetO.Picture

End Sub

Private Sub cmdReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

SortList 1

lvFiles.Sorted = False
'CurrentlyPlaying
Dim iTmx As ListItem
'Find an entry where the playingflag is set to "IsPlaying"
Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
If Not iTmx Is Nothing Then
  CurrentlyPlaying = iTmx.Index
End If

RenumberList

lvFiles.ListItems(CurrentlyPlaying).EnsureVisible

cmdReset.Picture = cmdResetF.Picture
DoEvents


End Sub

Private Sub cmdSavePlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSavePlaylist.Picture = cmdSaveO.Picture
DoEvents
End Sub

Private Sub cmdSavePlaylist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

SavePlaylist
cmdSavePlaylist.Picture = cmdSaveF.Picture
DoEvents

End Sub

Function GetDefaultSoundDevice() As Long

On Error Resume Next

GetDefaultSoundDevice = -1

lDeviceSingle = GetSetting(regMainKey, regSubKey, "SinglePlay Device")
sspDevice.Caption = GetSetting(regMainKey, regSubKey, "SinglePlay Device Description")

sspDevice.Tag = lDeviceSingle

If lDeviceSingle <> -99 Then GetDefaultSoundDevice = lDeviceSingle

Exit Function

err1:
MsgBox Err.Description

End Function

Private Sub cmdSetupDevice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSetupDevice.Picture = cmdSetupDeviceO.Picture
DoEvents

End Sub

Private Sub cmdSetupDevice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'  For i = 1 To 30
'    If Me.lblStatus(i).Caption = "Playing" Then
'      Exit Sub
'    End If
'  Next i

cmdSetupDevice.Picture = cmdSetupDeviceF.Picture
DoEvents

bSinglePlayer = True

frmSetupSoundCards.Show vbModal

DoEvents
If lDeviceSingle = CLng(sspDevice.Tag) Then Exit Sub

Screen.MousePointer = vbHourglass
Me.Enabled = False

If lDeviceSingle <> -99 Then
  'Free current device loaded
  BASS_SetDevice (CLng(sspDevice.Tag))
  BASS_Free
  ' setup NEW output devices
  BASS_SetDevice (lDeviceSingle)
  If (BASS_Init(lDeviceSingle, 44100, BASS_DEVICE_LATENCY, Me.hWnd, 0) = 0) Then
    MsgBox "Can't initialize device "
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
  End If
  
  'Set the buffer size...
  'Call BASS_GetInfo(info)
  buflen = BASS_GetConfig(BASS_CONFIG_BUFFER)
  If buflen < 1000 Then
    Call BASS_SetConfig(BASS_CONFIG_BUFFER, buflen * 5)  'Make buffer twice as large...
  End If
  
  GetDefaultSoundDevice
  
  'Recreate the streaminfo for each file in the list
  ReCreateStreamInfo

End If
  
Screen.MousePointer = vbDefault
Me.Enabled = True
DoEvents

bSinglePlayer = False

End Sub

Private Sub cmdShuffle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If cmdShuffle.Tag = "ON" Then
'  cmdShuffle.Tag = "OFF"
'  cmdShuffle.Picture = cmdShuffleF.Picture
'  SaveSetting regMainKey, regSubKey, "ShuffleState", 0
'Else
'  cmdShuffle.Tag = "ON"
  cmdShuffle.Picture = cmdShuffleO.Picture
'  SaveSetting regMainKey, regSubKey, "ShuffleState", 1
'End If

'ShuffleState = GetSetting(regMainKey, regSubKey, "ShuffleState")

'ShuffleList


End Sub

Sub ShuffleList(bShuffleList As Boolean)
Dim iRows As Long

For iRows = 1 To lvFiles.ListItems.Count
  If bShuffleList Then
    lvFiles.ListItems.Item(iRows).SubItems(4) = Format(Arr(iRows), "000")
  Else
    lvFiles.ListItems.Item(iRows).SubItems(4) = Format(iRows, "000")
  End If
Next iRows

End Sub

Sub SortList(SortColumn As Integer)

lvFiles.SortKey = SortColumn   '4 = shuffle, 5 = Time, 0 = Title
lvFiles.SortOrder = lvFiles.SortOrder 'Xor 1
' Set Sorted to True to sort the list.
lvFiles.Sorted = True

End Sub

Sub RenumberList()
Dim iR As Integer
Dim iPos As Integer

For iR = 1 To lvFiles.ListItems.Count
  iPos = InStr(1, lvFiles.ListItems(iR).Text, ".") + 1
  lvFiles.ListItems(iR).Text = iR & ". " & Trim(Mid(lvFiles.ListItems(iR).Text, iPos))
  lvFiles.ListItems.Item(iR).SubItems(4) = Format(iR, "000")
Next iR

End Sub

Sub ShuffleArray(MaxValue As Integer)

  Dim colCardsLeft As New Collection
  Dim c As Integer
  Dim rndCard As Integer
  
  If MaxValue = 0 Then Exit Sub
  
  ReDim Arr(1 To MaxValue)
  
  ' initialise collection of cards
  ' t0 make use of handly features of collections
  For c = 1 To MaxValue
    colCardsLeft.Add Str(c)
  Next c
  
  ' now shuffle by placing a card in each position
  ' in the deck.  Using the collection's remove method
  ' to prevent any repeats
  c = 1
  Randomize Timer
  While colCardsLeft.Count > 0
    rndCard = Int(Rnd() * colCardsLeft.Count) + 1
    Arr(c) = CInt(colCardsLeft(rndCard))
    colCardsLeft.Remove (rndCard)
    c = c + 1
  Wend

  ' take a look at the shuffle
'  For c = 1 To MaxValue
'    Debug.Print Arr(c); " ";
'  Next c
  
  'Debug.Print



End Sub

Private Sub cmdShuffle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ShuffleArray lvFiles.ListItems.Count
ShuffleList True
SortList 4


lvFiles.Sorted = False


'CurrentlyPlaying
Dim iTmx As ListItem
'Find an entry where the playingflag is set to "IsPlaying"
Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
If Not iTmx Is Nothing Then
  CurrentlyPlaying = iTmx.Index
End If

RenumberList

lvFiles.ListItems(CurrentlyPlaying).EnsureVisible

cmdShuffle.Picture = cmdShuffleF.Picture



DoEvents
End Sub

Private Sub cmdSortAZ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSortAZ.Picture = cmdSortAZO.Picture
End Sub

Private Sub cmdSortAZ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

SortList 11

'CurrentlyPlaying
Dim iTmx As ListItem
'Find an entry where the playingflag is set to "IsPlaying"
Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
If Not iTmx Is Nothing Then
  CurrentlyPlaying = iTmx.Index
End If

RenumberList

lvFiles.ListItems(CurrentlyPlaying).EnsureVisible

cmdSortAZ.Picture = cmdSortAZF.Picture

lvFiles.Sorted = False
DoEvents


End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.sspSongTitle.Caption = "" Then Exit Sub

Setstate "Stop"

End Sub

Private Sub cpvSlider1_MouseDown(Shift As Integer)
  tmrSlider.Enabled = False
End Sub

Private Sub cpvSlider1_MouseUp(Shift As Integer)
'
  Call BASS_ChannelSetPosition(PlayStreamHandle, BASS_ChannelSeconds2Bytes(PlayStreamHandle, cpvSlider1.value), BASS_POS_BYTE)  ' set the position
  'Call BASS_ChannelSetPosition(PlayStreamHandle, BASS_ChannelSeconds2Bytes(PlayStreamHandle, cpvSlider1.Value), BASS_POS_BYTE)  ' set the position
  tmrSlider.Enabled = True
End Sub

Private Sub cpvVolume_ValueChanged()
Dim iVol As Single

On Error Resume Next

iVol = cpvVolume.value / 10
Call BASS_ChannelSetAttribute(PlayStreamHandle, BASS_ATTRIB_VOL, iVol)
'Call BASS_ChannelSetAttribute(chan(4), BASS_ATTRIB_VOL, iVol)
iVol4 = iVol

End Sub

Private Sub Form_Load()
Dim FileToOpen As String
Dim iList As Integer

'EnableCloseButton Me.hWnd, False


GetDefaultSoundDevice

If lDeviceSingle = -99 Then
  MsgBox "No device was loaded previously. Using default device", vbInformation, "No Device found"
  lDeviceSingle = -1
End If

RepeatSong = False

lblType.Caption = ""
lblBitRate.Caption = ""


'Set the device where we are going to play on...
If BASS_SetDevice(lDeviceSingle) <> 0 Then
  'MsgBox "Error setting Single Player device..."
End If

If BASS_Init(lDeviceSingle, 44100, BASS_DEVICE_LATENCY, frmSinglePlayer.hWnd, 0) = BASSFALSE Then
    'MsgBox "Can't initialize Single Player device..."
End If

'Set the buffer size...
'Call BASS_GetInfo(info)
buflen = BASS_GetConfig(BASS_CONFIG_BUFFER)
If buflen < 2000 Then
  Call BASS_SetConfig(BASS_CONFIG_BUFFER, buflen * 10)   'Make buffer 25 times as large... 5000 = maximum
End If

lvFiles.View = lvwReport
lvFiles.ColumnHeaders.Add , , "Available files ", 4400
lvFiles.ColumnHeaders.Add , , "filename ", 0
lvFiles.ColumnHeaders.Add , , "Time ", 0         'Full time met slash
lvFiles.ColumnHeaders.Add , , "StreamHandle", 0
lvFiles.ColumnHeaders.Add , , "SongNumber", 0
lvFiles.ColumnHeaders.Add , , "Original Order", 0
lvFiles.ColumnHeaders.Add , , "Time", 700
lvFiles.ColumnHeaders.Add , , "OriginalByteTime", 0
lvFiles.ColumnHeaders.Add , , "State", 0 'Play state
lvFiles.ColumnHeaders.Add , , "SongTitle", 0
lvFiles.ColumnHeaders.Add , , "Artist", 0
lvFiles.ColumnHeaders.Add , , "FullSongtitle", 0

lvFiles.ColumnHeaders(7).Alignment = lvwColumnRight
'Set the color of the level indicators
iShowLevelColor = 5 'Will increase to 6, which is orange...
ChangeLevelColor
'Clear the captions
sspSongTitle.Caption = ""
sspSongTitle.Tag = ""
sspSongTitle.TagVariant = ""
lblType.Caption = ""
iLogo = -1
'Set the display time option
bShowTimePlayed = False

On Error Resume Next

''-------------------------
''initialize Picture style
''------------------------
'picBg.BackColor = LV.BackColor
'picBg.ScaleMode = vbTwips
'picBg.BorderStyle = vbBSNone
'picBg.AutoRedraw = True
'picBg.Visible = False
''---------------------------
     

'Preload last played list...
FileToOpen = GetSetting(regMainKey, regSubKey, "LastPlaylist")
If FileToOpen <> "" Then
  OpenPlaylist FileToOpen
  CurrentlyPlaying = GetSetting(regMainKey, regSubKey, "LastPlayEntry")
  'LastPlayed = GetSetting(regMainKey, regSubKey, "LastPlayEntry")
  'CurrentlyPlaying = GetNextSongToPlay(LastPlayed)
  
  'Deselect all by default
  For iList = 1 To lvFiles.ListItems.Count
    lvFiles.ListItems(iList).Selected = False
  Next iList
  'Set focus to the last played song...
  lvFiles.ListItems(CurrentlyPlaying).Selected = True
  ColorListView lvFiles, vbPlaylistSelColor, lvFiles.SelectedItem.Index, True, vbWhite, True, False
  PlayStreamHandle = CLng(lvFiles.SelectedItem.SubItems(3))
  LoadSong lvFiles.SelectedItem, lvFiles.SelectedItem.SubItems(1), PlayStreamHandle, lvFiles.SelectedItem.SubItems(2), lvFiles.SelectedItem.Index, lvFiles.SelectedItem.SubItems(9), lvFiles.SelectedItem.SubItems(10)
  
  For iList = 1 To lvFiles.ListItems.Count
    If lvFiles.ListItems(iList).SubItems(3) = "0" Or lvFiles.ListItems(iList).SubItems(3) = "" Then
      ColorListView lvFiles, vbRed, CLng(iList), False, vbWhite, False, False
    End If
  Next iList
   
End If

Me.Width = ScreenWidth '5640   '5670
Me.Height = ScreenHeight  '11070
Me.top = 0
Me.Left = Screen.Width - Me.Width


End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
  If Me.Height < 4000 Then
    Me.Height = 4000
    Exit Sub
  End If
  Me.Width = ScreenWidth
  
  sspBottom.top = Me.Height - 1000  '1170
  lvFiles.Height = Me.Height - 3345
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

'Stop any song playing...
Setstate "Stop"

'Save the last song playing so we can reload later with correct playlist and position...
If bPlaylistSaved = False Then
  SavePlaylist App.Path & "\tmp.pls"
  SaveSetting regMainKey, regSubKey, "LastPlaylist", App.Path & "\tmp.pls"
Else
  Kill App.Path & "\tmp.pls"
  SaveSetting regMainKey, regSubKey, "LastPlaylist", LastPlaylistSaved
End If

'Conditional test to see if this was called from Main Player screen, or started as loose standing app...
If MainApp Then
  frmPlayer.WindowState = 0
Else
  'Make sure we FREE all the instances of the devices we have allocated
  On Error Resume Next
  
  Dim c As Long
  Dim i As BASS_DEVICEINFO
  Dim iDevCnt As Long
  'Loop to get all devices...
  c = 1      ' device 1 = 1st real device
  While BASS_GetDeviceInfo(c, i)
    If (i.flags And BASS_DEVICE_ENABLED) Then
      c = c + 1
    End If
  Wend
  'Free all devices...
  For iDevCnt = 1 To c - 1
    Call BASS_SetDevice(iDevCnt)
    Call BASS_Free
  Next iDevCnt
  'Free all plugins...
  Call BASS_PluginFree(0)
  'End here...
  End
'''''Clears the ID3V2 Class
''''Set objTag = Nothing

End If



End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label5_Click()
ChangeLevelColor
End Sub

Private Sub Label6_Click()
ChangeLevelColor
End Sub

Private Sub lblTimePlayed_Click()
Dim pos As Single
Dim TimePlayed As String
Dim sTime As String

If bShowTimePlayed Then
  bShowTimePlayed = False
Else
  bShowTimePlayed = True
End If

SaveSetting regMainKey, regSubKey, "ShowTimePlayed", bShowTimePlayed

'If paused pressed, get the time already played...
pos = Format(bassTime.GetPlayingPos(PlayStreamHandle), "0")

sTime = Right(bassTime.GetTime(Duration(4) - pos), 5)
sTime = Format(Left(sTime, 2), "0") & Right(sTime, 3)

TimePlayed = Right(bassTime.GetTime(pos), 5)
TimePlayed = Format(Left(TimePlayed, 2), "0") & Right(TimePlayed, 3)


'If lblStatus.Caption = "Pause" Then
'  pos = Format(bassTime.GetPlayingPos(PlayStreamHandle), "0")
'
'  sTime = Right(bassTime.GetTime(Duration(4) - pos), 5)
'  sTime = Format(Left(sTime, 2), "0") & Right(sTime, 3)
'
'  TimePlayed = Right(bassTime.GetTime(pos), 5)
'  TimePlayed = Format(Left(TimePlayed, 2), "0") & Right(TimePlayed, 3)
  If bShowTimePlayed Then
    'Show Time Played
    lblTimePlayed.Caption = TimePlayed
  Else
    'Show Time left
    lblTimePlayed.Caption = "-" & sTime
  End If

'Else
'  ssTime = Format(bassTime.GetDuration(PlayStreamHandle), "0")
'  ssTime = (ssTime \ 60) & ":" & Format(ssTime Mod 60, "00")
'End If
'
'If GetSetting(regMainKey, regSubKey, "ShowTimePlayed") Then
'  lblTimePlayed.Caption = "0.00"
'Else
'
'  lblTimePlayed.Caption = "-" & ssTime
'End If


End Sub

Private Sub lvFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' When a ColumnHeader object is clicked, the ListView control is                    '
' sorted by the subitems of that column.                                            '
' Set the SortKey to the Index of the ColumnHeader - 1                              '
'====================================================================================
lvFiles.SortKey = ColumnHeader.Index - 1
lvFiles.SortOrder = lvFiles.SortOrder Xor 1
' Set Sorted to True to sort the list.
lvFiles.Sorted = True
End Sub

Private Sub lvFiles_DblClick()
Dim SongTitle As String
Dim Index As Integer


'if a song is playing, stop it first, then load new song
If sspSongTitle.TagVariant <> "" Then StopSong CLng(sspSongTitle.TagVariant)

'Get the songnumber to play
'PlayingListNo = lvFiles.SelectedItem.index
'Set the global variable for testing later in the timer...
PlayStreamHandle = CLng(lvFiles.SelectedItem.SubItems(3))
Index = lvFiles.SelectedItem.Index
'Load the text into the display panel...

LoadSong lvFiles.SelectedItem, lvFiles.SelectedItem.SubItems(1), PlayStreamHandle, lvFiles.SelectedItem.SubItems(2), lvFiles.SelectedItem.Index, lvFiles.SelectedItem.SubItems(9), lvFiles.SelectedItem.SubItems(10)
    
'Color the listview item seletced
'ColorListView lvFiles, vbPlaylistSelColor, lvFiles.SelectedItem.index, False, vbWhite, True

'Play the song selected
Setstate "Play"

End Sub

Private Sub lvFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Dim i As Integer
'
'picBg.Width = lvFiles.Width
'picBg.Height = lvFiles.ListItems(1).Height * (lvFiles.ListItems.Count)
'picBg.ScaleHeight = lvFiles.ListItems.Count
'picBg.ScaleWidth = 1
'picBg.DrawWidth = 1
'picBg.Cls
'
'For i = 1 To lvFiles.ListItems.Count
'  If lvFiles.ListItems(i).Selected = True Then
'    picBg.Line (0, i - 1)-(1, i), vbPlayListSelBackColor, BF
'  Else
'    picBg.Line (0, i - 1)-(1, i), vbPlayListBackColor, BF
'  End If
'Next
'
'picBg.Line (0, i - 1)-(1, i), vbRed, BF
'lvFiles.Picture = picBg.Image




End Sub

Private Sub lvFiles_KeyUp(KeyCode As Integer, Shift As Integer)
Dim iRow As Integer
Dim iMax As Integer
Dim iCnt As Integer
Dim aFind() As String
Dim iTmx As ListItem
Dim sFindStr As String

On Error Resume Next
If KeyCode = 46 Then  'Delete
  iCnt = 1
  iMax = lvFiles.ListItems.Count
  'Load items to be deleted into an array
  For iRow = 1 To iMax
    If lvFiles.ListItems(iRow).Selected Then
      iCnt = iCnt + 1
      ReDim Preserve aFind(iCnt)
      aFind(iCnt - 1) = CStr(lvFiles.ListItems(iRow).SubItems(9))
    End If
  Next iRow
  'Check each item in the array and remove if to do so...
  For iRow = 1 To UBound(aFind)
    If aFind(iRow) <> "" Then
      Set iTmx = lvFiles.FindItem(CStr(aFind(iRow)), lvwSubItem, , lvwPartial)
      
      If Not iTmx Is Nothing Then
        If CLng(lvFiles.ListItems(iTmx.Index).SubItems(3)) <> CLng(sspSongTitle.TagVariant) Then
          Call BASS_StreamFree(CLng(lvFiles.ListItems(iTmx.Index).SubItems(3)))
          lvFiles.ListItems.Remove iTmx.Index
        Else
          If lblStatus.Caption <> "Playing" Then
            Call BASS_StreamFree(CLng(lvFiles.ListItems(iTmx.Index).SubItems(3)))
            lvFiles.ListItems.Remove iTmx.Index
          End If
        End If
      End If
    End If
  Next iRow
  CalcTotTime
    
  If lvFiles.ListItems.Count = 0 Then
    ResetPlayer
    ResetControls
    ClearPlayer
    CurrentlyPlaying = 0
  Else
   RenumberList
  End If
  
  If lvFiles.ListItems.Count > 0 Then
    Set iTmx = lvFiles.FindItem(CLng(sspSongTitle.TagVariant), lvwSubItem, , lvwPartial)
    If Not iTmx Is Nothing Then
      CurrentlyPlaying = iTmx.Index
      lvFiles.ListItems(CurrentlyPlaying).EnsureVisible
    End If
  End If
  
ElseIf KeyCode = 70 And Shift = 2 Then  'Ctrl + F  (Find)
  sFindStr = UCase(InputBox("Please enter Search criteria. ", "Search"))
  If sFindStr = "" Then Exit Sub
  
  For iRow = 1 To lvFiles.ListItems.Count
    If InStr(1, UCase(lvFiles.ListItems(iRow).Text), sFindStr) > 0 Then
      lvFiles.ListItems(iRow).EnsureVisible
      ColorListView lvFiles, vbYellow, CLng(iRow), False, vbWhite, True, True
      Exit For
    End If
  Next iRow
End If

End Sub

Private Sub lvFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
  PopupMenu PopMenu
End If

End Sub

Private Sub lvFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iCnt As Integer
Dim fsMain As New FileSystemObject
Dim fsFolder As folder
Dim fsFile As file
Dim sExt As String
Dim iTmx As ListItem
Dim sFile As String

'Check if this song is moved in the current list, or added from external source

'If Not data.GetFormat(vbCFText) Then Exit Sub
'On Error Resume Next


If Effect = 3 Then
  lvFiles.Sorted = False
  'lvFiles.SortKey = 3
  Call LVDragDropSingle(lvFiles, X, Y)
  'SortList 4
  'lvFiles.Sorted = False
'  RenumberList
'  Exit Sub
Else
  If DirectoryExists(Data.Files.Item(1)) Then
    LoadDirectories Data.Files.Item(1), X, Y
  Else
    For iCnt = 1 To Data.Files.Count
      sExt = Right(Data.Files(iCnt), 4)
      If InStr(1, Filter, sExt) > 0 Then
        AddSongToList Data.Files(iCnt)
        lvFiles.ListItems(lvFiles.ListItems.Count).Selected = True
        Call LVDragDropSingle(lvFiles, X, Y)
      End If
    Next iCnt
  End If
   
   
'  'First check to see if we selected a directory
'  If DirectoryExists(data.Files.Item(1)) Then
'    Set fsFolder = fsMain.GetFolder(data.Files.Item(1))
'    'Load all songs in root folder here...
'    For Each fsFile In fsFolder.Files
'      sExt = Right(fsFile.name, 4)
'      If InStr(1, filter, sExt) > 0 Then
'        AddSongToList fsFolder & "\" & fsFile.name
'        lvFiles.ListItems(lvFiles.ListItems.Count).Selected = True
'        Call LVDragDropSingle(lvFiles, X, Y)
'      End If
'    Next fsFile
    
    
'    'Now load all subfolders as well...
'    For Each fsFolder In fsMain.GetFolder(data.Files.Item(1)).SubFolders
'      If fsFolder.Attributes = Directory Or fsFolder.Attributes = 48 Or fsFolder.Attributes = 2064 Or fsFolder.Attributes = 2096 Then
'
'        For Each fsFile In fsFolder.Files
'          sExt = Right(fsFile.name, 4)
'          If InStr(1, filter, sExt) > 0 Then
'            AddSongToList fsFolder & "\" & fsFile.name
'            lvFiles.ListItems(lvFiles.ListItems.Count).Selected = True
'            Call LVDragDropSingle(lvFiles, X, Y)
'          End If
'        Next fsFile
'      End If
'
'    Next fsFolder
'  Else
'    For iCnt = 1 To data.Files.Count
'      AddSongToList data.Files(iCnt)
'      lvFiles.ListItems(lvFiles.ListItems.Count).Selected = True
'      Call LVDragDropSingle(lvFiles, X, Y)
'    Next iCnt
'  End If
End If


'Find an entry where the playingflag is set to "IsPlaying"
Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
If Not iTmx Is Nothing Then
  CurrentlyPlaying = iTmx.Index
Else
  CurrentlyPlaying = 1
End If

If lvFiles.ListItems.Count > 0 Then
  For iCnt = 1 To lvFiles.ListItems.Count
    If iCnt <> CurrentlyPlaying Then lvFiles.ListItems(iCnt).Selected = False
  Next iCnt
End If

RenumberList

'''''Find an entry where the playingflag is set to "IsPlaying"
''''Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
''''If Not iTmx Is Nothing Then
''''  CurrentlyPlaying = iTmx.index
''''Else
''''  CurrentlyPlaying = 1
''''End If

'lvFiles.ListItems(CurrentlyPlaying).Selected = True
'Color the song playing
ColorListView lvFiles, vbPlaylistSelColor, CLng(CurrentlyPlaying), True, vbWhite, True, False
'lvFiles.ListItems(CurrentlyPlaying).EnsureVisible


bPlaylistSaved = False

End Sub

Function LoadDirectories(DirectoryToProcess As String, ByVal X As Single, ByVal Y As Single)
Dim iCnt As Integer
Dim fsMain As New FileSystemObject
Dim fsFolder As folder
Dim folder As folder
Dim folderlist As Folders
Dim filelist As Files
Dim fsFile As file
Dim sExt As String
Dim iTmx As ListItem
Dim sFile As String

Static running As Boolean

Dim AllDirs As New Collection
Dim next_dir As Integer
Dim dir_name As String
Dim sub_dir As String
Dim i As Integer
Dim txt As String
     
next_dir = 1
AllDirs.Add DirectoryToProcess ' Start here.
Do While next_dir <= AllDirs.Count
   ' Get the next directory to search.
   dir_name = AllDirs(next_dir)
   next_dir = next_dir + 1
   
   ' Read directories from dir_name.
   sub_dir = Dir$(dir_name & "\*", vbDirectory)
   Do While sub_dir <> ""
       ' Add the name to the list if
       ' it is a directory.
       If UCase$(sub_dir) <> "PAGEFILE.SYS" And sub_dir <> "." And sub_dir <> ".." Then
           sub_dir = dir_name & "\" & sub_dir
           On Error Resume Next
           If GetAttr(sub_dir) And vbDirectory Then AllDirs.Add sub_dir
       End If
       sub_dir = Dir$(, vbDirectory)
   Loop
   DoEvents
Loop

'Position of Item being dropped
Dim objFind As ListItem
Dim intIndex As Integer
Set objFind = lvFiles.HitTest(X, Y)

'Retrieve the drop position
intIndex = objFind.Index
If intIndex = 0 Then intIndex = lvFiles.ListItems.Count
intIndex = intIndex - 2
'Loop through each directory and load its files...
For i = 1 To AllDirs.Count
   Set fsFolder = fsMain.GetFolder(AllDirs(i))
   Set folderlist = fsFolder.SubFolders
   Set filelist = fsFolder.Files
   
   For Each fsFile In filelist
     sExt = Right(fsFile.name, 4)
     If InStr(1, Filter, sExt) > 0 Then
       AddSongToList fsFolder & "\" & fsFile.name
       intIndex = intIndex + 1
'       lvFiles.ListItems(intIndex).Selected = True
'       ColorListView lvFiles, vbRed, CLng(intIndex), False, vbWhite, False
       lvFiles.ListItems(lvFiles.ListItems.Count).Selected = True
       Call LVDragDropSingle(lvFiles, X, Y)
     End If
   Next fsFile
Next i

End Function

Function FileCheck(Filename As String) As Boolean

If Dir(Filename) <> "" Then
 FileCheck = True
Else
 FileCheck = False
End If

End Function

Public Function DirectoryExists(Dir As String) As Boolean
  Dim oDir As New Scripting.FileSystemObject
  DirectoryExists = oDir.FolderExists(Dir)
  
End Function

Sub AddSongToList(FileToLoad As String)
'Dim tags As New clsTags
Dim sListName As String
Dim bytes As Long
Dim time As Long
Dim iRow As Long
Dim totRes As Integer
Dim bId3V1Found As Boolean
Dim bId3V2Found As Boolean
Dim bTitleArtistLoaded As Boolean
Dim lStreamHandle As Long
Dim sArtist As String
Dim sTitle As String
Dim MaxVol As Long

On Error Resume Next

'''bId3V1Found = False
'''bId3V2Found = False
'''bTitleArtistLoaded = False
'''' Get ID3v1 Tag Information:
'''With m_cID3v1
'''  .MP3File = FileToLoad
'''  If .HasID3v1Tag Then
'''    bId3V1Found = True
'''    If Trim(.Artist) = "" Then
'''      If Trim(.Title) = "" Then
'''        bId3V1Found = False
'''      Else
'''        sListName = Trim(.Title)
'''        sTitle = .Title
'''      End If
'''    Else
'''      sListName = Trim(.Artist) & " - " & Trim(.Title)
'''      sArtist = .Artist
'''      sTitle = .Title
'''      bTitleArtistLoaded = True
'''    End If
'''  End If
'''End With
'''
'''If Not bTitleArtistLoaded Then
'''  With m_cID3v2
'''    .MP3File = FileToLoad
'''    If .HasID3v2Tag Then
'''      bId3V2Found = True
'''      If Trim(.Artist) = "" Then
'''        If Trim(.Title) = "" Then
'''          bId3V2Found = False
'''        Else
'''          sListName = Trim(.Title)
'''          sTitle = .Title
'''        End If
'''      Else
'''        sListName = Trim(.Artist) & " - " & Trim(.Title)
'''        sArtist = .Artist
'''        sTitle = .Title
'''      End If
'''    End If
'''  End With
'''End If
'''
'''If sArtist = "" And sTitle = "" Then
'''  For i = Len(FileToLoad) To 1 Step -1
'''    If Mid(FileToLoad, i, 1) = "\" Then
'''      sListName = Mid(FileToLoad, i + 1)
'''      iPos = InStr(1, sListName, ".")
'''      sListName = Mid(sListName, 1, iPos - 1)
'''      sArtist = ""
'''      sTitle = sListName
'''      Exit For
'''    End If
'''  Next i
'''End If
GetId3Tags FileToLoad

sListName = Id3TagArr(0)
sTitle = Id3TagArr(1)
sArtist = Id3TagArr(2)


    



'Load info into listview

lStreamHandle = CreateFileStreamHandle(FileToLoad)
If lStreamHandle = 0 Then
    MsgBox "File : " & FileToLoad & Chr(13) & Chr(13) & "Error: Invalid file format", vbExclamation
    Exit Sub
End If

Set mItem = lvFiles.ListItems.Add(, , sListName)
lvFiles.ListItems(lvFiles.ListItems.Count).Text = Format(lvFiles.ListItems.Count, "###") & ". " & lvFiles.ListItems(lvFiles.ListItems.Count).Text

DoEvents
mItem.SubItems(1) = FileToLoad   'Full filename
mItem.SubItems(3) = CStr(lStreamHandle)
'Get the Time (length) for this file...
bytes = BASS_ChannelGetLength(CLng(mItem.SubItems(3)), BASS_POS_BYTE)
time = CLng(BASS_ChannelBytes2Seconds(CLng(mItem.SubItems(3)), bytes))


MaxVol = BASS_ChannelGetLevel(lStreamHandle)

'Store the time in the listview subitem
mItem.SubItems(2) = "/" & (time \ 60) & ":" & Format(time Mod 60, "00")
mItem.SubItems(4) = Format(lvFiles.ListItems.Count, "000")
mItem.SubItems(5) = Format(lvFiles.ListItems.Count, "000")
mItem.SubItems(6) = Mid(mItem.SubItems(2), 2)
mItem.SubItems(7) = CStr(time)
mItem.SubItems(8) = ""
mItem.SubItems(9) = Trim(sTitle)
mItem.SubItems(10) = Trim(sArtist)
mItem.SubItems(11) = sListName

'If lvFiles.ListItems(lvFiles.ListItems.Count).SubItems(3) = "0" Or lvFiles.ListItems(lvFiles.ListItems.Count).SubItems(3) = "" Then
'  ColorListView lvFiles, vbRed, lvFiles.ListItems.Count, False, vbWhite, True
'End If

DoEvents

CalcTotTime

End Sub

Sub ReCreateStreamInfo()
Dim iCnt As Integer

For iCnt = 1 To lvFiles.ListItems.Count
  lvFiles.ListItems(iCnt).SubItems(3) = CStr(CreateFileStreamHandle(lvFiles.ListItems(iCnt).SubItems(1)))
Next iCnt
  
End Sub

Function CalcTotTime()

totTime = 0
'read through list and recalc the times...
For iRow = 1 To lvFiles.ListItems.Count
  totTime = totTime + CLng(lvFiles.ListItems(iRow).SubItems(7))
Next iRow

totHH = totTime \ 3600

totRes = totTime - (totHH * 3600)
totMM = totRes \ 60
totSS = totRes Mod 60

If totHH < 1 And totMM < 10 Then
  lblTotTime.Caption = totHH & ":" & totMM & ":" & Format(totSS, "00")
Else
  lblTotTime.Caption = totHH & ":" & Format(totMM, "00") & ":" & Format(totSS, "00")
End If

End Function

Private Sub mnuPop_Click(Index As Integer)
Dim sNo As String
Dim iPos As Integer
Dim FileToLoad As String
Dim SongTitle As String


On Error Resume Next

Select Case Index
  Case 1 'Play
    'if a song is playing, stop it first, then load new song
    If sspSongTitle.TagVariant <> "" Then StopSong CLng(sspSongTitle.TagVariant)
    
    'Set the global variable for testing later in the timer...
    PlayStreamHandle = CLng(lvFiles.SelectedItem.SubItems(3))
    'Load the text into the display panel...
    LoadSong lvFiles.SelectedItem, lvFiles.SelectedItem.SubItems(1), PlayStreamHandle, lvFiles.SelectedItem.SubItems(2), lvFiles.SelectedItem.Index, lvFiles.SelectedItem.SubItems(9), lvFiles.SelectedItem.SubItems(10)
    
    'Play the song selected
    Setstate "Play"

  Case 2 'Jump to
  
  
  Case 4 'View tags
    frmTestMp3Tags.Show vbModal
    'Get the tag info for this file, and update the listview's info
    FileToLoad = lvFiles.SelectedItem.SubItems(1)
    iPos = InStr(1, lvFiles.SelectedItem, ".")
    sNo = Mid(lvFiles.SelectedItem, 1, iPos)
    
    GetId3Tags FileToLoad
    
    lvFiles.SelectedItem.Text = sNo & " " & Id3TagArr(0)
    
    DoEvents
    lvFiles.SelectedItem.SubItems(1) = FileToLoad   'Full filename
    lvFiles.SelectedItem.SubItems(9) = Trim(Id3TagArr(1))
    lvFiles.SelectedItem.SubItems(10) = Trim(Id3TagArr(2))
    lvFiles.SelectedItem.SubItems(11) = Id3TagArr(0)
    
    If lvFiles.SelectedItem.SubItems(8) = "IsPlaying" Then
      'Set the global variable for testing later in the timer...
      PlayStreamHandle = CLng(lvFiles.SelectedItem.SubItems(3))
      'Load the text into the display panel...
      LoadSong lvFiles.SelectedItem, lvFiles.SelectedItem.SubItems(1), PlayStreamHandle, lvFiles.SelectedItem.SubItems(2), lvFiles.SelectedItem.Index, lvFiles.SelectedItem.SubItems(9), lvFiles.SelectedItem.SubItems(10)
    End If
    
    

  Case 6 'Remove from list
    RemoveFile
  Case 8
    Unload Me
    
End Select

End Sub

Private Sub RemoveFile()
Dim iRow As Integer
Dim iMax As Integer
Dim iCnt As Integer
Dim aFind() As String
Dim iTmx As ListItem
Dim sFindStr As String

iCnt = 1
iMax = lvFiles.ListItems.Count
'Load items to be deleted into an array
For iRow = 1 To iMax
  If lvFiles.ListItems(iRow).Selected Then
    iCnt = iCnt + 1
    ReDim Preserve aFind(iCnt)
    aFind(iCnt - 1) = CStr(lvFiles.ListItems(iRow).SubItems(9))
  End If
Next iRow
'Check each item in the array and remove if to do so...
For iRow = 1 To UBound(aFind)
  If aFind(iRow) <> "" Then
    Set iTmx = lvFiles.FindItem(CStr(aFind(iRow)), lvwSubItem, , lvwPartial)
    
    If Not iTmx Is Nothing Then
      If CLng(lvFiles.ListItems(iTmx.Index).SubItems(3)) <> CLng(sspSongTitle.TagVariant) Then
        Call BASS_StreamFree(CLng(lvFiles.ListItems(iTmx.Index).SubItems(3)))
        lvFiles.ListItems.Remove iTmx.Index
      Else
        If lblStatus.Caption <> "Playing" Then
          Call BASS_StreamFree(CLng(lvFiles.ListItems(iTmx.Index).SubItems(3)))
          lvFiles.ListItems.Remove iTmx.Index
        End If
      End If
    End If
  End If
Next iRow
CalcTotTime
  
If lvFiles.ListItems.Count = 0 Then
  ResetPlayer
  ResetControls
  ClearPlayer
  CurrentlyPlaying = 0
Else
 RenumberList
End If

If lvFiles.ListItems.Count > 0 Then
  Set iTmx = lvFiles.FindItem(CLng(sspSongTitle.TagVariant), lvwSubItem, , lvwPartial)
  If Not iTmx Is Nothing Then
    CurrentlyPlaying = iTmx.Index
    lvFiles.ListItems(CurrentlyPlaying).EnsureVisible
  End If
End If


End Sub

Private Sub pgLeft_Click()
ChangeLevelColor
End Sub

Private Sub pgRight_Click()
ChangeLevelColor
End Sub

Private Sub SSPanel5_Click()
ChangeLevelColor
End Sub

Private Sub TimerP1Level_Timer()
Dim LevelInd As Long
Dim LevelAve As Long
Dim LeftChan As Integer
Dim RightChan As Integer

On Error Resume Next

LevelInd = BASS_ChannelGetLevel(PlayStreamHandle)
'LevelInd = BASS_ChannelGetLevel(chan(4))

LeftChan = Round(LoWord(LevelInd) * iVol4)
RightChan = Round(HiWord(LevelInd) * iVol4)

'pgLeft.Value = LeftChan / 100 / 100 * 29
'pgRight.Value = RightChan / 100 / 100 * 29

sspVolL.value = LeftChan / 100 / 100 * 29
sspVolR.value = RightChan / 100 / 100 * 30

End Sub

Private Sub TimerF1_Timer()
'If sspSongTitle(Player1Index).ForeColor = vbBlack Then 'sspSongTitle(Player1Index).Tag Then
'  sspSongTitle(Player1Index).BackColor = vbBlack
'  sspSongTitle(Player1Index).ForeColor = vbYellow
'Else
'  sspSongTitle(Player1Index).ForeColor = vbBlack  'sspSongTitle(Player1Index).Tag
'  sspSongTitle(Player1Index).BackColor = vbYellow
'End If

TimerF1.Enabled = False
If Not RepeatSong Then
  LoadNextSong
Else
  TimerP1.Enabled = True
  TimerPlay.Enabled = True
End If

End Sub

Sub LoadNextSong()
Dim iNext As Integer
Dim iTmx As ListItem

iNext = 0

'Find an entry where the playingflag is set to "IsPlaying"
Set iTmx = lvFiles.FindItem("IsPlaying", lvwSubItem, , lvwPartial)
If Not iTmx Is Nothing Then
  iNext = iTmx.Index + 1
  If iNext > lvFiles.ListItems.Count Then
    If chkContinue.value = ssCBChecked Then
      iNext = 1
    Else
      iNext = -1
    End If
  End If
End If

If iNext > 0 Then
  PlayStreamHandle = CLng(lvFiles.ListItems(iNext).SubItems(3))
  LoadSong lvFiles.ListItems(iNext), lvFiles.ListItems(iNext).SubItems(1), PlayStreamHandle, lvFiles.ListItems(iNext).SubItems(2), iNext, lvFiles.ListItems(iNext).SubItems(9), lvFiles.ListItems(iNext).SubItems(10)
  Setstate "Play"
End If

End Sub

Private Sub TimerP1_Timer()
On Error Resume Next
Dim pos As Single
Dim TimeLeft As Long
Dim TimeElapsedPerc As Long
Dim sTime As String
Dim StreamHandle As Long
Dim TimePlayed As String


If BASS_ChannelIsActive(PlayStreamHandle) = 0 Then
  pos = -1 ' reached the end
Else
  'Get current possition of playing...
  pos = Format(bassTime.GetPlayingPos(PlayStreamHandle), "0")
End If

'Check if END Reached. Stop timer, stop playing and reset button...
If pos = -1 Then
  If RepeatSong Then
    TimerF1.Interval = 1
    TimerF1.Enabled = True
    Exit Sub
  End If
  'Test if the song needs to continue, but only at the end...
  If chkContinue.value = -1 Then
    If chk5SecMix.value <> -1 Then
      TimerF1.Interval = 1
      TimerF1.Enabled = True
    End If
  End If
  'Stop current stream playing
  TimerP1.Enabled = False
  StopSong StreamHandle
  ResetPlayer
  Exit Sub
End If

'Calculate the progress bar's position, as well as the time left to display
sTime = Right(bassTime.GetTime(Duration(4) - pos), 5)
sTime = Format(Left(sTime, 2), "0") & Right(sTime, 3)

TimePlayed = Right(bassTime.GetTime(pos), 5)
TimePlayed = Format(Left(TimePlayed, 2), "0") & Right(TimePlayed, 3)

If bShowTimePlayed Then
  'Show Time Played
  lblTimePlayed.Caption = TimePlayed
Else
  'Show Time left
  lblTimePlayed.Caption = "-" & sTime
End If

lblTimePlayed.Tag = TimePlayed
If Left(sTime, 1) = "0" Then
  If Not RepeatSong Then
    If chk5SecMix.value = -1 Then 'Mix 5 second prior to end
      If Val(Right(sTime, 2)) < 5 Then
        If chkContinue.value = -1 Then
          TimerF1.Enabled = True
        End If
'      Else
'        TimerF1.Enabled = True
      End If
    End If
  Else
    'TimerF1.Enabled = True
  End If
End If

End Sub

Private Sub TimerPlay_Timer()
  
  TimerPlay.Enabled = False
  Setstate "Play"
    
End Sub

Private Sub tmrLogo_Timer()

iLogo = iLogo + 1
If iLogo > 1 Then iLogo = 0
Me.Icon = imgLogo(iLogo).Picture

End Sub

Private Sub tmrSlider_Timer()
On Error Resume Next
'
cpvSlider1.value = BASS_ChannelBytes2Seconds(PlayStreamHandle, BASS_ChannelGetPosition(PlayStreamHandle, BASS_POS_BYTE)) ' update position
'sspVolumeL.DrawBar BASS_ChannelBytes2Seconds(PlayStreamHandle, BASS_ChannelGetPosition(PlayStreamHandle, BASS_POS_BYTE))
'cpvSlider1.value = BASS_ChannelBytes2Seconds(chan(4), BASS_ChannelGetPosition(chan(4), BASS_POS_BYTE)) ' update position
End Sub
