VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayer 
   BackColor       =   &H00000000&
   ClientHeight    =   11610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   26850
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmPlayer.frx":0000
   ScaleHeight     =   11610
   ScaleWidth      =   26850
   Begin Threed.SSPanel sspVol 
      Height          =   450
      Left            =   21570
      TabIndex        =   17
      Top             =   9150
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   794
      _Version        =   131074
      BackColor       =   3092271
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Outline         =   -1  'True
      Alignment       =   6
      Begin LilacProBackTraxPlayer.zzSlider cpvVol 
         Height          =   210
         Left            =   495
         TabIndex        =   73
         Top             =   165
         Width           =   2700
         _extentx        =   4763
         _extenty        =   370
         font            =   "frmPlayer.frx":3545
         slidercolor     =   49152
         maxvalue        =   100
         smallchange     =   20
         largechange     =   100
      End
      Begin VB.Image cmdCloseVol 
         Height          =   285
         Left            =   75
         Picture         =   "frmPlayer.frx":3571
         Stretch         =   -1  'True
         Top             =   105
         Width           =   285
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   225
         Index           =   6
         Left            =   1350
         TabIndex        =   37
         Top             =   1695
         Width           =   450
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "50%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   225
         Index           =   5
         Left            =   2610
         TabIndex        =   36
         Top             =   1245
         Width           =   450
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   4
         Left            =   2445
         TabIndex        =   35
         Top             =   2520
         Width           =   450
      End
      Begin VB.Label lblVolInd 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   60
         TabIndex        =   22
         Top             =   0
         Width           =   2685
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "50%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   225
         Index           =   3
         Left            =   1020
         TabIndex        =   20
         Top             =   2820
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   2430
         TabIndex        =   19
         Top             =   2775
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   18
         Top             =   1695
         Width           =   450
      End
      Begin VB.Label lblVolTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   195
         Left            =   645
         TabIndex        =   21
         Top             =   1995
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0FFC0&
         Index           =   0
         X1              =   2115
         X2              =   2520
         Y1              =   1575
         Y2              =   2265
      End
      Begin VB.Line Line1 
         BorderColor     =   &H008080FF&
         Index           =   1
         X1              =   390
         X2              =   1005
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         Index           =   2
         X1              =   1695
         X2              =   2040
         Y1              =   1695
         Y2              =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0FF&
         Index           =   3
         X1              =   2820
         X2              =   3030
         Y1              =   1680
         Y2              =   2265
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0C0&
         Index           =   4
         X1              =   660
         X2              =   1275
         Y1              =   1470
         Y2              =   1470
      End
   End
   Begin Threed.SSPanel lblPaleteName 
      Height          =   600
      Left            =   6570
      TabIndex        =   72
      Top             =   240
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   1058
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   16776960
      BackColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "TEST NEW CODE WITH MULTIPLE LINES ON THE HEADING"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   20235
      Top             =   2430
   End
   Begin Threed.SSPanel sspVolumeTot 
      Height          =   270
      Left            =   6555
      TabIndex        =   23
      Top             =   900
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   476
      _Version        =   131074
      BackColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      FloodColor      =   16777215
      Begin MSComctlLib.ProgressBar pgLeft 
         Height          =   90
         Left            =   135
         TabIndex        =   24
         Top             =   45
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   159
         _Version        =   393216
         Appearance      =   0
         Max             =   30473
      End
      Begin MSComctlLib.ProgressBar pgRight 
         Height          =   90
         Left            =   135
         TabIndex        =   25
         Top             =   120
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   159
         _Version        =   393216
         Appearance      =   0
         Max             =   30473
      End
      Begin VB.Label lblOn 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   1
         Left            =   60
         TabIndex        =   60
         Top             =   135
         Width           =   60
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808000&
         Height          =   225
         Left            =   0
         Top             =   15
         Width           =   2775
      End
      Begin VB.Label lblPeakML 
         BackColor       =   &H0047FED0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   0
         Left            =   2475
         TabIndex        =   30
         Top             =   60
         Width           =   45
      End
      Begin VB.Label lblPeakMR 
         BackColor       =   &H0046BAFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   1
         Left            =   2550
         TabIndex        =   34
         Top             =   135
         Width           =   30
      End
      Begin VB.Label lblPeakMR 
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   2
         Left            =   2610
         TabIndex        =   33
         Top             =   135
         Width           =   30
      End
      Begin VB.Label lblPeakMR 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   3
         Left            =   2670
         TabIndex        =   32
         Top             =   135
         Width           =   30
      End
      Begin VB.Label lblPeakMR 
         BackColor       =   &H0047FED0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   0
         Left            =   2475
         TabIndex        =   31
         Top             =   135
         Width           =   45
      End
      Begin VB.Label lblPeakML 
         BackColor       =   &H0046BAFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   1
         Left            =   2550
         TabIndex        =   29
         Top             =   60
         Width           =   30
      End
      Begin VB.Label lblOn 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   0
         Left            =   60
         TabIndex        =   28
         Top             =   60
         Width           =   60
      End
      Begin VB.Label lblPeakML 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   3
         Left            =   2670
         LinkTimeout     =   30
         TabIndex        =   27
         Top             =   60
         Width           =   30
      End
      Begin VB.Label lblPeakML 
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   60
         Index           =   2
         Left            =   2610
         TabIndex        =   26
         Top             =   60
         Width           =   30
      End
   End
   Begin VB.ListBox lstSystem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   3510
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   150
      Width           =   2460
   End
   Begin VB.Timer Timer4 
      Interval        =   250
      Left            =   20280
      Top             =   5670
   End
   Begin VB.Timer timFade 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   20280
      Top             =   6120
   End
   Begin Threed.SSPanel SSPanel17 
      Height          =   780
      Left            =   25185
      TabIndex        =   50
      Top             =   615
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   1376
      _Version        =   131074
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SSPanel17"
   End
   Begin VB.Timer tmFade 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   20280
      Top             =   6570
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   24600
      TabIndex        =   0
      Top             =   9285
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   21480
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   165
      Width           =   345
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   20280
      Top             =   7035
   End
   Begin Threed.SSPanel sspStream2 
      Height          =   225
      Left            =   25035
      TabIndex        =   15
      Top             =   5745
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   397
      _Version        =   131074
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ready..."
   End
   Begin Threed.SSPanel sspStream1 
      Height          =   225
      Left            =   25035
      TabIndex        =   14
      Top             =   5475
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   397
      _Version        =   131074
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ready"
   End
   Begin VB.Timer TimerMainLevel 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   20280
      Top             =   5040
   End
   Begin VB.Timer TimerF2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   20280
      Top             =   1995
   End
   Begin VB.Timer TimerF1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   20280
      Top             =   1560
   End
   Begin VB.Timer TimerP2Level 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   20280
      Top             =   4605
   End
   Begin VB.Timer TimerP1Level 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   20280
      Top             =   4200
   End
   Begin VB.Timer TimerP2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   20280
      Top             =   3735
   End
   Begin VB.Timer TimerP1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   20280
      Top             =   3300
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   20280
      Top             =   2880
   End
   Begin Threed.SSPanel cmdReset 
      Height          =   1620
      Left            =   24705
      TabIndex        =   1
      Top             =   2340
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   2858
      _Version        =   131074
      BackColor       =   3333
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   4
   End
   Begin Threed.SSPanel SSPanel8 
      Height          =   825
      Left            =   11805
      TabIndex        =   2
      Top             =   0
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1455
      _Version        =   131074
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Begin Threed.SSPanel lblDate 
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   555
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   12632064
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "26 December 2014"
         BevelOuter      =   0
      End
      Begin Threed.SSPanel sspTime 
         Height          =   540
         Left            =   285
         TabIndex        =   4
         Top             =   -15
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   953
         _Version        =   131074
         ForeColor       =   16776960
         BackColor       =   8421504
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "23:59"
         BevelWidth      =   2
         BevelOuter      =   0
         Alignment       =   1
         Begin Threed.SSPanel sspTimeInd 
            Height          =   375
            Left            =   540
            TabIndex        =   6
            Top             =   75
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   661
            _Version        =   131074
            ForeColor       =   8421376
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ":"
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   180
         Left            =   255
         TabIndex        =   5
         Top             =   1260
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   318
         _Version        =   131074
         ForeColor       =   12632256
         BackColor       =   8421504
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CPU:"
         BevelWidth      =   2
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel sspCpu 
         Height          =   135
         Left            =   630
         TabIndex        =   7
         Top             =   1200
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   238
         _Version        =   131074
         ForeColor       =   65535
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
         FloodType       =   1
         FloodShowPct    =   0   'False
         FloodColor      =   65280
      End
   End
   Begin Threed.SSPanel SSPanel15 
      Height          =   525
      Left            =   25950
      TabIndex        =   41
      Top             =   9675
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   926
      _Version        =   131074
      BackColor       =   14993249
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin Threed.SSPanel sspButtonPlayStop 
         Height          =   300
         Index           =   1
         Left            =   570
         TabIndex        =   42
         Top             =   15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   529
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   14993249
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Play  /  STOP"
         BevelOuter      =   0
      End
      Begin Threed.SSPanel sspButtonPlayStop 
         Height          =   300
         Index           =   0
         Left            =   15
         TabIndex        =   43
         Top             =   15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   529
         _Version        =   131074
         CaptionStyle    =   1
         BackColor       =   14993249
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Play  /  Pause"
         BevelOuter      =   0
      End
   End
   Begin Threed.SSPanel sspVersion 
      Height          =   210
      Index           =   3
      Left            =   21465
      TabIndex        =   52
      Top             =   630
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   370
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Serial:"
      BevelWidth      =   2
      BevelOuter      =   0
      Alignment       =   3
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   300
      Left            =   9615
      TabIndex        =   56
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   529
      _Version        =   131074
      BackColor       =   12632256
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
      BevelOuter      =   0
      Begin Threed.SSPanel lblTotPlayTime 
         Height          =   315
         Left            =   990
         TabIndex        =   57
         Top             =   45
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   131074
         ForeColor       =   16776960
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "00:00:00"
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   330
         Left            =   -45
         TabIndex        =   58
         Top             =   45
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   131074
         ForeColor       =   16777152
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Playlist Time :"
         BevelOuter      =   0
         Alignment       =   4
      End
      Begin Threed.SSPanel lblTotPlayTimeLeft 
         Height          =   315
         Left            =   2940
         TabIndex        =   74
         Top             =   45
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   131074
         ForeColor       =   16776960
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "00:00:00"
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   330
         Left            =   2130
         TabIndex        =   75
         Top             =   45
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   582
         _Version        =   131074
         ForeColor       =   16777152
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Time left :"
         BevelOuter      =   0
         Alignment       =   4
      End
   End
   Begin Threed.SSPanel SSPanel16 
      Height          =   945
      Left            =   13740
      TabIndex        =   44
      Top             =   180
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   1667
      _Version        =   131074
      BackColor       =   3092271
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin Threed.SSCommand cmdFadeOut 
         Height          =   825
         Left            =   990
         TabIndex        =   54
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1455
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   15194953
         BackColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPlayer.frx":39B3
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdPause 
         Height          =   825
         Left            =   75
         TabIndex        =   53
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1455
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   14149394
         BackColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPlayer.frx":3D65
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdSettings 
         Height          =   825
         Left            =   4650
         TabIndex        =   49
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1455
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   15194953
         BackColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPlayer.frx":40F5
         Alignment       =   8
         ButtonStyle     =   3
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdLoadPalette 
         Height          =   825
         Left            =   1950
         TabIndex        =   48
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1455
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   15194953
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Load Playlist"
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdSavePalette 
         Height          =   825
         Left            =   2865
         TabIndex        =   47
         Top             =   75
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1455
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   15194953
         BackColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Save Playlist"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdClearPalette 
         Height          =   825
         Left            =   3750
         TabIndex        =   46
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1455
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   15194953
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Clear Playlist"
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   825
         Left            =   5550
         TabIndex        =   45
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1455
         _Version        =   131074
         ForeColor       =   15194953
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "EXIT"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
   End
   Begin Threed.SSPanel PanMain 
      Height          =   9720
      Left            =   165
      TabIndex        =   8
      Top             =   1320
      Width           =   15960
      _ExtentX        =   28152
      _ExtentY        =   17145
      _Version        =   131074
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SSPanel1"
      ClipControls    =   0   'False
      BevelOuter      =   0
      Begin Threed.SSPanel cmdSong 
         Height          =   6435
         Index           =   0
         Left            =   75
         Negotiate       =   -1  'True
         TabIndex        =   9
         Top             =   15
         Visible         =   0   'False
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   11351
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   128
         PictureMaskColor=   65535
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1"
         ClipControls    =   0   'False
         BorderWidth     =   5
         Alignment       =   0
         PictureAlignment=   7
         FloodColor      =   49344
         Begin Threed.SSPanel sspSongTitle 
            Height          =   1035
            Index           =   0
            Left            =   225
            TabIndex        =   62
            Top             =   30
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   1826
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   0
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "SSPanel2"
            Begin VB.Image imgDirection 
               Height          =   240
               Index           =   0
               Left            =   3015
               Stretch         =   -1  'True
               Top             =   -15
               Width           =   240
            End
            Begin VB.Label lblButtonCnt 
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   345
               TabIndex        =   85
               Top             =   -15
               Width           =   285
            End
            Begin VB.Image imgCompleted3 
               Height          =   240
               Index           =   0
               Left            =   2880
               Picture         =   "frmPlayer.frx":45FD
               Top             =   645
               Width           =   240
            End
            Begin VB.Image imgCompleted2 
               Height          =   240
               Index           =   0
               Left            =   15
               Picture         =   "frmPlayer.frx":4B87
               Top             =   630
               Width           =   240
            End
            Begin VB.Image imgCompleted0 
               Height          =   240
               Index           =   0
               Left            =   30
               Picture         =   "frmPlayer.frx":5111
               Top             =   15
               Width           =   240
            End
            Begin VB.Image imgCompleted1 
               Height          =   240
               Index           =   0
               Left            =   2880
               Picture         =   "frmPlayer.frx":569B
               Top             =   15
               Width           =   240
            End
         End
         Begin Threed.SSPanel sspProgress11 
            Height          =   270
            Index           =   0
            Left            =   1845
            TabIndex        =   10
            Top             =   4665
            Visible         =   0   'False
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   476
            _Version        =   131074
            ForeColor       =   16777215
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin MSComctlLib.ProgressBar sspProgress 
            Height          =   215
            Index           =   0
            Left            =   1020
            TabIndex        =   40
            Top             =   1170
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   370
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Image imgSetup 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   3675
            Picture         =   "frmPlayer.frx":5C25
            Stretch         =   -1  'True
            Top             =   855
            Width           =   240
         End
         Begin VB.Label lblStream 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Stream  : 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   150
            Index           =   0
            Left            =   1125
            TabIndex        =   51
            Top             =   1440
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label lblSelect 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   0
            Left            =   1995
            TabIndex        =   39
            Top             =   2190
            Width           =   1140
         End
         Begin VB.Image imgVol 
            Height          =   450
            Index           =   0
            Left            =   210
            Stretch         =   -1  'True
            Top             =   1170
            Width           =   450
         End
         Begin VB.Label lblVol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   16
            Top             =   2655
            Width           =   270
         End
         Begin VB.Label lblTimeLeft 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   300
            Index           =   0
            Left            =   675
            TabIndex        =   11
            Top             =   1380
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblTimePlayed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   0
            Left            =   3210
            TabIndex        =   13
            Top             =   1095
            Visible         =   0   'False
            Width           =   315
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
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Index           =   0
            Left            =   3465
            TabIndex        =   12
            Top             =   1740
            Visible         =   0   'False
            Width           =   690
         End
      End
      Begin Threed.SSPanel sspDevice 
         Height          =   600
         Left            =   12585
         TabIndex        =   76
         Top             =   9045
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1058
         _Version        =   131074
         ForeColor       =   16776960
         BackColor       =   8421504
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sound card..."
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel sspSndHead 
         Height          =   600
         Left            =   11445
         TabIndex        =   77
         Top             =   9060
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1058
         _Version        =   131074
         ForeColor       =   16776960
         BackColor       =   8421504
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sound card :"
         BorderWidth     =   1
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel sspPageMain 
         Height          =   600
         Left            =   255
         TabIndex        =   78
         Top             =   9045
         Width           =   10710
         _ExtentX        =   18891
         _ExtentY        =   1058
         _Version        =   131074
         BackColor       =   3092271
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin Threed.SSCommand cmdPage 
            Height          =   480
            Index           =   6
            Left            =   8925
            TabIndex        =   84
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   847
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Page 6"
            ButtonStyle     =   3
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdPage 
            Height          =   480
            Index           =   5
            Left            =   7155
            TabIndex        =   83
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   847
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Page 5"
            ButtonStyle     =   3
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdPage 
            Height          =   480
            Index           =   4
            Left            =   5385
            TabIndex        =   82
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   847
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Page 4"
            ButtonStyle     =   3
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdPage 
            Height          =   480
            Index           =   3
            Left            =   3615
            TabIndex        =   81
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   847
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Page 3"
            ButtonStyle     =   3
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdPage 
            Height          =   480
            Index           =   2
            Left            =   1845
            TabIndex        =   80
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   847
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Page 2"
            ButtonStyle     =   3
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdPage 
            Height          =   480
            Index           =   1
            Left            =   75
            TabIndex        =   79
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   847
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Page 1"
            ButtonStyle     =   3
            BevelWidth      =   1
            RoundedCorners  =   0   'False
            Outline         =   0   'False
         End
      End
   End
   Begin Threed.SSPanel SSPanel9 
      Height          =   525
      Left            =   15
      TabIndex        =   61
      Top             =   195
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   926
      _Version        =   131074
      MarqueeStyle    =   3
      ForeColor       =   16776960
      MarqueeDelay    =   250
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DEMO MODE"
      BevelOuter      =   0
   End
   Begin Threed.SSPanel sspSecure 
      Height          =   1770
      Left            =   21480
      TabIndex        =   66
      Top             =   7320
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   3122
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   3333
      MarqueeDelay    =   300
      MarqueeScrollAmount=   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.CommandButton cmdCancelSecure 
         BackColor       =   &H00E4C761&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   4575
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   975
         Width           =   1620
      End
      Begin VB.CommandButton cmdOKSecure 
         BackColor       =   &H00E4C761&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   4575
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   315
         Width           =   1620
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         BackColor       =   &H00E4C761&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "l"
         TabIndex        =   68
         Top             =   780
         Width           =   3510
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E4C761&
         Height          =   285
         Left            =   375
         TabIndex        =   67
         Top             =   300
         Width           =   3315
      End
   End
   Begin Threed.SSPanel sspLoading 
      Height          =   1065
      Left            =   18735
      TabIndex        =   64
      Top             =   5100
      Visible         =   0   'False
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   1879
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   3333
      MarqueeDelay    =   300
      MarqueeScrollAmount=   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Loading Available Songs..."
      BevelOuter      =   1
      Begin Threed.SSPanel sspFlood 
         Height          =   150
         Left            =   465
         TabIndex        =   65
         Top             =   825
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   265
         _Version        =   131074
         MarqueeDirection=   1
         ForeColor       =   16777215
         BackColor       =   3333
         MarqueeDelay    =   5
         MarqueeScrollAmount=   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "l"
         BevelOuter      =   0
         FloodType       =   1
         FloodShowPct    =   0   'False
         Alignment       =   1
         FloodColor      =   65535
      End
   End
   Begin VB.Label lblCompiled 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled on 12121212"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   255
      TabIndex        =   71
      Top             =   960
      Width           =   2940
   End
   Begin VB.Label lblCursorPlaceHolder 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   21060
      TabIndex        =   63
      Top             =   1485
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgCompletedLarge 
      Height          =   360
      Left            =   23835
      Picture         =   "frmPlayer.frx":6067
      Top             =   4440
      Width           =   360
   End
   Begin VB.Image imgCompletedSmall 
      Height          =   240
      Left            =   23925
      Picture         =   "frmPlayer.frx":70E9
      Top             =   4035
      Width           =   240
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   255
      TabIndex        =   55
      Top             =   750
      Width           =   2940
   End
   Begin VB.Image imgDirectionSource 
      Height          =   210
      Index           =   1
      Left            =   23145
      Picture         =   "frmPlayer.frx":7673
      Stretch         =   -1  'True
      Top             =   5175
      Width           =   210
   End
   Begin VB.Image imgDirectionSource 
      Height          =   210
      Index           =   0
      Left            =   23160
      Picture         =   "frmPlayer.frx":8075
      Stretch         =   -1  'True
      Top             =   4845
      Width           =   210
   End
   Begin VB.Menu mnuPopUps 
      Caption         =   "PopMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPop 
         Caption         =   "Play Song"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Assign Song"
         Index           =   2
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Clear button"
         Index           =   3
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Clear All Buttons"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gPalette  As String
Dim iMnuFlag As Integer   '0 = Play, 1=Assign, 2 = Clear
Dim i As Integer
Dim iCntPlayers As Integer
Dim Player1Index As Integer
Dim Player2Index As Integer
Dim bPlayer1Playing As Boolean
Dim bPlayer2Playing As Boolean
Dim bAsigning As Boolean
Const PageButHeight As Integer = 600

Dim iVol1 As Single
Dim iVol2 As Single
Const PeakNormal As Single = 28835   '0.88   '32767 * 0.88
Const PeakMidLow As Single = 30473   '0.93
Const PeakMidHigh As Single = 31784  '0.97
Const PeakMax As Long = 32767
Const PeakDispMax As Long = 28834.96
Dim iBarColor As Integer
Dim bLoading As Boolean
Dim DragingButton As Boolean
Dim PlayingCurrently As Integer
Dim iFlood As Integer
Dim ButColors(0 To 16) As Long
Dim ButForColors(0 To 16) As Long
Dim buflen As Long
Dim sTotalPlaytime As String
Dim iVolIndex As Integer
Dim SpacebarPressed As Boolean
Dim LastPlaying As Integer

Dim fxBass(2) As Long        ' 3 eq band + reverb
Dim fxMid(2) As Long
Dim fxHigh(2) As Long

Const iBassFreq As Integer = 83
Const iMidFreq As Integer = 1000
Const iHighFreq As Integer = 8000

Dim RightMost As Long

Dim p As BASS_DX8_PARAMEQ

Dim bSecureClose As Boolean

Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean

''''Public Sub DrawGradient(ByRef dstObject As Control, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Direction As Long, Peak As Long)
'''''Public Sub DrawGradient(ByRef dstObject As Form, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Direction As Long)
''''
''''    'R, G, B variables for each end of the gradient
''''    Dim R As Long, G As Long, B As Long
''''    Dim R2 As Long, G2 As Long, B2 As Long
''''
''''    'Fill our RGB values from the longs supplied by the calling routine
''''    R = Color1 Mod 256
''''    G = (Color1 \ 256) And 255
''''    B = (Color1 \ 65536) And 255
''''    R2 = Color2 Mod 256
''''    G2 = (Color2 \ 256) And 255
''''    B2 = (Color2 \ 65536) And 255
''''
''''    'Always use variables for storing object values - it's loads faster
''''    Dim TempWidth As Long, TempHeight As Long
''''    TempWidth = dstObject.ScaleWidth
''''    TempHeight = dstObject.ScaleHeight
''''
''''    'Several calculation variables for generating the gradient
''''    Dim VR As Single, VG As Single, VB As Single
''''
''''    'Vertical gradient
''''    If Direction = 1 Then
''''
''''        'First, create a calculation variable for determining the step
''''        'between each level of the gradient (large if the destination form
''''        'is small, small if the destination form is large); for example, this
''''        'value will be exactly 1 for each variable if the control is 255 pixels
''''        'tall and the gradient is going from pure black to pure white
''''        VR = Abs(R - R2) / TempHeight
''''        VG = Abs(G - G2) / TempHeight
''''        VB = Abs(B - B2) / TempHeight
''''
''''        'If the second value is lower then the first value, make the step
''''        'negative (so that we subtract as we go along, not add)
''''        If R2 < R Then VR = -VR
''''        If G2 < G Then VG = -VG
''''        If B2 < B Then VB = -VB
''''
''''        'Lastly, run a loop through the height of the control, incrementing (or if
''''        'negative, decrementing) the gradient color according to the y-coordinate
''''        'of the current line of the control
''''        For Y = 0 To TempHeight
''''            R2 = R + VR * Y
''''            G2 = G + VG * Y
''''            B2 = B + VB * Y
''''            dstObject.Line (0, Y)-(TempWidth, Y), RGB(R2, G2, B2)
''''        Next Y
''''
''''    'Horizontal gradients work exactly the same, except that they (obviously)
''''    'run from left-to-right instead of up-and-down
''''    Else
''''
''''        VR = Abs(R - R2) / TempWidth
''''        VG = Abs(G - G2) / TempWidth
''''        VB = Abs(B - B2) / TempWidth
''''
''''        If R2 < R Then VR = -VR
''''        If G2 < G Then VG = -VG
''''        If B2 < B Then VB = -VB
''''
''''        For X = 0 To TempWidth
''''            R2 = R + VR * X
''''            G2 = G + VG * X
''''            B2 = B + VB * X
''''            dstObject.Line (X, 0)-(X, TempHeight), RGB(R2, G2, B2)
''''        Next X
''''
''''    End If
''''
''''End Sub


Private Sub chkAutoAdvance_Click()

End Sub

Private Sub cmdCancelSecure_Click()

txtPassword.text = ""

bSecureClose = True
sspSecure.Left = 20000

End Sub

Private Sub cmdClearPalette_Click()
   On Error Resume Next
   ClearButtons
   DoEvents
   If Not TimerMainLevel.Enabled Then
      InitialisePeaks 0
   End If
   PaletteName = "TMP001"
   Me.Caption = "TMP001"
   
   ClearPalet = True
   SavePalete Trim(Me.Caption), 1
   LoadPalette Trim(Me.Caption), 1, 5
   SetPageButton 1
   ClearPalet = False
   
   lblPaleteName.Caption = Me.Caption
   lblTotPlayTime.Caption = "0:00"
   lblTotPlayTimeLeft.Caption = "0:00"
   Command2.SetFocus

End Sub

'''''''Sub ResetProgress(Index As Integer)
'''''''
'''''''  SetProgress Index, 1
'''''''
'''''''End Sub
'''''''
'''''''Sub SetProgress(Index As Integer, XPercent As Long)
'''''''
'''''''   Dim lIncr As Single
'''''''   lIncr = (RightMost - lnProgress(Index).X1) / 100
'''''''   lnProgress(Index).X2 = lnProgress(Index).X1 + (lIncr * XPercent)
'''''''
'''''''
'''''''End Sub

Sub ClearButton(Index As Integer)

   On Error Resume Next
   
   'Closevolume
   
   imgVol(Index).Visible = False
   
   lblStream(Index).Caption = ""
   lblStream(Index).Tag = ""
   lblStream(Index).Visible = False
   
   
   sspSongTitle(Index).Caption = ""
   sspSongTitle(Index).Tag = ""
   sspSongTitle(Index).TagVariant = ""
   cmdSong(Index).TagVariant = ""
   sspProgress(Index).Tag = ""
   sspProgress(Index).ToolTipText = ""
   
   sspProgress(Index).max = 100
   cmdSong(Index).BackStyle = ssOpaque
   cmdSong(Index).BackColor = cmdReset.BackColor
   
   lblTimePlayed(Index).Caption = ""
   lblTimePlayed(Index).Tag = ""
   lblTimeLeft(Index).Caption = ""
   lblStatus(Index).Caption = "Ready"
   lblStatus(Index).ForeColor = vbWhite
   
   cmdSong(Index).TagVariant = "" 'Use this to keep the Average when song is loaded...
   
   sspSongTitle(Index).Visible = False
   lblTimePlayed(Index).Visible = False
   lblTimeLeft(Index).Visible = False
   lblStatus(Index).Visible = False
   sspProgress(Index).Visible = False
   imgDirection(Index).Visible = False
   imgCompleted0(Index).Visible = False
   imgCompleted1(Index).Visible = False
   imgCompleted2(Index).Visible = False
   imgCompleted3(Index).Visible = False
   'imgSetup(Index).Visible = False
   
   imgVol(Index).Visible = False
  ' If bDoEq Then imgEQ(Index).Visible = False
     
   cmdSong(Index).Picture = Nothing
   cmdSong(Index).BackColor = cmdReset.BackColor   'vbDefaultBack
   
   lblVol(Index).Caption = "0"
         
   cmdSong(Index).Caption = " " & GetNextButtonNumber + Index
   lblButtonCnt(Index).Caption = cmdSong(Index).Caption
   
   Closevolume True
   
'   GetTotalTime
   
   
   Exit Sub
   
err1:
   MsgBox Err.Description
       
End Sub

Sub SetButtonsLayout()
   Dim iBut As Integer
   Dim iDispBut As Integer 'The numnber that is going to be displayed on the button, ie 1-16 or 17 to 32 etc
   Dim iTop As Long
   Dim iLeft As Long
   Dim iMaxRows As Integer
   Dim iMaxCols As Integer
   Dim iCol As Integer
   Dim iRow As Integer
   Dim sRow As String
   Dim sCol As String
   Dim bMoveNextBut As Boolean
   Dim iLblWidth As Integer
   Dim AdjustButtonHeight As Integer
   Dim AdjustButtonWidth As Integer
   Dim AdjustHSpacing As Integer
   Dim AdjustVSpacing As Integer
   Dim ButtonHeight As Integer
   Dim ButtonWidth As Integer
   
   Dim PlayAreaTop As Integer
   Dim PlayAreaLeft As Integer
   Dim PlayAreaHeight As Integer
   Dim PlayAreaWidth As Integer
   
   Dim ProgressTop As Integer
   Dim ProgressLeft As Integer
   Dim ProgressWidth As Integer
   Dim ProgressHeight As Integer
   
   Dim TimePlayedLeft As Integer
   Dim TimePlayedTop As Integer
   Dim TimePlayedWidth As Integer
   Dim TimePlayedHeight As Integer
   
   Dim TimeRemainLeft As Integer
   Dim TimeRemainTop As Integer
   Dim TimeRemainHeight As Integer
   Dim TimeRemainWidth As Integer
   
   Dim VolumeTop As Integer
   Dim VolumeLeft As Integer
   Dim VolumeWidth As Integer
   Dim PlayFontSize As Integer
   Dim TimeFontSize As Integer

   
 '  iMaxBut = 20
   
   On Error GoTo err1
   
   iTop = 15
   iLeft = 75  '100  '120
   iCol = 0
   iRow = 0
   
   
   
   
   
   bMoveNextBut = False
   
   'Set the Adjustments values for butto sizes and positioning...
'   Select Case iMaxBut
'      Case 9
'         AdjustHSpacing = 10
'         AdjustVSpacing = 90
'         AdjustButtonWidth = 0
'         AdjustButtonHeight = 60
'         ButtonHeight = 3000
'         ButtonWidth = 6660
'      Case 16
'         AdjustHSpacing = 15
'         AdjustVSpacing = 90
'         AdjustButtonWidth = 0
'         AdjustButtonHeight = 60
'         ButtonHeight = 2235 - (PageButHeight / 5)
'         ButtonWidth = 4980
'      Case 25
'         AdjustHSpacing = 10
'         AdjustVSpacing = 80
'         AdjustButtonWidth = 0
'         AdjustButtonHeight = 60
'         ButtonHeight = 1775
'         ButtonWidth = 3972
'      Case 30
'         AdjustHSpacing = 5
'         AdjustVSpacing = 80
'         AdjustButtonWidth = 0
'         AdjustButtonHeight = 60
'         ButtonHeight = 1470
'         ButtonWidth = 3972
'   End Select
      

   
   If ApplyStandardTheme Then
      'Set the Adjustments values for butto sizes and positioning...
      Select Case iMaxBut
         Case 9
            AdjustHSpacing = 10
            AdjustVSpacing = 90
            AdjustButtonWidth = 0
            AdjustButtonHeight = 60
            ButtonHeight = 3000
            ButtonWidth = 6660
         Case 16
            AdjustHSpacing = 15
            AdjustVSpacing = 90
            AdjustButtonWidth = 0
            AdjustButtonHeight = 60
            ButtonHeight = 2235 - (PageButHeight / 5)
            ButtonWidth = 4980
         Case 25
            AdjustHSpacing = 10
            AdjustVSpacing = 80
            AdjustButtonWidth = 0
            AdjustButtonHeight = 60
            ButtonHeight = 1775
            ButtonWidth = 3972
         Case 30
            AdjustHSpacing = 5
            AdjustVSpacing = 80
            AdjustButtonWidth = 0
            AdjustButtonHeight = 60
            ButtonHeight = 1470
            ButtonWidth = 3972
      End Select
   Else
      'Set the Adjustments values for butto sizes and positioning...
      Select Case iMaxBut
         Case 9
            AdjustHSpacing = 75 'Top to Bottom
            AdjustVSpacing = 75 'Left to Right
            AdjustButtonHeight = -30
            AdjustButtonWidth = 0
            ButtonHeight = 3000
            ButtonWidth = 6660
         Case 16
            AdjustHSpacing = 75 'Top to Bottom
            AdjustVSpacing = 60 'Left to Right
            AdjustButtonHeight = -30
            AdjustButtonWidth = 0
            ButtonHeight = 2235 - (PageButHeight / 4)
            ButtonWidth = 4980
            
            iMaxRows = 4
            iMaxCols = 4
            lblStream(0).Font.Size = 7
            lblStream(0).Font.Bold = True
            
            PlayAreaTop = 30
            PlayAreaLeft = 30
            PlayAreaHeight = 1525
            'PlayAreaWidth = 4915
            PlayAreaWidth = ButtonWidth + AdjustButtonWidth - 60 '3930   '4000
                                   
            TimeRemainLeft = 660  '115
            TimeRemainTop = 1705
            TimeRemainWidth = 600  '1100
            TimeRemainHeight = 230
            
            ProgressLeft = 1295
            ProgressTop = 1720
            ProgressWidth = 2455 + 450
            ProgressHeight = 215
            
            TimePlayedLeft = 4260  '4305  '3905 + 400
            TimePlayedTop = 1705
            TimePlayedWidth = 600  '1100
            TimePlayedHeight = 230
            
            PlayFontSize = 12
            TimeFontSize = 10
         
            MaxWidth = 2900

         Case 20
            AdjustHSpacing = 75 'Top to Bottom
            AdjustVSpacing = 60 'Left to Right
            AdjustButtonHeight = -30
            AdjustButtonWidth = 0
            ButtonHeight = 2235 - (PageButHeight / 4)
            ButtonWidth = 4000
            
            iMaxRows = 4
            iMaxCols = 5
            
            lblStream(0).Font.Size = 7
            lblStream(0).Font.Bold = True
            
            PlayAreaTop = 30
            PlayAreaLeft = 30
            PlayAreaHeight = 1525
            PlayAreaWidth = ButtonWidth + AdjustButtonWidth - 60 '3930   '4000
                                   
            TimeRemainLeft = 600  '115
            TimeRemainTop = 1705
            TimeRemainWidth = 460  '1100
            TimeRemainHeight = 230
            
            ProgressLeft = 1120
            ProgressTop = 1720
            ProgressWidth = 2200  '2455 + 450
            ProgressHeight = 215
            
            TimePlayedLeft = 3375  '4260
            TimePlayedTop = 1705
            TimePlayedWidth = 500  '1100
            TimePlayedHeight = 230
            PlayFontSize = 11
            TimeFontSize = 9
         
            MaxWidth = 2900

         Case 30
            AdjustHSpacing = 45 'Top to Bottom
            AdjustVSpacing = 45 'Left to Right
            AdjustButtonHeight = 0
            AdjustButtonWidth = 15
            ButtonHeight = 1660
            ButtonWidth = 3315
            iMaxRows = 5
            iMaxCols = 6
            
            lblStream(0).Font.Size = 7
            lblStream(0).Font.Bold = True
            
            PlayAreaTop = 30
            PlayAreaLeft = 30
            PlayAreaHeight = 1130
            PlayAreaWidth = ButtonWidth + AdjustButtonWidth - 60 '3930   '4000
                                   
            TimeRemainLeft = 545  '115
            TimeRemainTop = 1360
            TimeRemainWidth = 460  '1100
            TimeRemainHeight = 230
            
            ProgressLeft = 1060
            ProgressTop = 1345
            ProgressWidth = 1660  '2455 + 450
            ProgressHeight = 215
            
            TimePlayedLeft = ProgressLeft + ProgressWidth + 60  '4260
            TimePlayedTop = 1360
            TimePlayedWidth = 500  '1100
            TimePlayedHeight = 230
            
            PlayFontSize = 10
            TimeFontSize = 7
         
            MaxWidth = 2900
      End Select
   End If
      
   LoadDataIntoFile 114, App.Path & "\tmpClear" 'Clear
   LoadDataIntoFile 115, App.Path & "\tmpOpen"  'Open
   LoadDataIntoFile 116, App.Path & "\tmpColor" 'Color
   LoadDataIntoFile 132, App.Path & "\tmpDrive" 'Drive
   LoadDataIntoFile 124, App.Path & "\tmpDrive" 'Folder with subfolder
   LoadDataIntoFile 127, App.Path & "\tmpDrive" 'Empty Folder
   LoadDataIntoFile 129, App.Path & "\tmpDrive" 'Folder with Audio in root
   LoadDataIntoFile 131, App.Path & "\tmpDrive" 'Audio files
   LoadDataIntoFile 123, App.Path & "\tmpDrive" 'Video Files
   LoadDataIntoFile 134, App.Path & "\tmpClose" 'Close
   LoadDataIntoFile 135, App.Path & "\tmpSpkr"  'Close
        
   '*****************************************
   'Set the initial button size and layout...
   '*****************************************
   cmdSong(0).Top = 15
   cmdSong(0).Height = ButtonHeight + AdjustButtonHeight
   cmdSong(0).Width = ButtonWidth + AdjustButtonWidth
         
   '*********************************
   'Play AREA
   '*********************************
   sspSongTitle(0).Top = PlayAreaTop
   sspSongTitle(0).Left = PlayAreaLeft
   sspSongTitle(0).Height = PlayAreaHeight
   sspSongTitle(0).Width = PlayAreaWidth
   sspSongTitle(0).BackStyle = ssOpaque
   sspSongTitle(0).BevelInner = ssNoneBevel
   sspSongTitle(0).BevelOuter = ssNoneBevel
   sspSongTitle(0).Font.Size = PlayFontSize
   
   'Button Count label
   lblButtonCnt(0).Caption = "1"
   lblButtonCnt(0).Left = -15
   lblButtonCnt(0).Top = -15
   lblButtonCnt(0).BackStyle = 0
   lblButtonCnt(0).ForeColor = vbWhite
   lblButtonCnt(0).Font.Size = 7
   lblButtonCnt(0).Font.Bold = True
   'Direction Image
   If iButtonDirection = 1 Then 'Top To Bottom
      imgDirection(0).Picture = imgDirectionSource(1).Picture
      'imgDirection(0).Left = sspSongTitle(0).Width - imgDirection(0).Width - 30
   Else
      imgDirection(0).Picture = imgDirectionSource(0).Picture
      'imgDirection(0).Left = sspSongTitle(0).Width - imgDirection(0).Width - 30
   End If
   imgDirection(0).Top = 15
   imgDirection(0).Left = sspSongTitle(0).Width - imgDirection(0).Width - 75
      
   '*********************************
   'Labels for Playing Time/progress
   '*********************************
   lblTimePlayed(0).Left = TimePlayedLeft '145
   lblTimePlayed(0).Top = TimePlayedTop '180
   lblTimePlayed(0).Height = TimePlayedHeight
   lblTimePlayed(0).Width = TimePlayedWidth
   lblTimePlayed(0).Caption = "00:00"
   lblTimePlayed(0).FontSize = TimeFontSize
   lblTimePlayed(0).FontBold = False
   lblTimePlayed(0).Alignment = 0   'Left
   
   lblTimePlayed(0).BorderStyle = 0
   lblTimeLeft(0).BorderStyle = 0
   
   lblTimeLeft(0).Left = TimeRemainLeft
   lblTimeLeft(0).Top = TimeRemainTop
   lblTimeLeft(0).Height = TimeRemainHeight
   lblTimeLeft(0).Width = TimeRemainWidth
   lblTimeLeft(0).AutoSize = False
   lblTimeLeft(0).Caption = "00:00"
   lblTimeLeft(0).FontSize = lblTimePlayed(0).FontSize
   lblTimeLeft(0).FontBold = False
   lblTimeLeft(0).ForeColor = &HFFC0C0
   lblTimeLeft(0).Alignment = 1     'Right

   
   imgVol(0).Left = 100  '150
   imgVol(0).Top = ProgressTop - 160
   
   sspProgress(0).Top = ProgressTop
   sspProgress(0).Height = ProgressHeight
   sspProgress(0).Width = ProgressWidth
   sspProgress(0).Left = ProgressLeft
   Change_pb_ForeColor sspProgress(0).hWnd, vbGreen    '&HFFAE27
   Change_pb_Color sspProgress(0).hWnd, &H260F35   'Default back color

   lblSelect(0).Left = 0
   lblSelect(0).BackStyle = 0
   
   imgVol(0).Height = 450
   imgVol(0).Width = 450
   
   lblSelect(0).Top = lblTimeLeft(0).Top - 150
   lblSelect(0).Width = lblTimeLeft(0).Width
   
   imgCompleted0(0).Picture = imgCompletedLarge.Picture
   imgCompleted1(0).Picture = imgCompletedLarge.Picture
   imgCompleted2(0).Picture = imgCompletedLarge.Picture
   imgCompleted3(0).Picture = imgCompletedLarge.Picture

   
''''   lblPeakL(0).Left = 60
''''   lblPeakR(0).Left = 210
''''   lblMidHL(0).Left = lblPeakL(0).Left
''''   lblMidLL(0).Left = lblPeakL(0).Left
''''   lblMidHR(0).Left = lblPeakR(0).Left
''''   lblMidLR(0).Left = lblPeakR(0).Left
''''   picLevelL(0).Left = lblPeakL(0).Left - 15  '45
''''   picLevelR(0).Left = lblPeakR(0).Left - 15   '195
   

      
   'lblStream(0).Top = cmdSong(0).Height - lblStream(0).Height   'sspProgress(0).Top + sspProgress(0).Height
   lblStream(0).Top = sspProgress(0).Top + sspProgress(0).Height + 30
   lblStream(0).Left = sspProgress(0).Left   '30
   lblStream(0).Width = sspProgress(0).Width
      
      
''''      lnProgress(0).BorderColor = vbMagenta
   
   '======================================================================
   'General settings of original button - After Initial resizing above ...
   '======================================================================
'''''   lblPeakR(0).Top = lblPeakL(0).Top
'''''   lblMidHL(0).Top = lblPeakL(0).Top + 105
'''''   lblMidHR(0).Top = lblMidHL(0).Top
'''''   lblMidLL(0).Top = lblMidHL(0).Top + 105
'''''   lblMidLR(0).Top = lblMidLL(0).Top
'''''   picLevelL(0).Top = lblMidLL(0).Top + lblMidLL(0).Height
'''''   picLevelR(0).Top = picLevelL(0).Top
'''''   picLevelR(0).Height = picLevelL(0).Height
      
   MaxWidth = sspProgress(0).Width '- 50 'Progress control value
   
   '*********************************
   'IMAGE when Completed song
   '*********************************
   imgCompleted0(0).Top = 200 '30    '(sspSongTitle(0).Height / 2) - (imgCompleted(0).Height / 2)
   imgCompleted0(0).Left = 30   'sspSongTitle(0).Width - imgCompleted0(0).Width - 200
   imgCompleted0(0).Visible = False
   
   imgCompleted1(0).Top = imgCompleted0(0).Top
   imgCompleted1(0).Left = sspSongTitle(0).Width - imgCompleted0(0).Width - 200
   imgCompleted1(0).Visible = False
   
   imgCompleted2(0).Top = (sspSongTitle(0).Height) - (imgCompleted2(0).Height)
   imgCompleted2(0).Left = 30  'sspSongTitle(0).Width - imgCompleted0(0).Width - 200
   imgCompleted2(0).Visible = False
   
   imgCompleted3(0).Top = imgCompleted2(0).Top
   imgCompleted3(0).Left = imgCompleted1(0).Left
   imgCompleted3(0).Visible = False
   
'   iLblWidth = lblTimeLeft(0).Width
'   lblTimeLeft(0).AutoSize = False
   lblTimeLeft(0).Caption = ""
'   lblTimeLeft(0).Width = iLblWidth
   
     
''   If iButtonDirection = 1 Then 'Top To Bottom
''      imgDirection(0).Picture = imgDirectionSource(1).Picture
''      imgDirection(0).Left = sspSongTitle(0).Width - imgDirection(0).Width - 30  '370   '  sspLevel(0).Left + 100
''   Else
''      imgDirection(0).Picture = imgDirectionSource(0).Picture
''      imgDirection(0).Left = sspSongTitle(0).Width - imgDirection(0).Width - 30   '370    'sspLevel(0).Left + 100
''   End If
''
''   imgDirection(0).Top = 15  '60

   
   
   
   
   
   
   '***************************************
   'L O A D   A L L   N E W   B U T T O N S
   '***************************************
   For iBut = 1 To iMaxBut
      '==================================
      'Create the controls dinamicaly
      Load cmdSong(iBut)
      cmdSong(iBut).Left = 25000
      cmdSong(iBut).Visible = True

      
      Load sspSongTitle(iBut)
      Set sspSongTitle(iBut).Container = cmdSong(iBut)
      sspSongTitle(iBut).Visible = True
      
      Load imgDirection(iBut)
      Set imgDirection(iBut).Container = sspSongTitle(iBut)
      'imgDirection(iBut).Visible = True
           
      Load lblButtonCnt(iBut)
      Set lblButtonCnt(iBut).Container = sspSongTitle(iBut)
      lblButtonCnt(iBut).Visible = True
      
      Load imgCompleted0(iBut)
      'Set imgCompleted(iBut).Container = cmdSong(iBut)
      Set imgCompleted0(iBut).Container = sspSongTitle(iBut)
      imgCompleted0(iBut).Visible = False
      
      Load imgCompleted1(iBut)
      Set imgCompleted1(iBut).Container = sspSongTitle(iBut)
      imgCompleted1(iBut).Visible = False
      
      Load imgCompleted2(iBut)
      Set imgCompleted2(iBut).Container = sspSongTitle(iBut)
      imgCompleted2(iBut).Visible = False
      
      Load imgCompleted3(iBut)
      Set imgCompleted3(iBut).Container = sspSongTitle(iBut)
      imgCompleted3(iBut).Visible = False
      
      
      Load lblTimePlayed(iBut)
      Set lblTimePlayed(iBut).Container = cmdSong(iBut)
      lblTimePlayed(iBut).Visible = True
            
      Load lblTimeLeft(iBut)
      Set lblTimeLeft(iBut).Container = cmdSong(iBut)
      lblTimeLeft(iBut).Visible = True
      
      Load lblSelect(iBut)
      Set lblSelect(iBut).Container = cmdSong(iBut)
      lblSelect(iBut).Visible = True
            
      Load lblStatus(iBut)
      Set lblStatus(iBut).Container = cmdSong(iBut)
      lblStatus(iBut).Visible = False
   
      Load imgVol(iBut)
      Set imgVol(iBut).Container = cmdSong(iBut)
      imgVol(iBut).Visible = False
      
'      If bDoEq Then
'        Load imgEQ(iBut)
'        Set imgEQ(iBut).Container = cmdSong(iBut)
'        imgEQ(iBut).Visible = False
'      End If
      
      Load sspProgress(iBut)
      Set sspProgress(iBut).Container = cmdSong(iBut)
      sspProgress(iBut).Visible = True
      
      Load lblStream(iBut)
      Set lblStream(iBut).Container = cmdSong(iBut)
      lblStream(iBut).Visible = False
      
'''      Load lnProgress(iBut)
'''      Set lnProgress(iBut).Container = cmdSong(iBut)
'''      lnProgress(iBut).Visible = True
      
'''      Load sspLevel(iBut)
'''      Set sspLevel(iBut).Container = cmdSong(iBut)
'''      sspLevel(iBut).Visible = False
      
''''      Load picLevelL(iBut)
''''      Set picLevelL(iBut).Container = sspLevel(iBut)
''''      picLevelL(iBut).Visible = True
''''
''''      Load picLevelR(iBut)
''''      Set picLevelR(iBut).Container = sspLevel(iBut)
''''      picLevelR(iBut).Visible = True
            
'''      Load cpvVolume(iBut)
'''      Set cpvVolume(iBut).Container = cmdSong(iBut)
'''      cpvVolume(iBut).Visible = True
      
''''      Load lblPeakL(iBut)
''''      Set lblPeakL(iBut).Container = sspLevel(iBut)
''''      lblPeakL(iBut).Visible = True
''''
''''      Load lblPeakR(iBut)
''''      Set lblPeakR(iBut).Container = sspLevel(iBut)
''''      lblPeakR(iBut).Visible = True
''''
''''      Load lblMidHL(iBut)
''''      Load lblMidLL(iBut)
''''      Set lblMidHL(iBut).Container = sspLevel(iBut)
''''      Set lblMidLL(iBut).Container = sspLevel(iBut)
''''      lblMidHL(iBut).Visible = True
''''      lblMidLL(iBut).Visible = True
''''
''''      Load lblMidHR(iBut)
''''      Load lblMidLR(iBut)
''''      Set lblMidHR(iBut).Container = sspLevel(iBut)
''''      Set lblMidLR(iBut).Container = sspLevel(iBut)
''''      lblMidHR(iBut).Visible = True
''''      lblMidLR(iBut).Visible = True
      
      Load lblVol(iBut)
      lblVol(iBut).Visible = True
      
''''      Load imgSetup(iBut)
''''      Set imgSetup(iBut).Container = cmdSong(iBut)
''''      imgSetup(iBut).Visible = True
      
     
      '===================================
      'Set the button layout
      '===================================
      cmdSong(iBut).Height = cmdSong(0).Height
      cmdSong(iBut).Width = cmdSong(0).Width
      cmdSong(iBut).DragMode = vbAutomatic
      cmdSong(iBut).OLEDropMode = ssOLEDropManual
      cmdSong(iBut).Picture = Nothing
      cmdSong(iBut).BackColor = cmdReset.BackColor
      cmdSong(iBut).BackStyle = ssOpaque
      cmdSong(iBut).Font.Size = 7
      cmdSong(iBut).Font.Bold = True
      'cmdSong(iBut).Caption = iBut
      'Add the button number here
'      Select Case iPageno
'         Case 1
'            iDispBut = 0
'         Case 2
'            iDispBut = 16
'         Case 3
'            iDispBut = 32
'         Case 4
'            iDispBut = 48
'         Case 5
'            iDispBut = 64
'         Case 6
'            iDispBut = 80
'      End Select
      cmdSong(iBut).Caption = " " & (GetNextButtonNumber + iBut)
      lblButtonCnt(iBut).Caption = cmdSong(iBut).Caption
      'cmdSong(iBut).Caption = " " & iBut
      cmdSong(iBut).ForeColor = vbWhite
      
      sspSongTitle(iBut).Top = sspSongTitle(0).Top
      sspSongTitle(iBut).Left = sspSongTitle(0).Left
      sspSongTitle(iBut).Width = sspSongTitle(0).Width '- 50
      sspSongTitle(iBut).Height = sspSongTitle(0).Height
      sspSongTitle(iBut).Font = sspSongTitle(0).Font
      sspSongTitle(iBut).Font.Size = sspSongTitle(0).Font.Size
      sspSongTitle(iBut).Font.Bold = sspSongTitle(0).Font.Bold
      'sspSongTitle(iBut).PictureAlignment = ssCenterMiddle
     ' sspSongTitle(iBut).OLEDropMode = ssOLEDropManual
      
      lblTimePlayed(iBut).Top = lblTimePlayed(0).Top
      lblTimePlayed(iBut).Left = lblTimePlayed(0).Left
      lblTimePlayed(iBut).Width = lblTimePlayed(0).Width
      lblTimePlayed(iBut).Height = lblTimePlayed(0).Height
      lblTimePlayed(iBut).FontSize = lblTimePlayed(0).FontSize
      lblTimePlayed(iBut).Alignment = 0
           
      lblTimeLeft(iBut).Top = lblTimeLeft(0).Top
      lblTimeLeft(iBut).Left = lblTimeLeft(0).Left
      lblTimeLeft(iBut).Width = lblTimeLeft(0).Width
      lblTimeLeft(iBut).Height = lblTimeLeft(0).Height
      lblTimeLeft(iBut).FontSize = lblTimeLeft(0).FontSize
      lblTimeLeft(iBut).Visible = True
      lblTimeLeft(iBut).Alignment = 1
      
      lblStatus(iBut).Top = lblStatus(0).Top
      lblStatus(iBut).Left = lblStatus(0).Left
      lblStatus(iBut).Width = lblStatus(0).Width
      lblStatus(iBut).Height = lblStatus(0).Height
      lblStatus(iBut).FontSize = lblStatus(0).FontSize
      lblStatus(iBut).ForeColor = vbWhite
      lblStatus(iBut).Visible = False
     
      'imgVol(iBut).Picture = imgVolSource.Picture  ' LoadResPicture(133, vbResIcon)    'LoadPicture(App.Path & "\tmpSpkr")
      
     ' imgVol(iBut).Picture = LoadPicture(App.Path & "\tmpSpkr")
      
      imgVol(iBut).Picture = LoadResPicture(133, vbResIcon)      'imgVolSource.Picture
    '  If bDoEq Then imgEQ(iBut).Picture = imgEQ(0).Picture
      


    '  imgCompleted(iBut).Picture = LoadResPicture(140, vbResIcon)
      
      sspProgress(iBut).Top = sspProgress(0).Top
      sspProgress(iBut).Left = sspProgress(0).Left
      sspProgress(iBut).Width = sspProgress(0).Width
      sspProgress(iBut).Height = sspProgress(0).Height
      
    '  sspProgress(iBut).BevelOuter = ssInsetBevel   ' ssRaisedBevel
      sspProgress(iBut).Visible = False
      sspProgress(iBut).ToolTipText = ""
      
      
      sspProgress(iBut).Appearance = ccFlat  '     cc3D
      sspProgress(iBut).BorderStyle = ccNone
      'Flood Color
      Change_pb_ForeColor sspProgress(iBut).hWnd, vbProgressGreen   '&HFB91&        '&H6FFBB           'vbYellow
      'Control Background color
      Change_pb_Color sspProgress(iBut).hWnd, &H260F35       '&H800000
      
      

      ''''SetProgress iBut, 100
      
   
    '  sspProgress(iBut).BevelWidth = 1
    '  sspProgress(iBut).Font.Size = 8
      'sspProgress(iBut).RoundedCorners = True
    '  sspProgress(iBut).FloodShowPct = False
    '  sspProgress(iBut).FloodPercent = 0
    '  sspProgress(iBut).FloodFillStyle = ssSolid
    '  sspProgress(iBut).FloodType = ssLeftToRight
    '  sspProgress(iBut).BackColor = &H404040
    '  sspProgress(iBut).BackStyle = ssOpaque    'ssTransparent
      'sspProgress(iBut).Picture = sspProgressTemplate.Picture
    
''''      picLevelL(iBut).value = 0
''''      picLevelR(iBut).value = 0
''''      sspLevel(iBut).Left = sspLevel(0).Left
''''      sspLevel(iBut).Top = sspLevel(0).Top
''''      picLevelL(iBut).Top = picLevelL(0).Top
''''      picLevelL(iBut).Left = picLevelL(0).Left
''''
''''      picLevelL(iBut).BorderStyle = ccNone
''''      picLevelL(iBut).Appearance = ccFlat
''''      picLevelR(iBut).BorderStyle = ccNone
''''      picLevelR(iBut).Appearance = ccFlat
    
'''      cpvVolume(iBut).Top = cpvVolume(0).Top
'''      cpvVolume(iBut).Left = cpvVolume(0).Left
'''      cpvVolume(iBut).Width = cpvVolume(0).Width
'''      cpvVolume(iBut).Height = cpvVolume(0).Height
'''      cpvVolume(iBut).RailStyle = SunkenSoft
'''      cpvVolume(iBut).max = 100
'''      cpvVolume(iBut).ShowValueTip = True
'''      cpvVolume(iBut).value = 0
'''      cpvVolume(iBut).Tag = ""
'''      cpvVolume(iBut).BackColor = &H404040
      


'   iTop = 15
'   iLeft = 100  '120
'   iCol = 0
'   iRow = 0
'   bMoveNextBut = False
'   If ApplyStandardTheme Then
'      AdjustButtonHeight = 60
'   Else
'      AdjustButtonHeight = 45
'   End If
   

      
      'Now determine where the next button goes...
      bMoveNextBut = True
      '===========================================================================
      'Top to bottom first
      '===========================================================================
      If iButtonDirection = 1 Then
         iRow = iRow + 1
         If iRow = 1 Then

         ElseIf iRow Mod iMaxRows = 0 Then
            'set the top and left of the current button before we increase the columns
            iTop = iTop + cmdSong(0).Height + AdjustHSpacing
            'iTop = iTop + cmdSong(0).Height + AdjustButtonHeight + AdjustHSpacing
            cmdSong(iBut).Top = iTop
            cmdSong(iBut).Left = iLeft
            bMoveNextBut = False
            iCol = iCol + 1 'Increases the row
            iRow = 0        'sets the column back to 1
            'Sets the top for next row
            iTop = 15   '90
            iLeft = iLeft + cmdSong(iBut).Width + AdjustVSpacing    '30 + 30
         Else
            iTop = iTop + cmdSong(0).Height + (AdjustHSpacing - 30)
            'iTop = iTop + cmdSong(0).Height + AdjustButtonHeight + AdjustHSpacing
         End If
      '===========================================================================
      '  Left to right, then top to bottom
      '===========================================================================
      Else
         iCol = iCol + 1
         If iCol = 1 Then

         ElseIf iCol Mod iMaxCols = 0 Then
            'set the top and left of the current button before we increase the columns
            iLeft = iLeft + cmdSong(iBut).Width + AdjustVSpacing    '30 + 30
            cmdSong(iBut).Top = iTop
            cmdSong(iBut).Left = iLeft
            bMoveNextBut = False
            iRow = iRow + 1 'Increases the row
            iCol = 0        'sets the column back to 1
            'Sets the top for next row
            iLeft = 75  '100
            iTop = iTop + cmdSong(0).Height + (AdjustHSpacing - 30)
            'iTop = iTop + cmdSong(0).Height + AdjustButtonHeight + AdjustHSpacing
         Else
            iLeft = iLeft + cmdSong(iBut).Width + AdjustVSpacing    '30 + 30
         End If
      End If
      
      'set the top and left of the next button
      If bMoveNextBut Then
         cmdSong(iBut).Top = iTop
         cmdSong(iBut).Left = iLeft
      End If
           
      'Debug.Print "Button " & iBut & "  LEFT:" & cmdSong(iBut).Left & "  TOP:" & cmdSong(iBut).Top
       
   Next iBut
'   'Now move the page buttons in place
'   cmdPage(1).Left = 60
'   For i = 1 To 6  'This will give max of 96 songs, dont wanna go over then code changes for 3 digits all over needs to happen
'    cmdPage(i).Top = iTop ' - 30
'    cmdPage(i).Height = cmdPage(0).Height
'    cmdPage(i).Width = cmdPage(0).Width
'    cmdPage(i).ForeColor = cmdPage(0).ForeColor
'    cmdPage(i).BackColor = cmdPage(0).BackColor
'    cmdPage(i).Font.Bold = cmdPage(0).Font.Bold
'    cmdPage(i).Font.Size = cmdPage(0).Font.Size
'    cmdPage(i).BevelWidth = cmdPage(0).BevelWidth
'   Next i
'   cmdPage(2).Left = cmdPage(1).Left + cmdPage(1).Width + 30
'   cmdPage(3).Left = cmdPage(2).Left + cmdPage(2).Width + 30
'   cmdPage(4).Left = cmdPage(3).Left + cmdPage(3).Width + 30
'   cmdPage(5).Left = cmdPage(4).Left + cmdPage(4).Width + 30
'   cmdPage(6).Left = cmdPage(5).Left + cmdPage(5).Width + 30
   
   
   Exit Sub
   
err1:
   MsgBox "Error in Module : SetButtonsLayout " & Chr(13) & Chr(13) & Err.Description, vbExclamation

   Resume Next

End Sub



Private Sub cmdClearPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'cmdClearPalette.Picture = LoadPicture(App.Path & "\tmpInvCLS")
'Sleep 100
'DoEvents

cmdClearPalette.BackColor = &HE7DB49
cmdClearPalette.ForeColor = vbBlack

End Sub

Private Sub cmdClearPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'cmdClearPalette.Picture = LoadPicture(App.Path & "\tmpCLS")
'Sleep 100
'DoEvents

cmdClearPalette.BackColor = vbBlack
cmdClearPalette.ForeColor = &HE7DB49
End Sub

Private Sub cmdCloseVol_Click()

'ButLeft = 22000
'sspVol.Left = ButLeft


End Sub

Private Sub cmdCloseVol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoEvents
Screen.MousePointer = 14  '14
DoEvents
End Sub

Private Sub cmdCloseVol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Closevolume
End Sub

Private Sub cmdExit_Click()

   On Error GoTo err1

   Dim bStopOk As Boolean
   
   bStopOk = True
   For i = 1 To iMaxBut
     If lblStatus(i).Caption = "Playing" Then
       bStopOk = False
       Exit For
     End If
   Next i
   
   If Not bStopOk Then
     Resp = MsgBox(vbTab & "The music is still playing..." & Chr(13) & Chr(13) & "   Are you sure you want to quit  BackTrax Player ??", vbYesNo + vbExclamation, "Exit Program")
     If Resp = vbNo Then Exit Sub
   End If
   
   Unload Me
   
   Exit Sub
   
err1:
   MsgBox "Error in Module : cmdExit_click " & Chr(13) & Chr(13) & Err.Description, vbExclamation


End Sub

Sub ClearButtons()
   On Error Resume Next
   For i = 1 To iMaxBut
     If lblStatus(i).Caption = "Ready" Then
       ClearButton i
     End If
   Next i
   
   GetTotalTime
   
   
End Sub

Sub ShowButtonPlayArea(bShow As Boolean)

   On Error Resume Next
  
   For i = 1 To iMaxBut
      If lblStatus(i).Caption = "Ready" Then
         If bShow Then
            sspSongTitle(i).BorderStyle = 1
         Else
            sspSongTitle(i).BorderStyle = 0
         End If
     End If
   Next i
   
   
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdExit.BackColor = &HE7DB49
cmdExit.ForeColor = vbBlack

End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Shift = 3 Then
   frmSerial.Show vbModal
End If

cmdExit.BackColor = vbBlack
cmdExit.ForeColor = &HE7DB49
End Sub

Private Sub cmdFadeOut_Click()

cmdFadeOut.Enabled = False
FadeSong

End Sub

Sub FadeSong(Optional OverrideFadeValue As Integer)
On Error GoTo err1

Fadeout = GetFadeOutValue

If Fadeout > 7 Then Fadeout = 6
If OverrideFadeValue <> 0 Then Fadeout = OverrideFadeValue

DetermineLastPlaying

If LastPlaying = 0 Then
  cmdFadeOut.Enabled = True
  Exit Sub
End If
'Call BASS_ChannelSlideAttribute(chan(CLng(cmdSong(LastPlaying).Tag)), BASS_ATTRIB_VOL, -1, iDuration)

   cmdFadeOut.BackColor = vbDirectionColor
   cmdFadeOut.ForeColor = vbBlack
   
   cmdFadeOut.Picture = LoadPicture(App.Path & "\tmpInvFade")

fVolume = lblVol(LastPlaying).Caption  'cpvVol.value
'iVolIndex = LastPlaying
tmFade.Interval = iInterval
'SetFadeColor LastPlaying
timFade.Enabled = True
tmFade.Enabled = True

Exit Sub
err1:
MsgBox "Error in Module : cmdFadeOut_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation
End Sub


Function DeterminePlayingCurrently() As Integer
   Dim MaxButs As Integer
   On Error GoTo err1
   
   
   MaxButs = DetermineTotButtons
   For i = 1 To MaxButs
      If lblStatus(i).Caption = "Playing" Then
         DeterminePlayingCurrently = i
         Exit For
      End If
   Next i
   
Exit Function
err1:
MsgBox "Error in Module : DeterminePlayingCurrently " & Chr(13) & Chr(13) & Err.Description, vbExclamation
   
End Function

Sub DetermineLastPlaying()
   Dim MaxButs As Integer
   On Error GoTo err1
   
   LastPlaying = 0
   
   MaxButs = DetermineTotButtons
   For i = 1 To MaxButs
      If lblStatus(i).Caption = "Playing" Then
         LastPlaying = i
      End If
   Next i
   
Exit Sub
err1:
MsgBox "Error in Module : DetermineLastPlaying " & Chr(13) & Chr(13) & Err.Description, vbExclamation
   
End Sub

Sub DetermineLastPaused()
   Dim MaxButs As Integer
   On Error GoTo err1
   
   LastPlaying = 0
   
   MaxButs = DetermineTotButtons
   For i = 1 To MaxButs
      If lblStatus(i).Caption = "Pause" Then
         LastPlaying = i
      End If
   Next i
   
Exit Sub
err1:
MsgBox "Error in Module : DetermineLastPaused " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub cmdLoadPalette_Click()
   On Error GoTo err1
   
   PanMain.Left = 25000
   frmShowPalettes.Show vbModal
    
   If PaletteName <> "" Then
      ClearButtons
      PanMain.Left = -10
      DoEvents
      LoadPalette PaletteName, 1, 1 '1=ALL
      SetPageButton 1
      
      DoEvents
      Me.Caption = UCase(PaletteName)
      lblPaleteName.Caption = Me.Caption
   Else
      PanMain.Left = -10
   End If
   DoEvents
   Command2.SetFocus
   
Exit Sub
err1:
MsgBox "Error in Module : cmdLoadPalette_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub cmdOK_Click()
   On Error GoTo err1
   Dim i As Integer
   For i = 1 To iMaxBut
     If cmdSong(i).Tag <> "" Then
       StopSong i
     End If
   Next i
   
Exit Sub
err1:
MsgBox "Error in Module : cmdOK_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Sub ProcessPause()

On Error GoTo err1

DetermineLastPlaying

If LastPlaying = 0 Then DetermineLastPaused

''''If LastPlaying = 0 Then
''''   If iButtonPlayStopPause = 1 Then
''''      LastPlaying = DetermineLastPlayedSong(DetermineTotButtons)
''''      'Make sure we set the play flag to 0 (play)
''''      iMnuFlag = 0
''''      If LastPlaying <> 0 Then
''''         'Call the routine to start playing a song
''''         Setstate LastPlaying + 1
''''      Else
''''         Setstate 1
''''      End If
''''      Exit Sub
''''   End If
''''End If

'   cmdPause.Picture = LoadPicture(App.Path & "\tmpPause")
'   cmdFadeOut.Picture = LoadPicture(App.Path & "\tmpFade")
   

If LastPlaying = 0 Then Exit Sub
   
'Call BASS_ChannelSlideAttribute(CLng(cmdSong(LastPlaying).Tag), BASS_ATTRIB_VOL, -1, 5)  chan(CLng(cmdSong(LastPlaying).Tag))
If lblStatus(LastPlaying).Caption = "Playing" Then
   lblStatus(LastPlaying).Caption = "Pause"
   cmdPause.Tag = "Resume Song"
   cmdPause.Picture = LoadPicture(App.Path & "\tmpInvPause")
   SetPauseColor LastPlaying
   Call BASS_ChannelPause(chan(CLng(cmdSong(LastPlaying).Tag)))
   SetMainPeakLevel 0, 0
Else
   lblStatus(LastPlaying).Caption = "Playing"
   cmdPause.Tag = "Pause Song"
   cmdPause.Picture = LoadPicture(App.Path & "\tmpPause")
   SetPlayingColor LastPlaying
   Call BASS_ChannelPlay(chan(CLng(cmdSong(LastPlaying).Tag)), BASSFALSE)
   Command2.SetFocus
End If

Exit Sub
err1:
MsgBox "Error in Module : ProcessPause " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub cmdLoadPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdLoadPalette.BackColor = &HE7DB49
cmdLoadPalette.ForeColor = vbBlack

End Sub

Private Sub cmdLoadPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdLoadPalette.ForeColor = &HE7DB49
cmdLoadPalette.BackColor = vbBlack
End Sub

Private Sub cmdOKSecure_Click()

ProcessClickOK

End Sub

Private Sub cmdPage_Click(Index As Integer)
 
 If lblStatus(Player1Index).Caption = "Playing" Or lblStatus(Player2Index).Caption = "Playing" Then Exit Sub
  
  iPageno = Index
  SetPageButton Index
  
  If PaletteName <> "" Then
     ClearButtons
     PanMain.Left = -10
     DoEvents
     LoadPalette PaletteName, iPageno, 2 '2=ONLY page from array
     
     DoEvents
     Me.Caption = UCase(PaletteName)
     lblPaleteName.Caption = Me.Caption
  Else
     PanMain.Left = -10
  End If
  DoEvents
   

End Sub

Sub SetPageButton(iPageno As Integer)

For i = 1 To 6
  cmdPage(i).ForeColor = vbSelected
  cmdPage(i).BackColor = vbBlack
Next i

cmdPage(iPageno).ForeColor = vbBlack
cmdPage(iPageno).BackColor = vbSelected

End Sub

Private Sub cmdPage1_Click(Index As Integer)

End Sub

Private Sub cmdPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ProcessPause
End Sub

Sub ProcessClickOK()

If iSecureMode = 2 Then
  If Trim(txtPassword.text) <> sSecurePWD Then
    MsgBox "The PASSWORD you have entered is INVALID!" & Chr(13) & Chr(13) & "Please try again.", vbExclamation, "Invalid Password"
    txtPassword.text = ""
    txtPassword.SetFocus
    Exit Sub
  End If
  'Only exits if password is valid
  sspSecure.Left = 20000
  bSecureClose = True
End If

End Sub


Private Sub cmdSavePalette_Click()
   On Error GoTo err1
   Dim bPaletteActive As Boolean
   bPaletteActive = False
   
   For i = 1 To iMaxBut
     If sspSongTitle(i).Tag <> "" Then
       bPaletteActive = True
       Exit For
     End If
   Next i
     
   If Not bPaletteActive Then GoTo ExitHere
   
   'Sets the Tag to true, so we can put focus on the correct control in showSongs form, and disable the select option.
   bSavePalette = True
   
  ' PanMain.Left = 25000
   frmShowPalettes.Show vbModal
    
   If PaletteName <> "" Then
      ClearButtons
      PanMain.Left = -10
      DoEvents
      LoadPalette PaletteName, iPageno, 3  '3 = Only reload the array from file
      LoadPalette PaletteName, iPageno, 2  '2=Reload this page from array
      
      DoEvents
      Me.Caption = UCase(PaletteName)
      lblPaleteName.Caption = Me.Caption
   Else
      PanMain.Left = -10
   End If


ExitHere:
   Command2.SetFocus
   DoEvents
Exit Sub
err1:
MsgBox "Error in Module : cmdSavePalette_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub cmdSavePalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSavePalette.BackColor = &HE7DB49
cmdSavePalette.ForeColor = vbBlack

End Sub

Private Sub cmdSavePalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSavePalette.ForeColor = &HE7DB49
cmdSavePalette.BackColor = vbBlack
End Sub

Sub SetSecureMode()

If iSecureMode = 2 Then
   cmdLoadPalette.Enabled = False
   cmdSavePalette.Enabled = False
   cmdClearPalette.Enabled = False
   cmdFadeOut.Enabled = False
Else
   cmdLoadPalette.Enabled = True
   cmdSavePalette.Enabled = True
   cmdClearPalette.Enabled = True
   cmdFadeOut.Enabled = True
End If

End Sub

Private Sub cmdSettings_Click()
   On Error GoTo err1
   Dim sDeviceName As String
   
  txtPassword.text = ""
  bSecureClose = False
  
  If iSecureMode = 2 Then
    sspSecure.Left = vbLoadingLeft
    DoEvents
    txtPassword.SetFocus
    sspSecure.ZOrder 0
RetrySecure:
    Do Until bSecureClose
      Sleep 250
      DoEvents
    Loop
    If Trim(txtPassword.text) <> sSecurePWD Then
      Exit Sub
    End If
    'sspSecure.Left = 20000
  End If
  
   BusyPlaying = False
   
   For i = 1 To iMaxBut
      If Me.lblStatus(i).Caption = "Playing" Then
         'GoTo ExitHere
         BusyPlaying = True
      End If
   Next i
  
  frmSetupSoundCards.Show vbModal
  
  If iButtonMaxSelected <> ScreenOptions(ButMaxSel) Or iButtonDirection <> ScreenOptions(ButDirection) Then
     ReloadScreen
  End If
  
  If lDeviceNo = CLng(Val(sspDevice.Tag)) Or lDeviceNo = -99 Then GoTo ExitHere
   
  DoEvents
  Screen.MousePointer = vbHourglass
  Me.Enabled = False
  If lDeviceNo <> -99 Then
    If Val(sspDevice.Tag) <> 0 Then
        Call BASS_SetDevice(CLng(sspDevice.Tag))  ' set the device to free
        BASS_Free
    End If
    ' setup output devices
    Call BASS_SetDevice(lDeviceNo)  ' set the device to create stream on
    If BASS_Init(lDeviceNo, 44100, BASS_DEVICE_LATENCY, frmPlayer.hWnd, 0) = BASSFALSE Then
      Screen.MousePointer = vbDefault
      sDeviceName = Trim(GetSetting(regMainKey, regSubKey, "Current Device Description"))
      MsgBox "Can't initialize device (" & sDeviceName & ")." & Chr(13) & Chr(13) & "Reason : Device may already be assigned or does not exists." & Chr(13) & Chr(13) & "Reverting to previous sound card.", vbExclamation, "ERROR Setting Sound card"
      Me.Enabled = True
      '============================
      'Reset the device to previous
      lDeviceNo = CLng(Val(sspDevice.Tag))
      SaveSetting regMainKey, regSubKey, "Current Device", lDeviceNo
      SaveSetting regMainKey, regSubKey, "Current Device Description", sspDevice.Caption
      '============================
      
      GoTo ExitHere
    End If
    
    'Set the buffer size...
    buflen = BASS_GetConfig(BASS_CONFIG_BUFFER)
    If buflen < 2000 Then
      Call BASS_SetConfig(BASS_CONFIG_BUFFER, buflen * 5)  'Make buffer 5 times as large...
    End If
    GetDefaultSoundDevice

  End If
    
  Screen.MousePointer = vbDefault
  Me.Enabled = True

ExitHere:
   Command2.SetFocus
   DoEvents
Exit Sub
err1:
MsgBox "Error in Module : cmdSettings_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation
  
End Sub

Private Sub cmdSettings_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSettings.Picture = LoadPicture(App.Path & "\tmpInvSetup")

End Sub

Private Sub cmdSettings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSettings.Picture = LoadPicture(App.Path & "\tmpSetup")
Sleep 100
DoEvents

End Sub

Private Sub cmdSong_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   On Error GoTo err1
   
   If iSecureMode = 2 Then Exit Sub
   
   If TypeOf Source Is SSPanel Then
      If Source.Index = Index Then                                      'Check if it is myself...
        cmdSong_MouseUp Index, 1, 1, 0, 0
        Exit Sub
      End If
      If Trim(sspSongTitle(Index).TagVariant) = "" Then                 'Make sure we do NOT drag button onto an EXISTING one...
         If lblStatus(Source.Index).Caption = "Ready" Then              'Make sure this one is NOT playing
            'If Trim(sspSongTitle(Source.Index).LinkItem) <> "" Then   'Make sure we do NOT drag an EMPTY button...
            If Trim(sspSongTitle(Source.Index).TagVariant) <> "" Then   'Make sure we do NOT drag an EMPTY button...
               Closevolume
               iVolIndex = Index
               ClearButton Index
               sspProgress(Index).Tag = sspProgress(Source.Index).Tag
               sspProgress(Index).max = sspProgress(Source.Index).max
               sspProgress(Index).ToolTipText = sspProgress(Source.Index).ToolTipText
               lblTimeLeft(Index).Tag = lblTimeLeft(Source.Index).Tag
               'sspProgress(Index).TagVariant = sspProgress(Source.Index).TagVariant
               lblVol(Index).Caption = lblVol(Source.Index).Caption
'''''               'SetupButton Index, sspSongTitle(Source.Index).LinkItem, "", sspSongTitle(Source.Index).Tag
               SetupButton Index, sspSongTitle(Source.Index).TagVariant, "", sspSongTitle(Source.Index).Tag
               ClearButton Source.Index
               Source.Caption = " " & Source.Index
               Source.Tag = ""
               'Source.TagVariant = ""
               'Source.LinkItem = ""
               
'''               palletArr(0) = ""
'''               If Trim(lblCurPalette.Caption) = "" Then
'''                  lblCurPalette.Caption = "tmp001"
'''                  palletArr(0) = Trim(lblCurPalette.Caption)
'''               End If
'''
               SavePalete Trim(Me.Caption), iPageno
               LoadPalette Trim(Me.Caption), iPageno, 3 'We only need to reload the array from file
               
               If Not TimerMainLevel.Enabled Then
                  InitialisePeaks 0
               End If
               
            End If
         End If
      Else
        'cmdSong_MouseUp Index, 1, 1, 0, 0
        Exit Sub
      End If
   End If
   
Exit Sub
err1:
MsgBox "Error in Module : cmdSong_DragDrop " & Chr(13) & Chr(13) & Err.Description, vbExclamation
      
End Sub

Sub FixPaletteArray(SourceIndex As Integer, TargetIndex As Integer)

  Dim i As Integer
  PlArr(PLA.efTtle, TargetIndex) = PlArr(PLA.efTtle, SourceIndex) 'Keep Full title here
  PlArr(PLA.eTtle, TargetIndex) = PlArr(PLA.eTtle, SourceIndex)   'Fix the above to show nice title ('Determine if there are "-" in the title array, if so, split the 2)
  PlArr(PLA.eFN, TargetIndex) = PlArr(PLA.eFN, SourceIndex)       'Keep the filename here
  PlArr(PLA.eVol, TargetIndex) = PlArr(PLA.eVol, SourceIndex)     'Get the volume
  PlArr(PLA.eAve, TargetIndex) = PlArr(PLA.eAve, SourceIndex)     'Use this to keep the Average when song is loaded...
  PlArr(PLA.eClr, TargetIndex) = PlArr(PLA.eClr, SourceIndex)     'Color
  'Clear the old data
  For i = 0 To 10
    PlArr(i, SourceIndex) = ""
  Next i
            
End Sub

Sub FixPaletteButtonArray(iOption As Integer, TargetIndex As Integer, sValue As String)

  Select Case iOption
    Case 0
      PlArr(PLA.eTtle, TargetIndex) = sValue   'Fix the above to show nice title ('Determine if there are "-" in the title array, if so, split the 2)
    Case 1
      PlArr(PLA.efTtle, TargetIndex) = sValue  'Keep Full title here
    Case 2
      PlArr(PLA.eFN, TargetIndex) = sValue       'Keep the filename here
    Case 3
      PlArr(PLA.eVol, TargetIndex) = CInt(sValue)     'Get the volume
    Case 4
      PlArr(PLA.eAve, TargetIndex) = CInt(sValue)     'Use this to keep the Average when song is loaded...
    Case 5
      PlArr(PLA.eClr, TargetIndex) = CInt(sValue)     'Color
    Case 6
    
  End Select
            
End Sub

Private Sub cmdSong_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If sspSongTitle(Index).Visible Then
   If lblStatus(Index).Caption = "Playing" Then
      Screen.MousePointer = vbDefault
   Else
      Screen.MousePointer = 15  '99
      'Screen.MouseIcon = lblCursorPlaceHolder.MouseIcon
   End If
Else
   Screen.MousePointer = vbDefault
End If

DoEvents

End Sub

Private Sub cmdSong_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo err1
   Dim iSongs As Integer
   Dim i As Integer

'   Timer3.Enabled = False
'   If LongPress > 40 Then
'      LongPress = 0
'      Exit Sub
'   End If

   If iSecureMode = 2 Then Exit Sub

   If Button = 2 Then 'Right-clicked
     iMnuFlag = 2
     Setstate Index
   Else
      If sspSongTitle(Index).Caption = "" Then
                  ''''         '==================================================================================
                  ''''         'For Demo system, only allow load of 5 songs
                  ''''         If DemoFlag Then
                  ''''            If DetermineTotButtons(True) >= DemoMax Then
                  ''''               MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3, vbExclamation, DemoHeading
                  ''''               Exit Sub
                  ''''            End If
                  ''''         End If
                  ''''         '==================================================================================
                  ''''
                  ''''         iMnuFlag = 2 'Make sure we only load a new song when button is not initialised
                  ''''         Setstate Index
                  ''''
                  ''''         palletArr(0) = ""
                  ''''         If Trim(Me.Caption) = "" Then
                  ''''            Me.Caption = "tmp001"
                  ''''            palletArr(0) = Trim(Me.Caption)
                  ''''         End If
                  ''''         SavePalete Trim(Me.Caption)
         LoadNewSong Index
      Else
'         bTagEditMP3 = UCase(Right(sspSongTitle(Index).Tag, 3)) = "MP3"
'         'FilenameToLoad = sspSongTitle(Index).Tag
'         If lblStatus(Index).Caption = "Ready" Then ShowOptionScreen Index
'
'         SavePalete Trim(Me.Caption), iPageno

      End If
   End If


   Exit Sub

err1:

MsgBox "Error in Module : cmdSong_MouseUp " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Sub LoadNewSong(Index As Integer)

'If sspSongTitle(Index).Caption = "" Then
   '==================================================================================
   'For Demo system, only allow load of 5 songs
   If DemoFlag Then
      If DetermineTotButtons(True) >= DemoMax Then
         MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3, vbExclamation, DemoHeading
         Exit Sub
      End If
   End If
   '==================================================================================
   
   iMnuFlag = 2 'Make sure we only load a new song when button is not initialised
   Setstate Index
   
   palletArr(0) = ""
   If Trim(Me.Caption) = "" Then
      Me.Caption = "tmp001"
      palletArr(0) = Trim(Me.Caption)
   End If
   If iButtonDefaultColor = 2 Then
   
   End If
   SavePalete Trim(Me.Caption), iPageno
   LoadPalette Trim(Me.Caption), iPageno, 3 'We only need to reload the array from file
   
End Sub

Sub ShowPopMenu(Index As Integer)
   On Error GoTo err1
  'Show the popup menu

  If lblStatus(Index).Caption = "Playing" Then
    Exit Sub
  End If
  
  PopupMenu mnuPopUps
  Setstate Index
  
   Exit Sub
   
err1:

MsgBox "Error in Module : ShowPopMenu " & Chr(13) & Chr(13) & Err.Description, vbExclamation
  
End Sub

Sub Setstate(Index As Integer)
   Dim cFreePlayer As Integer
   Dim min As Long, sec As Long, Duration As Long
   'Dim tags As New clsTags
   Dim sListName As String
   Dim sTitle As String
   Dim sArtist As String
    
   On Error GoTo ErrTrap
   
     'Choose the option selected
     Select Case iMnuFlag
       '======================================================================================
       Case 0  'Play                                                                         '
       '======================================================================================
         
         If sspSongTitle(Index).Tag = "" Then Exit Sub
         'If bAsigning Then Exit Sub
         
         If lblStatus(Index).Caption = "Playing" Then
           'Call play song routine here
           StopSong Index
           fVolume = 0
            If Not TimerMainLevel.Enabled Then
               InitialisePeaks 0
            End If
           
         ElseIf lblStatus(Index).Caption = "Pause" Then
            lblStatus(Index).Caption = "Playing"
            cmdPause.Tag = "Pause Song"
            InitialisePeaks 1
            DoEvents
            SetPlayingColor Index
            cmdPause.Picture = LoadPicture(App.Path & "\tmpPause")
            Call BASS_ChannelPlay(chan(CLng(cmdSong(Index).Tag)), BASSFALSE)
         Else
           'determine which player is free
           cFreePlayer = DetermineFreePlayer(Index)
           'Both players are playing, skip out
           If cFreePlayer = 0 Then Exit Sub
           
           
          If iAutoAdvance = 3 Then  'Stop currently playing and start plating selected song
            'find index of playing song
            LastPlaying = 0
            LastPlaying = DeterminePlayingCurrently
            'FadeSong
            If LastPlaying <> 0 Then
              If sspVol.Left <> 22000 Then
                Closevolume True
              End If
              FadeSong 20  'FASTER fadeout and then stop. The Stop is a bit abrubt...
            End If
            'If LastPlaying <> 0 Then StopSong LastPlaying
            DoEvents
          End If
           
           
           
           'Set button statusses etc...
           lblStatus(Index).Caption = "Playing"
           iCntPlayers = cFreePlayer
                   
           'Set the player (1 or 2) to the button
           cmdSong(Index).Tag = iCntPlayers
           If iCntPlayers = 1 Then
             Player1Index = Index
           Else
             Player2Index = Index
           End If
           'Call the play song function
           Playsong iCntPlayers, Index
           
         End If
       '======================================================================================
       Case 2  'Assign                                                                       '
       '======================================================================================
       
         Timer1.Enabled = False
         iFlood = 0
         
         'Timer2.Enabled = True
         Screen.MousePointer = vbHourglass
         DoEvents
         'sspLoading.Left = vbLoadingLeft  '3855
         'sspLoading.ZOrder 0

         DoEvents
         Load frmExplorer
         DoEvents
         'sspLoading.Left = 30000
         
         'Timer2.Enabled = False
         'sspFlood.FloodPercent = 0
         frmExplorer.Show vbModal
         DoEvents
         
         Timer1.Enabled = True
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         
         If FilenameToLoad = "" Or Len(Trim(FilenameToLoad)) < 3 Then
           Screen.MousePointer = vbDefault
           Me.Enabled = True
           Exit Sub
         End If
         
         ClearButton Index
   
         'Reset button to Loaded state
         ResetButton Index
         'Add caption to button
         
         'tags.MP3File = FilenameToLoad
         'GetId3Tags FilenameToLoad
         lblStatus(Index).Caption = "Loading..."
         
         
         ExtractTagInfo FilenameToLoad
         
         SetupButton Index, Id3TagArr(1), Id3TagArr(2), FilenameToLoad
         
         If Not TimerMainLevel.Enabled Then
            InitialisePeaks 0
         End If
         
         
         Screen.MousePointer = vbDefault
         DoEvents
         
       '======================================================================================
       Case 4  'Loadfrom drag and drop                                                       '
       '======================================================================================
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         DoEvents
         If FilenameToLoad = "" Then
           Screen.MousePointer = vbDefault
           Me.Enabled = True
           Exit Sub
         End If
   
         'Reset button to Loaded state
         ResetButton Index
'         If Not TimerMainLevel.Enabled Then
'             pgLeft.value = MaxLevelVal  '30473
'             pgRight.value = MaxLevelVal  '30473
'         End If
         'Add caption to button
         'GetId3Tags FilenameToLoad
         lblStatus(Index).Caption = "Loading..."
         lblTimePlayed(Index).Visible = False
         ExtractTagInfo FilenameToLoad
         SetupButton Index, Id3TagArr(1), Id3TagArr(2), FilenameToLoad
         
         If Not TimerMainLevel.Enabled Then
            InitialisePeaks 0
         End If
       
       '======================================================================================
       Case 3  'Clear                                                                        '
       '======================================================================================
         ClearButton Index
         
       '======================================================================================
       Case 5  'Clear the palette                                                            '
       '======================================================================================
         ClearButtons
     End Select
     
     Exit Sub
     
ErrTrap:
   bAsigning = False
   Screen.MousePointer = vbDefault
   Me.Enabled = True

MsgBox "Error in Module : SetState " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Sub InitialisePeaks(Index As Integer)
Const lIdleColor As Long = &H404000  '&H3514&      '&H242400      '&H32400       '&H404000
Dim i As Integer
On Error GoTo err1

If Index = 1 Then  'Means the full colors
   
   pgLeft.value = 1
   pgRight.value = 1
   DoEvents
   
   Change_pb_ForeColor pgLeft.hWnd, vbGreen    '&HFFAE27
   Change_pb_Color pgLeft.hWnd, vbBlack
   Change_pb_ForeColor pgRight.hWnd, vbGreen   '&HFFAE27
   Change_pb_Color pgRight.hWnd, vbBlack
 
   lblPeakML(0).BackColor = &H47FED0
   lblPeakMR(0).BackColor = &H47FED0
   lblPeakML(1).BackColor = &H46BAFF
   lblPeakMR(1).BackColor = &H46BAFF
   lblPeakML(2).BackColor = &H80FF&
   lblPeakMR(2).BackColor = &H80FF&
   lblPeakML(3).BackColor = &HFF&
   lblPeakMR(3).BackColor = &HFF&
      
   For i = 0 To 3
      lblPeakML(i).Visible = False
      lblPeakMR(i).Visible = False
   Next i
Else
   pgLeft.value = MaxLevelVal  '30473
   pgRight.value = MaxLevelVal  '30473
   
   Change_pb_ForeColor pgLeft.hWnd, lIdleColor    '&HFFAE27
   Change_pb_Color pgLeft.hWnd, vbBlack
   Change_pb_ForeColor pgRight.hWnd, lIdleColor   '&HFFAE27
   Change_pb_Color pgRight.hWnd, vbBlack

   For i = 0 To 3
      lblPeakML(i).BackColor = lIdleColor
      lblPeakMR(i).BackColor = lIdleColor
   Next i
      
   For i = 0 To 3
      lblPeakML(i).Visible = True
      lblPeakMR(i).Visible = True
   Next i
   DoEvents
End If

Exit Sub

err1:
MsgBox Err.Description, vbExclamation, "InitialisePeaks"
            
End Sub

Function GetVolVariables(MaxVol As Single) As String
   On Error Resume Next
   Dim DiffMax As Single
   Dim lMaxVol As Single
   
   On Error GoTo err1
   
   lMaxVol = Round(MaxVol * 100)
   DiffMax = Round(100 - lMaxVol)
   
   GetVolVariables = CStr(lMaxVol) & "|" & CStr(((DiffMax * lMaxVol) / 100) / 10)
   
   Exit Function
   
err1:

MsgBox "Error in Module : GetVolVariables " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Function

Sub SetupButton(Index As Integer, SongTitle As String, SongArtist As String, FileToLoad As String)
   On Error GoTo err1
   
   Dim MaxVol As Single
   Dim PercChange As Single
   Dim iPos As Integer
   Dim iStartPos As Single
   Dim iRandColor As Integer
   
'   If InStr(Trim(SongTitle), "-") > 0 Then
'      iPos = InStr(1, SongTitle, "-")
'      'sspSongTitle(Index).Caption = Trim(Mid(Trim(SongTitle), 1, iPos - 1)) & Chr(13) & " " & Trim(Mid(Trim(SongTitle), iPos + 1))
'      'sspSongTitle(Index).TagVariant = Trim(SongTitle)
'      sspSongTitle(Index).LinkItem = Trim(SongTitle)
'   Else
'     If Trim(SongArtist) = "" Then
'      ' sspSongTitle(Index).Caption = Trim(SongTitle)
'       sspSongTitle(Index).LinkItem = Trim(SongTitle)
'       'sspSongTitle(Index).TagVariant = Trim(SongTitle)
'     Else
'      ' sspSongTitle(Index).Caption = Trim(SongArtist) & Chr(13) & " " & Trim(SongTitle)
'       sspSongTitle(Index).LinkItem = Trim(Replace(SongArtist, "-", "~")) & " - " & Trim(SongTitle)
'       'sspSongTitle(Index).TagVariant = Trim(Replace(SongArtist, "-", "~")) & " - " & Trim(SongTitle)    TagVariant
'     End If
'   End If
   
   If SongTitle = "" And SongArtist = "" Then
      ClearButton Index
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      DoEvents
      Exit Sub
   End If
   
   If Trim(SongArtist) <> "" Then
      'sspSongTitle(Index).LinkItem = Trim(Replace(SongArtist, "-", "~")) & " - " & Trim(SongTitle)
      sspSongTitle(Index).TagVariant = Trim(Replace(SongArtist, "-", "~")) & " - " & Trim(SongTitle)
   Else
      If InStr(Trim(SongTitle), "-") > 0 Then
         iPos = InStr(1, SongTitle, "-")
         'sspSongTitle(Index).LinkItem = Trim(SongTitle)
         sspSongTitle(Index).TagVariant = Trim(SongTitle)
      Else
         'sspSongTitle(Index).LinkItem = Trim(SongTitle)
         sspSongTitle(Index).TagVariant = Trim(SongTitle)
      End If
   End If
   
   
   'sspSongTitle(Index).Caption = FixSongTitle(sspSongTitle(Index).LinkItem)
   sspSongTitle(Index).Caption = FixSongTitle(sspSongTitle(Index).TagVariant)
   
   sspSongTitle(Index).Tag = FileToLoad
   
   
   cmdSong(Index).Tag = 3
   
   If lblVol(Index).Caption = "" Or lblVol(Index).Caption = "0" Then
      MaxVol = Round(GetAverageLevel(FileToLoad, Index) * 100)
      'MaxVol = Round(GetPeakLevel(FileToLoad, Index) * 100)
      If MaxVol > 100 Then MaxVol = 100
      lVolume = MaxVol
    '  If MaxVol < 100 Then
        'If iAdjustVol = 2 Then lVolume = 100 + (100 - lVolume)
        'VolumeChanged Index, True
        SetInitialVolume Index, True
        MaxVol = lVolume
    '  End If
      
      cmdSong(Index).TagVariant = MaxVol    'Use this to keep the Average when song is loaded...
   Else
      MaxVol = CInt(lblVol(Index).Caption)
      'lblStatus(Index).Tag = MaxVol 'This should be loaded from the loadFile method
   End If
   
   If MaxVol < 0 Then  'Error in file, could not load...
     Screen.MousePointer = vbDefault
     MsgBox "Cannot Load the file  !!!! " & Chr(13) & Chr(13) & "The File seems corrupt / zero length", vbCritical, "Error Loading file"
     ClearButton Index
     Me.Enabled = True
     DoEvents
     Exit Sub
   End If
   
   lblTimeLeft(Index).Tag = CStr(ScanForLeadingSilences(FileToLoad, Index))     'CStr(cStartPos)
   lblVol(Index).Caption = MaxVol
   
   'Now determine how much this songs volume needs to be adjusted to become 0
   
   'BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX
   Call BASS_StreamFree(chan(3))    ' free the old stream
   chan(3) = BASS_StreamCreateFile(BASSFALSE, StrPtr(FileToLoad), 0, 0, 0)  'BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX)
   
   Dim Bytes As Long
   Bytes = BASS_ChannelGetLength(chan(3), BASS_POS_BYTE)
   Dim time As Long
   time = BASS_ChannelBytes2Seconds(chan(3), Bytes)
   
   lblTimeLeft(Index).Caption = "00:00"  'Trim((time \ 60) & ":" & Format(time Mod 60, "00"))
   lblTimePlayed(Index).Tag = Trim(Format((time \ 60), "00") & ":" & Format(time Mod 60, "00"))
   lblTimePlayed(Index).Caption = lblTimePlayed(Index).Tag
   'lblTimeLeft(Index).Caption = "/ " & Trim((time \ 60) & ":" & Format(time Mod 60, "00"))
   
   'Set song volume to 1 (which means the vol will increase/decrease...
   'Call BASS_ChannelSetAttribute(chan(3), BASS_ATTRIB_VOL, 1)
   
   'Get the starting position by finding the silence in the front and set start after silence
   'sspProgress(Index).TagVariant = CStr(ScanSilence(FileToLoad))
'   sspProgress(Index).ToolTipText = CStr(ScanSilence(FileToLoad))

   
        
   Call BASS_StreamFree(chan(3))    ' free the old stream
   chan(3) = 0
   
      
   cmdSong(Index).Tag = ""
   ResetButton Index
   
   'cpvVolume(Index).Tag = cpvVolume(Index).value
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
   'cpvVolume(Index).Visible = True
   imgVol(Index).Visible = True
'   If bDoEq Then imgEQ(Index).Visible = True
   
   imgDirection(Index).Visible = True
   
'''   imgSetup(Index).Visible = True
   
   GetTotalTime
   
   If Not TimerMainLevel.Enabled Then
      InitialisePeaks 0
   End If
   
   If vbKeepColor <> 0 Then
      'cmdSong(Index).BackColor = vbKeepColor
      SetButtonColor Index, vbKeepColor
   ElseIf iButtonDefaultColor = 2 Then  'Random
      Randomize
      iRandColor = GenRndNumber(1, 16)
      SetButtonColor Index, iRandColor
   End If
   
   Exit Sub
   
err1:

MsgBox "Error in Module : SetupButton " & Chr(13) & Chr(13) & Err.Description, vbExclamation
'Resume 0

End Sub

Sub GetTotalTime()
Dim sHH As String
Dim sMM As String
Dim sSS As String
Dim sWrk As String
Dim iPos As Integer
Dim iPos1 As Integer
Dim sTot As String
Dim sTotSS As Integer
Dim sTotMM As Integer
Dim sTotHH As Integer
Dim iTotButs As Integer
Dim iTotRec As Integer

On Error GoTo err1

lblTotPlayTime.Caption = "0:00"
lblTotPlayTimeLeft.Caption = "0:00"

''''For iTotRec = 1 To 4
''''   If sspOption3(iTotRec).BackColor = vbDirectionColor Then
''''      iTotButs = sspOption3(iTotRec).Caption
''''      Exit For
''''   End If
''''Next iTotRec

iTotButs = iMaxBut

sTotHH = 0
sTotMM = 0
sTotSS = 0

For iTotRec = 1 To iTotButs
   If lblTimePlayed(iTotRec).Visible = True And Trim(lblTimePlayed(iTotRec).Tag) <> "" Then
      'lblTimePlayed(iInner).Tag
      sWrk = Trim(lblTimePlayed(iTotRec).Tag)
      iPos = InStr(1, sWrk, ":")
      sMM = Mid(sWrk, 1, iPos - 1)
      sSS = Mid(sWrk, iPos + 1)
      
      'Now get total and add this to it
      iPos = InStr(1, lblTotPlayTime.Caption, ":")
      iPos1 = InStr(iPos + 1, lblTotPlayTime.Caption, ":")
      
      If iPos1 > 0 Then  'Means we have an hour in the caption
         sTotHH = Mid(lblTotPlayTime.Caption, 1, iPos - 1)
         sTotMM = Mid(lblTotPlayTime.Caption, iPos + 1, 2)
         sTotSS = Mid(lblTotPlayTime.Caption, iPos1 + 1)
      Else
         sTotMM = Mid(lblTotPlayTime.Caption, 1, iPos - 1)
         sTotSS = Mid(lblTotPlayTime.Caption, iPos + 1)
      End If
      

      sTotMM = sTotMM + sMM
      sTotSS = sTotSS + sSS
      
      If sTotSS > 59 Then
         sTotSS = sTotSS - 60
         sTotMM = sTotMM + 1
      End If
      If sTotMM > 59 Then
         sTotMM = sTotMM - 60
         sTotHH = sTotHH + 1
      End If
      
      If sTotHH > 0 Then
         lblTotPlayTime.Caption = sTotHH & ":" & Format(sTotMM, "00") & ":" & Format(sTotSS, "00")
      Else
         lblTotPlayTime.Caption = sTotMM & ":" & Format(sTotSS, "00")
      End If
   End If
Next iTotRec

sTotHH = 0
sTotMM = 0
sTotSS = 0

For iTotRec = 1 To iTotButs
   If lblTimePlayed(iTotRec).Visible = True And Trim(lblTimePlayed(iTotRec).Tag) <> "" And imgCompleted0(iTotRec).Visible = False Then
      'lblTimePlayed(iInner).Tag
      sWrk = Trim(lblTimePlayed(iTotRec).Tag)
      iPos = InStr(1, sWrk, ":")
      sMM = Mid(sWrk, 1, iPos - 1)
      sSS = Mid(sWrk, iPos + 1)
      
      'Now get total and add this to it
      iPos = InStr(1, lblTotPlayTimeLeft.Caption, ":")
      iPos1 = InStr(iPos + 1, lblTotPlayTimeLeft.Caption, ":")
      
      If iPos1 > 0 Then  'Means we have an hour in the caption
         sTotHH = Mid(lblTotPlayTimeLeft.Caption, 1, iPos - 1)
         sTotMM = Mid(lblTotPlayTimeLeft.Caption, iPos + 1, 2)
         sTotSS = Mid(lblTotPlayTimeLeft.Caption, iPos1 + 1)
      Else
         sTotMM = Mid(lblTotPlayTimeLeft.Caption, 1, iPos - 1)
         sTotSS = Mid(lblTotPlayTimeLeft.Caption, iPos + 1)
      End If
      

      sTotMM = sTotMM + sMM
      sTotSS = sTotSS + sSS
      
      If sTotSS > 59 Then
         sTotSS = sTotSS - 60
         sTotMM = sTotMM + 1
      End If
      If sTotMM > 59 Then
         sTotMM = sTotMM - 60
         sTotHH = sTotHH + 1
      End If
      
      If sTotHH > 0 Then
         lblTotPlayTimeLeft.Caption = sTotHH & ":" & Format(sTotMM, "00") & ":" & Format(sTotSS, "00")
      Else
         lblTotPlayTimeLeft.Caption = sTotMM & ":" & Format(sTotSS, "00")
      End If
      
   End If
Next iTotRec

   Exit Sub
   
err1:

MsgBox "Error in Module : GetTotalTime " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Function DetermineTotButtons(Optional DemoMode As Boolean) As Integer
Dim iTotButs As Integer
Dim iTotRec As Integer
Dim iCntButs As Integer
On Error GoTo err1

DetermineTotButtons = 0

'''For iTotRec = 1 To 4
'''   If sspOption3(iTotRec).BackColor = vbDirectionColor Then
'''      iTotButs = sspOption3(iTotRec).Caption
'''      Exit For
'''   End If
'''Next iTotRec

iTotButs = iMaxBut

If DemoMode Then
   For iTotRec = 1 To iTotButs
      If lblTimePlayed(iTotRec).Visible = True Then
         iCntButs = iCntButs + 1
      End If
   Next iTotRec
   DetermineTotButtons = iCntButs
Else
   DetermineTotButtons = iTotButs
End If

Exit Function
err1:

MsgBox "Error in Module : DetermineTotButtons " & Chr(13) & Chr(13) & Err.Description, vbExclamation


End Function

Sub AddStreamInfo(Index As Integer)

   Dim StreamHandle As Long
   On Error GoTo err1
   
   StreamHandle = BASS_StreamCreateFile(BASSFALSE, StrPtr(sspSongTitle(Index).Tag), 0, 0, 0)
   
   If StreamHandle = 0 Then
       MsgBox "Can't open stream"
   End If
   
Exit Sub
err1:

MsgBox "Error in Module : AddStreamInfo " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub


Function DetermineFreePlayer(Index As Integer) As Integer
   On Error GoTo err1
   
   DetermineFreePlayer = 0
   'Allow only one player because user selected only one stream...
   If iButtonStreams = 1 Then
      If chan(1) <> 0 Or chan(2) <> 0 Then Exit Function
   End If
   
   'Check which channel is open
   If chan(1) = 0 Then
     DetermineFreePlayer = 1
   ElseIf chan(2) = 0 Then
     DetermineFreePlayer = 2
   End If
   
   Exit Function

err1:

MsgBox "Error in Module : DetermineFreePlayer " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Function

Sub StopPlayingSong(Index As Integer)
On Error Resume Next
   
   'If SkipBassStop Then Call BASS_ChannelStop(chan(CInt(cmdSong(Index).Tag)))
   chan(CInt(cmdSong(Index).Tag)) = 0
   
   ResetButton Index, True
   
   lblStatus(Index).Caption = "Ready"
   cmdPause.Tag = ""
   cmdPause.BackColor = vbBlack
   cmdPause.Picture = LoadPicture(App.Path & "\tmpPause")

   If Not TimerMainLevel.Enabled Then
      InitialisePeaks 0
   End If
   
End Sub

Sub StopSong(Index As Integer)

   On Error Resume Next
   
   Call BASS_ChannelStop(chan(CInt(cmdSong(Index).Tag)))
   chan(CInt(cmdSong(Index).Tag)) = 0
   
   ResetButton Index
   
'   TimerMainLevel.Enabled = TimerP1Level.Enabled Or TimerP2Level.Enabled
   
   lblStatus(Index).Caption = "Ready"
   cmdPause.Tag = ""
   cmdPause.BackColor = vbBlack
   cmdPause.Picture = LoadPicture(App.Path & "\tmpPause")
   
   
''''   If InStr(1, sspStream1.Caption, "Ready") > 0 And _
''''      InStr(1, sspStream2.Caption, "Ready") > 0 Then
''''      'sspButtonDirMain.Enabled = True
''''      SSPanel3.Enabled = True
''''   Else
''''      'sspButtonDirMain.Enabled = False
''''      SSPanel3.Enabled = False
''''   End If

   'SetMainPeakLevel 0, 0
   If Not TimerMainLevel.Enabled Then
      InitialisePeaks 0
   End If

End Sub

Sub StartNextAvalableSong(Index As Integer)
On Error Resume Next

'First check if something else is playing TagVariant
For i = 1 To iMaxBut
   If lblStatus(i).Caption = "Playing" Then
      Exit Sub
   End If
Next i
For i = Index + 1 To iMaxBut
   'If sspSongTitle(I).LinkItem <> "" Then
   If sspSongTitle(i).TagVariant <> "" Then
      If imgCompleted0(i).Visible <> True Then
         iMnuFlag = 0
         Setstate i
         Exit Sub
      End If
   End If
Next i

End Sub

Sub SetPlayingColor(Index As Integer)
   sspSongTitle(Index).BackColor = vbYellow
   sspSongTitle(Index).ForeColor = vbBlack
   
   lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
   
   cmdPause.BackColor = vbBlack
   cmdPause.ForeColor = vbDirectionColor
End Sub

Sub SetPauseColor(Index As Integer)
   
   sspSongTitle(Index).BackColor = vbYellow
   sspSongTitle(Index).ForeColor = vbBlack
   
   lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
   
   cmdPause.BackColor = vbDirectionColor
   cmdPause.ForeColor = vbBlack
   
   
'   sspSongTitle(Index).BackColor = vbBlack
'   sspSongTitle(Index).ForeColor = vbGreen
'
End Sub

Sub SetPlayColor(Index As Integer, channel As Integer)
   Dim iColor As Integer
   On Error Resume Next
     
   lblStatus(Index).Caption = "Playing"
   lblStatus(Index).ForeColor = vbBlack
   'lblTimePlayed(Index).FontSize = 9
   lblTimePlayed(Index).FontBold = True
   lblTimePlayed(Index).FontSize = lblTimePlayed(Index).FontSize + 1
'''   lblTimePlayed(Index).Height = 195
'''   lblTimePlayed(Index).Top = lblTimePlayed(0).Top - 10
   lblTimePlayed(Index).Caption = lblTimePlayed(Index).Tag  '   Trim(Replace(lblTimeLeft(Index).Caption, "/", ""))
   If Val(sspProgress(Index).Tag) = 0 Then
      lblTimePlayed(Index).ForeColor = vbYellow  'vbBlack
   Else
      lblTimePlayed(Index).ForeColor = vbYellow
   End If
   
   imgCompleted0(Index).Visible = False
   imgCompleted1(Index).Visible = False
   imgCompleted2(Index).Visible = False
   imgCompleted3(Index).Visible = False
   
''''   lblPeakL(Index).Visible = False
''''   lblPeakR(Index).Visible = False
''''   lblMidHL(Index).Visible = False
''''   lblMidHR(Index).Visible = False
''''   lblMidLL(Index).Visible = False
''''   lblMidLR(Index).Visible = False
   
   cmdSong(Index).Picture = Nothing
   'cpvVolume(Index).BackColor = cmdSong(Index).BackColor
   
   iColor = Val(sspProgress(Index).Tag) 'VAL()   Will also force to 0 if nothing setup
   cmdSong(Index).BackColor = ButColors(iColor)
  ' cpvVolume(Index).BackColor = ButColors(iColor)
   
   sspSongTitle(Index).ForeColor = vbBlack
   lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
  ' sspSongTitle(Index).BevelOuter = ssInsetBevel
   sspSongTitle(Index).BackStyle = 1
   sspSongTitle(Index).ForeColor = vbBlack
   sspSongTitle(Index).BackColor = vbNDYellow
   sspSongTitle(Index).Font.Bold = True
  'sspSongTitle(Index).BevelWidth = 1
  ' sspSongTitle(Index).Outline = False
   
   InitialisePeaks 1


   Select Case channel
      Case 1
          TimerF1.Enabled = False
          'If Len(sspSongTitle(Index).LinkItem) > 63 Then 'LinkItem
          If Len(sspSongTitle(Index).TagVariant) > 63 Then
            'sspStream1.Caption = Left(sspSongTitle(Index).LinkItem, 63) & "..."
            sspStream1.Caption = Left(sspSongTitle(Index).TagVariant, 63) & "..."
          Else
            'sspStream1.Caption = sspSongTitle(Index).LinkItem
            sspStream1.Caption = sspSongTitle(Index).TagVariant
          End If
          
      Case 2
         TimerF2.Enabled = False
         'If Len(sspSongTitle(Index).LinkItem) > 63 Then
         If Len(sspSongTitle(Index).TagVariant) > 63 Then
           'sspStream2.Caption = Left(sspSongTitle(Index).LinkItem, 63) & "..."
           sspStream2.Caption = Left(sspSongTitle(Index).TagVariant, 63) & "..."
         Else
           'sspStream2.Caption = sspSongTitle(Index).LinkItem
           sspStream2.Caption = sspSongTitle(Index).TagVariant
         End If

   End Select
   lblStream(Index).ForeColor = ButForColors(iColor)
   
   sspProgress(Index).Visible = True
   
End Sub

Sub Playsong(channel As Integer, Index As Integer)
   Dim iVol As Single
   Dim DataLength As Long
   
   On Error GoTo err1
   
   SetPlayColor Index, channel
  ' If imgVol(iVolIndex).Visible = False And sspVol.Left = 22000 Then Closevolume
   
   lblStream(Index).Caption = "Stream  :  " & channel
   lblStream(Index).Tag = channel
 '  lblStream(Index).Visible = True
      
   Call BASS_StreamFree(chan(channel))
   'Call BASS_SetDevice(lDeviceNo)  ' set the device to create stream on
   'chan(channel) = BASS_StreamCreateFile(BASSFALSE, StrPtr(sspSongTitle(Index).Tag), 0, 0, IIf(bDoEq, 0 Or floatable, 0)) ' Or BASS_SAMPLE_FX)   ' Or BASS_SAMPLE_FX)  ' Or BASS_STREAM_AUTOFREE)
   chan(channel) = BASS_StreamCreateFile(BASSFALSE, StrPtr(sspSongTitle(Index).Tag), 0, 0, 0)

  ' chan(channel) = BASS_StreamCreateFile(BASSFALSE, StrPtr(sspSongTitle(Index).Tag), 0, 0, BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX)
   'chan(channel) = BASS_StreamCreateFile(BASSFALSE, StrPtr(sspSongTitle(Index).Tag), 0, 0, 0 Or BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX)
   cmdSong(Index).Tag = channel
   
   DataLength = FileLen(sspSongTitle(Index).Tag)
   
   lVolume = lblVol(Index).Caption
      
   VolumeChanged Index
      
   iPlayingChan = chan(channel)
      
'''   'Setup the Effects Chan info so we can do EQ
'''   SetupEqFx iPlayingChan
'''   'If any EQ in the array is NOT = 10 then Do this
'''   For iEqBands = 1 To MaxEqs
'''      If aEQ(Index, iEqBands) <> 10 Then
'''        UpdateSongFreqEq iEqBands, aEQ(Index, iEqBands)
'''      End If
'''   Next iEqBands
     
   If CDbl(lblTimeLeft(Index).Tag) > 0 And iButtonRemoveSilence = 1 Then
    '  Debug.Print "Starting at : " & CDbl(lblTimeLeft(Index).Tag)
      If lblTimeLeft(Index).Tag <> "" Then
         Call BASS_ChannelSetPosition(chan(channel), BASS_ChannelSeconds2Bytes(chan(channel), CDbl(lblTimeLeft(Index).Tag)), BASS_POS_BYTE)      ' set the position
      End If
   End If
   Duration(channel) = Format(bassTime.GetDuration(chan(channel)), "0")
   
   Select Case channel
     Case 1
         With bassTime
         '    lblTimeLeft(index).Tag = Format(.GetDuration(chan(channel)), "0.0")    '& " seconds / " & bassTime.GetTime(bassTime.GetDuration(chan))
         ''    frmMemory.lblFreq.Caption = "Frequency: " & .GetFrequency(chan) & " Hz, " & .GetBits(chan) & " bits, " & .GetMode(chan)
         ''    frmMemory.lblBPS.Caption = "Bytes/s: " & .GetBytesPerSecond(chan)
         ''    frmMemory.lblBitsPS.Caption = "Kbp/s: " & .GetBitsPerSecond(chan, DataLength) & " [average kbp/s for vbr mp3s]"
            sspStream1.Caption = sspStream1.Caption & " (" & .GetBitsPerSecond(chan(channel), DataLength) & " Kbp/s)"
            lblStream(Index).Caption = "Stream  :  " & channel & "     (" & .GetBitsPerSecond(chan(channel), DataLength) & "  Kbp/s)"
         End With

       TimerP1.Enabled = True
       TimerP1Level.Enabled = True
     
     Case 2
         With bassTime
            sspStream2.Caption = sspStream2.Caption & " (" & .GetBitsPerSecond(chan(channel), DataLength) & " Kbp/s)"
            lblStream(Index).Caption = "Stream  :  " & channel & "     (" & .GetBitsPerSecond(chan(channel), DataLength) & "  Kbp/s)"
         End With
       TimerP2.Enabled = True
       TimerP2Level.Enabled = True
   End Select
   
   Call BASS_ChannelPlay(chan(channel), BASSFALSE)           ' play new stream
   
   TimerMainLevel.Enabled = TimerP1Level.Enabled Or TimerP2Level.Enabled
      
Exit Sub
err1:

MsgBox "Error in Module : PlaySong " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub


Function ShowSongTitleOnly(sFilename As String) As String
   Dim sWork As String
   Dim iPos As Integer
   
   On Error Resume Next
   
   ShowSongTitleOnly = ""
   iPos = InStr(1, sFilename, ".") - 1
   
   ShowSongTitleOnly = Mid(sFilename, 1, iPos)

End Function

Private Sub cmdSong_OLEDragDrop(Index As Integer, Data As Threed.SSDataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo err1
   If Effect = 7 Then
      '==================================================================================
      'For Demo system, only allow load of 5 songs
      If DemoFlag Then
         If DetermineTotButtons(True) >= DemoMax Then
            MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3, vbExclamation, DemoHeading
            Exit Sub
         End If
      End If
      '==================================================================================
   
     FilenameToLoad = Data.Files(1)
     iMnuFlag = 4
     Setstate Index
           
      palletArr(0) = ""
      If Trim(Me.Caption) = "" Then
         Me.Caption = "tmp001"
         palletArr(0) = Trim(Me.Caption)
      End If
     SavePalete Trim(Me.Caption), iPageno
     LoadPalette Trim(Me.Caption), iPageno, 3
     
      If Not TimerMainLevel.Enabled Then
         InitialisePeaks 0
      End If
     
   End If
   
Exit Sub
err1:

MsgBox "Error in Module : cmdSong_OLEDragDrop " & Chr(13) & Chr(13) & Err.Description, vbExclamation
   
End Sub

Private Sub SetInitialVolume(Index As Integer, Optional bLoading As Boolean)
  Dim iVol As Single
   Dim bSetVol As Boolean
   
   On Error Resume Next
      
   iVol = lVolume / 100
   
   If iVol > 0.99 Then iVol = 0.99
   If iAdjustVol = 2 Then  'iAdjustVol = 1 :Set to 75%, iAdjustVol = 2 : Set to 100%
    iVol = 0.99
    'iVol = lVolume / 100
    'If lVolume < 100 Then lVolume = 100
    lVolume = 100
   Else 'Set to 75%, unless less than 75%, then make it 100%
    If iVol < 0.75 Then
      iVol = 0.99
      lVolume = 100
    Else
      iVol = 0.75
      lVolume = 75
    End If
   End If
   
   If cmdSong(Index).Tag = 1 Then
      iVol1 = iVol
   Else
      iVol2 = iVol
   End If
   
   lblVol(Index).Caption = lVolume
   lblVolInd.Caption = lVolume & " %"

   If Val(cmdSong(Index).Tag) <> 0 Then
       Call BASS_ChannelSetAttribute(chan(CInt(cmdSong(Index).Tag)), BASS_ATTRIB_VOL, iVol)
   End If
End Sub

Private Sub VolumeChanged(Index As Integer, Optional bLoading As Boolean)
   Dim iVol As Single
   Dim bSetVol As Boolean
   
   On Error Resume Next
      
   iVol = lVolume / 100
   
   If iVol > 0.99 Then iVol = 0.99
'   If iAdjustVol = 2 And bLoading Then
'    iVol = 0.99
'    'iVol = lVolume / 100
'    If lVolume < 100 Then lVolume = 100
'   Else
'    If iVol < 0.75 Then
'      iVol = 0.99
'      lVolume = 100
'    Else
'      iVol = 0.75
'      lVolume = 75
'    End If
'   End If
   
   If cmdSong(Index).Tag = 1 Then
      iVol1 = iVol
   Else
      iVol2 = iVol
   End If

   lblVol(Index).Caption = lVolume
   lblVolInd.Caption = lVolume & " %"
'   DoEvents
   
   If Val(cmdSong(Index).Tag) <> 0 Then
    '  bSetVol = BASS_SetVolume(iVol)
       Call BASS_ChannelSetAttribute(chan(CInt(cmdSong(Index).Tag)), BASS_ATTRIB_VOL, iVol)
   End If

End Sub

Sub SetSongVolume(Index As Integer)
   Dim iVol As Single
   
   On Error Resume Next
    
   iVol = fVolume / 100
   
'   If iVol > 0.99 Then iVol = 0.99
'   If cmdSong(Index).Tag = 1 Then
'      iVol1 = iVol
'   Else
'      iVol2 = iVol
'   End If
   
  ' If Val(cmdSong(Index).Tag) <> 0 Then
       Call BASS_ChannelSetAttribute(chan(CInt(cmdSong(Index).Tag)), BASS_ATTRIB_VOL, iVol)
  ' End If
   
End Sub

Private Sub Command1_Click()

'If Me.WindowState = 0 Then
'   Me.WindowState = 1
'   Me.Caption = "BackTrax Player"
'End If

End Sub

Private Sub Command2_Click()
'Closevolume
End Sub

'''Private Sub cpvBass_ValueChanged()
'''
'''   lBass = cpvBass.value
'''   lblBass(iVolIndex).Caption = lBass
'''   If cmdSong(iVolIndex).Tag <> "" Then Call UpdateFX(CInt(cmdSong(iVolIndex).Tag), 0, cpvBass.value)   ' bass
'''End Sub
'''
'''Private Sub cpvHigh_ValueChanged()
'''
'''   lHigh = cpvHigh.value
'''   lblHigh(iVolIndex).Caption = lHigh
'''   If cmdSong(iVolIndex).Tag <> "" Then Call UpdateFX(CInt(cmdSong(iVolIndex).Tag), 2, cpvHigh.value)   ' treble
'''End Sub
'''
'''Private Sub cpvMid_ValueChanged()
'''
'''   lMid = cpvMid.value
'''   lblMid(iVolIndex).Caption = lMid
'''   If cmdSong(iVolIndex).Tag <> "" Then Call UpdateFX(CInt(cmdSong(iVolIndex).Tag), 1, cpvMid.value)    ' mid
'''End Sub

'''Private Sub cpvVol_ValueChanged()
'''
''''lblVol(iVolIndex).Caption = cpvVol.value
'''lVolume = Format(Slider(Index).sldCrntVal, "0.0")
'''VolumeChanged iVolIndex
'''
'''End Sub

'''''Private Sub cpvVol_Change()
'''''lVolume = cpvVol.value
'''''VolumeChanged iVolIndex
'''''End Sub

'''Private Sub cpvVol_Scroll()
'''lVolume = cpvVol.value
'''VolumeChanged iVolIndex
'''End Sub

'''Private Sub cpvVol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''Dim NewPerc As Single
'''Dim NewStart As Integer
'''Const MaxTwip = 3900
'''
'''On Error Resume Next
'''
'''   'If Me.lblStatus(Index).Caption <> "Playing" Then Exit Sub
'''  If X > MaxTwip Then
'''    cpvVol.Value = 100
'''  Else
'''    cpvVol.Value = (X * 100) / MaxTwip
'''  End If
'''  lVolume = cpvVol.Value
'''  VolumeChanged iVolIndex
   
'   SetSongPos chan(CLng(cmdSong(Index).Tag)), Index, CInt(x)
'
'
'   NewPerc = (sPos / MaxWidth) * 100  'Percentage of where I need to start
'   NewStart = ((NewPerc * Duration(CLng(cmdSong(Index).Tag))) / 100)
   
'End Sub

Private Sub cpvVol_Change()

lVolume = cpvVol.value
VolumeChanged iVolIndex

End Sub

Private Sub cpvVol_SliderScroll()
lVolume = cpvVol.value

lblVolInd.Caption = lVolume & " %"
lblVolInd.Refresh

VolumeChanged iVolIndex
End Sub

Private Sub Form_Activate()
Dim errCnt As Integer
   On Error GoTo err1
   
   WriteLog " "
   WriteLog "  - - - - - - - - - - - - - - - - -"
   WriteLog "  frmPlayer : Form_Activate START ..."
   WriteLog "  frmPlayer : SetSecureMode ..."
   SetSecureMode
   
  ButColors(0) = vbNDefault
  ButForColors(0) = vbNDefaultFore
  
  cpvVol.BackColor = vbBlack
  cpvVol.SliderColor = vbCyan  'vbProgressGreen
   
   errCnt = 0
 '  EnableCloseButton Me.hWnd, False
   errCnt = errCnt + 1  '1
   If PaletteName <> "" Then
      Me.Caption = UCase(PaletteName)
   Else
      PaletteName = Me.Caption
   End If
   
   errCnt = errCnt + 1  '2
   If DemoFlag Then
      SSPanel9.Caption = "DEMO MODE"
      SSPanel9.MarqueeStyle = ssBlinkingMarquee
      sspVersion(3).Caption = ""
   Else
      sspVersion(3).Caption = "Serial: " & GetSetting(regMainKey, regSubKey, "SerialNumber")
      SSPanel9.Caption = ""
      SSPanel9.MarqueeStyle = ssNoneMarquee
   End If
   
   errCnt = errCnt + 1  '3
   'Check if the SETTING values have changed, if so, redo the relevant settings and screen changes
   'ScreenOptions(ButMaxSel) = iButtonMaxSelected
'''   If iButtonMaxSelected <> ScreenOptions(ButMaxSel) Or iButtonDirection <> ScreenOptions(ButDirection) Then
'''      ReloadScreen
'''   End If
   
'''   If iButtonMaxSelected <> Val(GetSetting(regMainKey, regSubKey, "Max Buttons")) Or _
'''      iButtonDirection <> Val(GetSetting(regMainKey, regSubKey, "ButtonDirection")) Then
'''      ReloadScreen
'''   End If
   
   errCnt = errCnt + 1  '4
   WriteLog "  frmPlayer : LoadStats ..."
   LoadStats
   lstSystem.Visible = False
'   sspDevice.Visible = False
   errCnt = errCnt + 1  '5
   If Not TimerMainLevel.Enabled Then
      WriteLog "  frmPlayer : InitialisePeaks 0 ..."
      InitialisePeaks 0
   End If
   
   errCnt = errCnt + 1  '6
   Me.lblPaleteName.Caption = Me.Caption
   'Me.SetFocus
'   Me.Command2.SetFocus
   WriteLog "  frmPlayer : Form_Activate END ..."
   WriteLog "  - - - - - - - - - - - - - - - - -"
   WriteLog " "
   Exit Sub
   
err1:

MsgBox "ERROR " & Chr(13) & Chr(13) & "The following error has occurred at error count : " & errCnt & Chr(13) & Chr(13) & "ERROR : " & Err.Description & " (" & Err.Number & ")", vbExclamation, "MAIN"
' Resume

End Sub

Sub ReloadScreen()
On Error GoTo err1

'Exit when no change detected...
'If iButtonMaxSelected = Val(GetSetting(regMainKey, regSubKey, "Max Buttons")) Then Exit Sub
   
'Save this in order for next change to be able to test  correctly
'SaveSetting regMainKey, regSubKey, "ButtonDirection", iButtonDirection
'SaveSetting regMainKey, regSubKey, "Max Buttons", iButtonMaxSelected
   
   
'Reload screen since the buttons to display has changed...
Select Case iButtonMaxSelected
   Case 1
      iMaxBut = 9
   Case 2
      iMaxBut = 16
   Case 3
      iMaxBut = 20
   Case 4
      iMaxBut = 30
End Select

'Hide the volume screen, just in case...
ButLeft = 22000
sspVol.Left = ButLeft

'Disable the form
PanMain.Visible = False
frmPlayer.Enabled = False
Screen.MousePointer = vbHourglass
DoEvents
DoEvents
'Clear and load buttons in new layout
RemoveButtons
DoEvents
SetButtonsLayout
DoEvents
ClearButtons
DoEvents
LoadPalette PaletteName, iPageno, 2 'Only load current page from array
'Enable the screen again...
frmPlayer.Enabled = True
PanMain.Visible = True
DoEvents

'Set the global variables to test against
ScreenOptions(ButMaxSel) = iButtonMaxSelected
ScreenOptions(ButDirection) = iButtonDirection

Screen.MousePointer = vbDefault
DoEvents

Exit Sub
err1:
MsgBox "Error in Module : sspOption3_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub



'''Private Sub SetProgressBackground(PgIndex As Integer, Index As Integer)
'''
'''    Dim LngNew As Long
'''
'''On Error GoTo err1
'''
'''    'Use a picture with the FLAT TB (Toolbar1)
'''    LngNew = CreatePatternBrush(picProgress(Index).Picture.handle) 'Creates the background from a Picture Handle
'''    ChangeTBBack sspProgress(PgIndex), LngNew, enuTB_FLAT
'''
''''''    'Change Backcolor to STANDARD TB (Toolbar2)
''''''    LngNew = CreateSolidBrush(RGB(240, 120, 120))        'Creates the background from a Color (Long)
''''''    ChangeTBBack Toolbar2, LngNew, enuTB_STANDARD
'''
'''    'Refresh Screen to see changes
'''    InvalidateRect 0&, 0&, False
'''
'''
'''Exit Sub
'''err1:
'''
'''MsgBox "Error in Module : SetProgressBackground " & Chr(13) & Chr(13) & Err.Description, vbExclamation
'''
'''End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
   If sspVol.Left < 21000 Then
      Closevolume
   End If
End If



End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim LastPlayed As Integer
Dim MaxButs As Integer

On Error GoTo err1

If KeyCode = 32 Then 'Space pressed, start next song
   
'''   If iButtonPlayStopPause = 1 Then    'Pause
'''      ProcessPause
'''      Exit Sub
'''   End If
   
   MaxButs = DetermineTotButtons
   
   LastPlayed = DetermineLastPlayedSong(MaxButs)
   
'   For i = 1 To MaxButs
'      'If cmdSong(i).Picture = cmdReset.Picture Then
'      'Check the imgCompleted position, or visibility
'      If imgCompleted(i).Visible = True Then
'         LastPlayed = i
'      End If
'   Next i
   'Make sure the buttons on top does NOT have focus, otherwise wgron screens are loaded
   Command2.SetFocus
   'Make sure we do NOT try to play past last button on screen
   If LastPlayed + 1 > MaxButs Then Exit Sub
   'Determine which song to play next. Will check the title of buttons
   For i = LastPlayed + 1 To MaxButs
      If sspSongTitle(i).Caption <> "" Then
         Exit For
      End If
   Next i
   'Make sure we set the play flag to 0 (play)
   iMnuFlag = 0
   'Call the routine to start playing a song
   Setstate i
End If

Exit Sub
err1:

MsgBox "Error in Module : Form_KeyUp " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Function DetermineLastPlayedSong(MaxButs As Integer) As Integer
Dim i As Integer
On Error GoTo err1

   DetermineLastPlayedSong = 0
   For i = 1 To MaxButs
      If imgCompleted0(i).Visible = True Then
         DetermineLastPlayedSong = i
      End If
   Next i
   
Exit Function
err1:

MsgBox "Error in Module : DetermineLastPlayedSong " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Function

Sub LoadSoundDevices()

On Error Resume Next

Dim c As Integer
Dim i As BASS_DEVICEINFO

c = 1      ' device 1 = 1st real device
frmLoader.Label2.Visible = True
'DoEvents
While BASS_GetDeviceInfo(c, i)
  If (i.flags And BASS_DEVICE_ENABLED) Then  ' enabled, so add it...
      frmLoader.lstPlugins(1).Caption = frmLoader.lstPlugins(1).Caption & VBStrFromAnsiPtr(i.name) & Chr(13)
  End If
  c = c + 1
Wend
   
End Sub

Private Sub Form_Load()
   Dim iDeviceCnt As Long
   Dim sDevices As String
   Dim errCnt As Integer
   
   errCnt = 0
   
   On Error GoTo err1
   errCnt = errCnt + 1 '1

  ' MsgBox "Start Form Load"
   WriteLog " "
   WriteLog "**************************************************************************************"
   WriteLog "frmPlayer : Form_Load START ..."
   
   bLoading = True
   
   errCnt = errCnt + 1 '16
   On Error Resume Next
   WriteLog "frmPlayer : Create Palets directory ..."
   MkDir App.Path & "\Palets"
   DoEvents
   On Error GoTo 0
   On Error GoTo err1
   
   'Always check to convert any older format files
   WriteLog "frmPlayer : ConvertOldPalettes ..."
   ConvertOldPalettes
   
   lblCompiled.Caption = "Created on : " & FileDateTime(Win3(App.Path) & "\" & App.EXEName & ".EXE")
   
   If IsThemeActive() = False Then
      frmLoader.lblTheme.Caption = "THEME : FALSE"
   Else
      frmLoader.lblTheme.Caption = "THEME : TRUE"
   End If
   
   frmLoader.ShowLoading MsgLits.ChkDisk, 150 'Checking Disk Space
   Sleep 100
   DoEvents

   errCnt = errCnt + 1 '2
   WriteLog "frmPlayer : GetFreeSpace ..."
   GetFreeSpace Left(App.Path, 3)
   WriteLog "frmPlayer : GetFreeSpace = " & Round((DiskSpaceFreeMB / 1024), 1) & " GB"
   If CCur(DiskSpaceFreeMB) < 100 Then
      MsgBox "Disk space too low !!!!" & Chr(13) & Chr(13) & "Space Available : " & DiskSpaceFreeMB & " mb" & Chr(13) & Chr(13) & "Please free up some space before continuing.", vbCritical, "Low Disk Space"
      End
   End If

   lblPaleteName.BackStyle = 0
   lblPaleteName.BackColor = vbBlack
      
'   DiskSpaceTot = Format$(Tot * 10000, "###,###,###,##0")
'   DiskSpaceFree = Format$(Free * 10000, "###,###,###,##0")
   frmLoader.ShowLoading MsgLits.LoadEnv, 150 'Loading Environment Variables...
   Sleep 100
   DoEvents

   'Set Program version label
   WriteLog "frmPlayer : Set Version ..."
   lblVersion.Caption = "Version   :    " & App.Major & "." & App.Minor & "." & App.Revision
'   lblVersion.Top = 705   '750
   lblVersion.Left = 255
   lblVersion.Width = 2940
   lblVersion.FontSize = 7
''   lblCompiled.Top = 855
'''   sspDevice.Top = 975

   frmLoader.ShowLoading MsgLits.LoadSoundCards, 250  'Load Sound Cards Devices...
   Sleep 300
   DoEvents

   LoadSoundDevices
   Sleep 200
   DoEvents

   'sspVersion.Caption = "Disk Space: " & DiskSpaceTot & " gb       Free Space: " & DiskSpaceFree & " gb       Version " & App.Major & "." & App.Minor & "." & App.Revision

   'sspVersion.Caption = "Disk Space: " & DiskSpaceTot & " gb       Free Space: " & DiskSpaceFree & " gb       Version " & App.Major & "." & App.Minor & "." & App.Revision
   errCnt = errCnt + 1 '3
   WriteLog "frmPlayer : Create temporary Images ..."
'   LoadDataIntoFile 150, App.Path & "\tmpBanner" 'Main banner pic
'   Me.Picture = LoadPicture(App.Path & "\tmpBanner")

   LoadDataIntoFile 151, App.Path & "\tmpPause"    'Pause
   LoadDataIntoFile 152, App.Path & "\tmpInvPause" 'Inverted Pause
   LoadDataIntoFile 153, App.Path & "\tmpFade"     'fade out
   LoadDataIntoFile 154, App.Path & "\tmpInvFade"  'Inverted Fade out

   frmLoader.ShowLoading MsgLits.LoadEnv, 150 'Loading Environment Variables...
   Sleep 100
   DoEvents

   LoadDataIntoFile 155, App.Path & "\tmpSetup"     'Setup
   LoadDataIntoFile 156, App.Path & "\tmpInvSetup"  'Inverted Setup

'   LoadDataIntoFile 157, App.Path & "\tmpCLS"     'Clear screen
'   LoadDataIntoFile 158, App.Path & "\tmpInvCLS"  'Inverted Clear Screen
   WriteLog "frmPlayer : Load Images into controls ..."
   cmdPause.Picture = LoadPicture(App.Path & "\tmpPause")
   cmdFadeOut.Picture = LoadPicture(App.Path & "\tmpFade")
   cmdSettings.Picture = LoadPicture(App.Path & "\tmpSetup")
'   cmdClearPalette.Picture = LoadPicture(App.Path & "\tmpCLS")

   frmLoader.ShowLoading MsgLits.GetReg, 150  'Retreiving Registry entries...
   Sleep 250
   DoEvents
   errCnt = errCnt + 1 '4
   WriteLog "frmPlayer : Get Registry entries : Max Buttons"
   'Set Max Buttons variables
   iButtonMaxSelected = Val(GetSetting(regMainKey, regSubKey, "Max Buttons"))
   If iButtonMaxSelected = 0 Then iButtonMaxSelected = 2  '16 buttons is the default...
   ScreenOptions(ButMaxSel) = iButtonMaxSelected
   Select Case iButtonMaxSelected
      Case 1
         iMaxBut = 9
      Case 2
         iMaxBut = 16 'DEFAULT !!!
      Case 3
         iMaxBut = 20
      Case 4
         iMaxBut = 30
   End Select


   'SetMaxButtonLayout
   'Set the Auto Adjust Volume variabe
   WriteLog "frmPlayer : Get Registry entries : AdjustVolume"
   iAdjustVol = Val(GetSetting(regMainKey, regSubKey, "AdjustVolume"))
   If iAdjustVol = 0 Then iAdjustVol = 1  'Leave unchanged is default
   ScreenOptions(ButAdjVol) = iAdjustVol
   'Set the Button direction
   WriteLog "frmPlayer : Get Registry entries : ButtonDirection"
   iButtonDirection = Val(GetSetting(regMainKey, regSubKey, "ButtonDirection"))
   If iButtonDirection = 0 Then iButtonDirection = 1  'Left to right is default
   ScreenOptions(ButDirection) = iButtonDirection
   'SetCorrectLayoutButton iButtonDirection

   'Set the Default Pause/Advance next button
   WriteLog "frmPlayer : Get Registry entries : PlayStopPause"
   iButtonPlayStopPause = Val(GetSetting(regMainKey, regSubKey, "PlayStopPause"))
   If iButtonPlayStopPause = 0 Then iButtonPlayStopPause = 1  'Stop after each song is default
   ScreenOptions(ButPlaystop) = iButtonPlayStopPause
   'Set AutoAdvance default
   WriteLog "frmPlayer : Get Registry entries : AutoAdvance"
   iAutoAdvance = Val(GetSetting(regMainKey, regSubKey, "AutoAdvance"))
   If iAutoAdvance = 0 Then iAutoAdvance = 1  'False is default
   ScreenOptions(ButAutoAdvance) = iAutoAdvance
   'SetAutoAdvance iAutoAdvance

   'RemoveSilence default
   WriteLog "frmPlayer : Get Registry entries : Remove Silence"
   iButtonRemoveSilence = Val(GetSetting(regMainKey, regSubKey, "Remove Silence"))
   If iButtonRemoveSilence = 0 Then iButtonRemoveSilence = 1  '
   ScreenOptions(ButRemSilence) = iButtonRemoveSilence
   'SetAutoRemoveSilences iButtonRemoveSilence

   'Streams default
   WriteLog "frmPlayer : Get Registry entries : Streams"
   iButtonStreams = Val(GetSetting(regMainKey, regSubKey, "Streams"))
   If iButtonStreams = 0 Then iButtonStreams = 2   '2 is the default
   ScreenOptions(ButStreams) = iButtonStreams
   'SetButtonStreams iButtonStreams

   'Button Default Colours
   WriteLog "frmPlayer : Get Registry entries : ButtonColor"
   iButtonDefaultColor = Val(GetSetting(regMainKey, regSubKey, "ButtonColor"))
   If iButtonDefaultColor = 0 Then iButtonDefaultColor = 1   '1 is the default
   ScreenOptions(ButDefColor) = iButtonDefaultColor
   'Button Default Secure mode
   WriteLog "frmPlayer : Get Registry entries : SecureMode"
   iSecureMode = Val(GetSetting(regMainKey, regSubKey, "SecureMode"))
   If iSecureMode = 0 Then iSecureMode = 1   '1 is the default (OFF)
   sSecurePWD = GetSetting(regMainKey, regSubKey, "SecureModePWD")
   ScreenOptions(ButSecure) = iSecureMode
   
   'Get the default colors
   WriteLog "frmPlayer : Get Registry entries : ButtonDefColor"
   vbNDefault = Val(GetSetting(regMainKey, regSubKey, "ButtonDefColor"))
   vbNDefaultFore = Val(GetSetting(regMainKey, regSubKey, "ButtonDefForColor"))
   
   If vbNDefault = 0 Then
    vbNDefault = &HE2FFB3
    vbNDefaultFore = &H80000008
  End If
   
   errCnt = errCnt + 1 '5

'''   Dim Ret As Long
'''   IsWow64Process GetCurrentProcess, Ret
'''   If Ret = 0 Then
'''      MsgBox "This application is not running on an x86 emulator for a 64-bit computer!"
'''   Else
'''      Dim SysInfo64 As SYSTEM_INFO
'''      GetNativeSystemInfo SysInfo64
'''      MsgBox "Number of processors on your 64-bit system: " + CStr(SysInfo64.dwNumberOrfProcessors)
'''   End If

   WriteLog "frmPlayer : Set Help file ..."
   App.HelpFile = App.Path & "\BackTrax.chm"

   frmLoader.ShowLoading MsgLits.FormatLayout, 150 'Format layouts
   Sleep 100
   DoEvents

   errCnt = errCnt + 1 '6
   HelpContextID = hlpIntro
   WriteLog "frmPlayer : FormatScreen ..."
   FormatScreen
   
   frmLoader.ShowLoading MsgLits.FormatLayout, 150 'Format layouts
   Sleep 100
   DoEvents

   errCnt = errCnt + 1 '7
   WriteLog "frmPlayer : LoadButtonColors ..."
   LoadButtonColors


   frmLoader.ShowLoading MsgLits.InitSoundCard, 150 'Load Sound Cards Devices...
   Sleep 100
   DoEvents


   
   errCnt = errCnt + 1 '8
   WriteLog "frmPlayer : GetDefaultSoundDevice ..."
   lDeviceNo = -99
   lDeviceNo = GetDefaultSoundDevice

   If lDeviceNo = -99 Then
     MsgBox "No device was loaded previously. Using default device", vbInformation, "No Device found"
     lDeviceNo = -1
   End If
   
   
   
   

   errCnt = errCnt + 1 '9
   WriteLog "frmPlayer : Call BASS_SetDevice(lDeviceNo) ..."
   Call BASS_SetDevice(lDeviceNo)  ' set the device to create stream on
  
   WriteLog "frmPlayer : BASS_Init ..."
   If BASS_Init(lDeviceNo, 44100, BASS_DEVICE_LATENCY, frmPlayer.hWnd, 0) = BASSFALSE Then
      ' MsgBox "Can't initialize device...", vbExclamation, Me.Caption
   End If

   errCnt = errCnt + 1 '10
   'Set the buffer size...
   WriteLog "frmPlayer : buflen = BASS_GetConfig(BASS_CONFIG_BUFFER) ..."
   buflen = BASS_GetConfig(BASS_CONFIG_BUFFER)

   'Dim bi As BASS_INFO
   WriteLog "frmPlayer : Buffer length = " & buflen
   If buflen < 2000 Then
      WriteLog "frmPlayer : Set Buffer length to 2500 ..."
      Call BASS_SetConfig(BASS_CONFIG_BUFFER, 2500)    'Make buffer 25 times as large... 5000 = maximum
      'Call BASS_SetConfig(BASS_CONFIG_BUFFER, buflen * 10)    'Make buffer 25 times as large... 5000 = maximum
   End If
'''  ' enable floating-point DSP
'''  Call BASS_SetConfig(BASS_CONFIG_FLOATDSP, BASSTRUE)
'''
'''  ' check for floating-point capability
'''  floatable = BASS_StreamCreate(44100, 2, BASS_SAMPLE_FLOAT, 0, 0)
'''  If (floatable) Then
'''      Call BASS_StreamFree(floatable)  ' woohoo!
'''      floatable = BASS_SAMPLE_FLOAT
'''  End If
    

   errCnt = errCnt + 1 '11
   WriteLog "frmPlayer : Check Theme ..."
  ' If IsThemeActive() = False Then
  '    ApplyStandardTheme = True
  ' Else
      ApplyStandardTheme = False
      'MsgBox "Windows Classic Theme detected"
  ' End If

   errCnt = errCnt + 1 '12
   WriteLog "frmPlayer : Set System Icon ..."
   SetIcon Me.hWnd, "AAA", False

   PanMain.Left = -10
   If ApplyStandardTheme Then
      PanMain.Top = 1300
   Else
      PanMain.Top = 1260
   End If





   errCnt = errCnt + 1 '13
   frmLoader.ShowLoading MsgLits.OpenPlayList, 150 'Loading Last Opened Playlist...
   Sleep 100
   DoEvents
   'Set the button layout according to master button
   WriteLog "frmPlayer : SetButtonsLayout ..."
   iPageno = 1  'Set to default page = 1
   SetButtonsLayout


   errCnt = errCnt + 1 '14
   'Make sure all the values are cleared
   WriteLog "frmPlayer : ClearButtons ..."
   ClearButtons



   'Show the form without the buttons yet visible
   errCnt = errCnt + 1 '15
   Dim lLeft As Long
   Dim lTop As Long
   lLeft = Val(GetSetting(regMainKey, regSubKey, "MoveLeft"))
   lTop = Val(GetSetting(regMainKey, regSubKey, "MoveTop"))

   If MonCount = 1 Then
      If lLeft < 0 Or lLeft > (Screen.Width - 300) Then
         lLeft = 0
         lTop = 0
      End If
      If lTop < 0 Or lTop > (Screen.Height - 200) Then
         lLeft = 0
         lTop = 0
      End If
   End If

   Me.Left = lLeft    'Val(GetSetting(regMainKey, regSubKey, "MoveLeft"))
   Me.Top = lTop      'Val(GetSetting(regMainKey, regSubKey, "MoveTop"))


   WriteLog "frmPlayer : Show the form without the buttons yet visible"
   PanMain.Height = 9800 - 60
   If iButtonMaxSelected = 3 Then
      PanMain.Width = Me.Width + 115
   Else
      PanMain.Width = Me.Width + 100
   End If
   
   
   'Now set the Page buttons and soundcard positions
   sspPageMain.Left = cmdSong(1).Left
   
   'If iButtonDirection = 2 Then
    sspPageMain.Top = 8490  'cmdSong(PageButNo).Top + cmdSong(PageButNo).Height + 90
   'Else
   ' sspPageMain.Top = cmdSong(4).Top + cmdSong(4).Height + 90
   'End If
   
'''   sspButMax.Left = sspPageMain.Left + sspPageMain.Width + 200
'''   sspButMax.Top = sspPageMain.Top
   
   sspSndHead.Left = (sspPageMain.Left + sspPageMain.Width) + 300
   sspDevice.Left = sspSndHead.Left + sspSndHead.Width + 90    '(cmdSong(16).Left + cmdSong(16).Width) - sspDevice.Width - 60   'cmdPage(6).Left + cmdPage(6).Width + 90 '
   sspDevice.Top = sspPageMain.Top
   sspSndHead.Top = sspPageMain.Top
   
   
   Me.Show
   DoEvents
   
   PanMain.Height = 9800 - 60
   If iButtonMaxSelected = 3 Then
      PanMain.Width = Me.Width + 115
   Else
      PanMain.Width = Me.Width + 100
   End If
   

   



   cmdSong(0).Left = 20000
   cmdReset.Left = 20000
   DoEvents
   
   frmLoader.Show
   DoEvents
   Sleep 100
   



'''
'''   errCnt = errCnt + 1 '16
'''   On Error Resume Next
'''   WriteLog "frmPlayer : Create Palets directory ..."
'''   MkDir App.Path & "\Palets"
'''   On Error GoTo 0
'''   On Error GoTo err1
   errCnt = errCnt + 1 '17
   WriteLog "frmPlayer : LoadDefaultPalette ..."
   LoadDefaultPalette
   cmdPage(1).BackColor = vbSelected
   cmdPage(1).ForeColor = vbBlack


'   'Show the form without the buttons yet visible
'   Me.Left = Val(GetSetting(regMainKey, regSubKey, "MoveLeft"))
'   Me.Top = Val(GetSetting(regMainKey, regSubKey, "MoveTop"))
'
'   Me.Show
'   DoEvents
'   frmLoader.Show
'   DoEvents


   frmLoader.ShowLoading MsgLits.Finalise, 150 'Finalising...
   Sleep 100
   DoEvents


'''   Change_pb_Color cpvVol.hwnd, vbBlue    'vbNCompleted
'''   Change_pb_ForeColor cpvVol.hwnd, vbRed


  ' picChanL.BorderStyle = 0
  ' picChanL.Appearance = prgFlat
'
'   picChanR.BorderStyle = 0
'   picChanR.Appearance = prgFlat


''   picChan1(0).Width = 0
''   picChan1(1).Width = 0
''   picChan2(0).Width = 0
''   picChan2(1).Width = 0

   errCnt = errCnt + 1 '18
   WriteLog "frmPlayer : SetupMainVolumeLights ..."
   SetupMainVolumeLights

   frmLoader.ShowLoading MsgLits.Finalise, 150 'Finalising...
   Sleep 100
   DoEvents

  ' pgtmp(0).ZOrder 0

'   Load pgtmp(0)
'   Load pgtmp(1)

''   pgtmp(0).Visible = True
''   pgtmp(0).value = 30473
''   pgtmp(1).Visible = True
''   pgtmp(1).value = 30473
''
''   Change_pb_ForeColor pgtmp(0).hwnd, vbRed    '&H404000
''   Change_pb_Color pgtmp(0).hwnd, vbBlack
''   Change_pb_ForeColor pgtmp(1).hwnd, vbRed   '&HFFAE27
''   Change_pb_Color pgtmp(1).hwnd, vbBlack
''
''   pgtmp(0).Left = pgLeft.Left
''   pgtmp(1).Left = pgRight.Left

'   pgtmp(0).ZOrder 0
'   pgtmp(1).ZOrder 0



   errCnt = errCnt + 1 '19
   sspTime.Caption = Format(Now, "HH:MM")
   lblDate.Caption = Format(Now, "DD MMMM YYYY")

   ButLeft = 22000
   sspVol.Left = ButLeft
   frmLoader.ShowLoading MsgLits.Finalise, 150 'Finalising...
   Sleep 100
   DoEvents
   WriteLog "frmPlayer : SetMainPeakLevel 0,0  ..."
   SetMainPeakLevel 0, 0

   errCnt = errCnt + 1 '20
   WriteLog "frmPlayer : Check DEMO Flag ..."

   SSPanel9.Left = 0
   SSPanel9.Top = 180

   If DemoFlag Then
      SSPanel9.Caption = "DEMO MODE"
      SSPanel9.MarqueeStyle = ssBlinkingMarquee
      sspVersion(3).Caption = ""
   Else
      WriteLog "frmPlayer : Get SerialNumber ..."
      SerialNumber = GetSetting(regMainKey, regSubKey, "SerialNumber")
      sspVersion(3).Caption = SerialNumber
      SSPanel9.Caption = ""
      SSPanel9.MarqueeStyle = ssNoneMarquee
   End If


'''   fx(0) = BASS_ChannelSetFX(chan, BASS_FX_DX8_PARAMEQ, 0) ' bass
'''   fx(1) = BASS_ChannelSetFX(chan, BASS_FX_DX8_PARAMEQ, 0) ' mid
'''   fx(2) = BASS_ChannelSetFX(chan, BASS_FX_DX8_PARAMEQ, 0) ' treble
'''   fx(3) = BASS_ChannelSetFX(chan, BASS_FX_DX8_REVERB, 0)  ' reverb
'''
'''   p.fGain = 0
'''   p.fBandwidth = 18
'''
'''   p.fCenter = 125                     ' bass   [125hz]
'''   Call BASS_FXSetParameters(fx(0), p)
'''
'''   p.fCenter = 1000                    ' mid    [1khz]
'''   Call BASS_FXSetParameters(fx(1), p)
'''
'''   p.fCenter = 8000                    ' treble [8khz]
'''   Call BASS_FXSetParameters(fx(2), p)
'''
'''   ' you can add more EQ bands with changing:
'''   ' p.fCenter = N [Hz] N>=80 and N<=16000
'''
'''   Call UpdateFX(0) ' bass
'''   Call UpdateFX(1) ' mid
'''   Call UpdateFX(2) ' treble
'''   Call UpdateFX(3) ' reverb

'''   'Clear the EQ Handle Arrays
'''   fxBass(1) = 0
'''   fxMid(1) = 0
'''   fxHigh(1) = 0
'''
'''   fxBass(2) = 0
'''   fxMid(2) = 0
'''   fxHigh(2) = 0

   'Load the last palette used

  ' Me.Caption = UCase(PaletteName)
   errCnt = errCnt + 1 '21
   WriteLog "frmPlayer : ShowButtonPlayArea ..."
   ShowButtonPlayArea False
   frmLoader.ShowLoading MsgLits.Finalise, 150 'Finalising...
   Sleep 100
   DoEvents

    
   
   WriteLog "frmPlayer : SetSecureMode ..."
   SetSecureMode

   DragingButton = False
  
   bLoading = False
   WriteLog "frmPlayer : Load Complete ..."
   WriteLog "**************************************************************************************"
   WriteLog " "
  ' Me.Show
   DoEvents
   Sleep 1000
   DoEvents
   frmLoader.Hide
   Screen.MousePointer = vbDefault
   DoEvents
 '  MsgBox "END Form Load"
Exit Sub

err1:
WriteLog "frmPlayer : -------------------------------------------------------------------------------------------------------------------------"
WriteLog "frmPlayer : An ERROR has Occurred ..."
WriteLog "frmPlayer : " & "ERROR " & Chr(13) & Chr(13) & "The following error has occurred at error count : " & errCnt & Chr(13) & Chr(13) & "ERROR : " & Err.Description & " (" & Err.Number & ")"
'MsgBox "Error in Module : Form_Load " & Chr(13) & Chr(13) & Err.Description, vbExclamation
MsgBox "ERROR " & Chr(13) & Chr(13) & "The following error has occurred at error count : " & errCnt & Chr(13) & Chr(13) & "ERROR : " & Err.Description & " (" & Err.Number & ")", vbExclamation, "frmPlayer_LOAD"
WriteLog "frmPlayer : -------------------------------------------------------------------------------------------------------------------------      "

End Sub

Sub LoadStats()
On Error GoTo err1
lstSystem.Clear

lstSystem.AddItem "Space : " & DiskSpaceTot & " GB"
lstSystem.AddItem "Free  : " & DiskSpaceFree & " GB"

lstSystem.AddItem "Direction  :  " & IIf(iButtonDirection = 1, "Top to Bottom", "Left To Right")
lstSystem.AddItem "Player Behaviour  :  " & IIf(iAutoAdvance = 2, "Auto Advance", "Play and Stop")
lstSystem.AddItem "Streams  :  " & iButtonStreams

Exit Sub
'lstSystem.AddItem "Buttons  :  " & iMaxBut
err1:
MsgBox Err.Description, vbExclamation, "LoadStats"
End Sub

Sub SetupMainVolumeLights()
Dim X As Integer

On Error GoTo err1

'''Exit Sub
'''
''''Set the Left and Width of the Peak Lights
'''For X = 0 To 3
'''   lblPeakML(X).Left = 45  '75
'''   lblPeakML(X).Width = 105
'''   lblPeakMR(X).Left = lblPeakML(X).Left + 125  '170  '210
'''   lblPeakMR(X).Width = 105
'''   lblPeakML(X).BackColor = &H5C7B0D
'''   lblPeakMR(X).BackColor = &H5C7B0D
'''Next X
''''On lights
'''lblOn(0).Left = lblPeakML(0).Left + 5
'''lblOn(1).Left = lblPeakMR(0).Left
'''lblOn(0).Width = lblPeakML(0).Width - 5
'''lblOn(1).Width = lblPeakMR(0).Width - 5
'''lblOn(0).Height = 165
'''lblOn(1).Height = lblOn(0).Height
''''Progress Bars
'''pgLeft.Left = 30    '60
'''pgRight.Left = pgLeft.Left + 115   '145  '195
'''pgLeft.Width = 135
'''pgRight.Width = 135
'''
''''Set the TOPS of each
'''lblPeakML(3).Top = 45
'''lblPeakML(2).Top = lblPeakML(3).Top + 90
'''lblPeakML(1).Top = lblPeakML(2).Top + 90
'''lblPeakML(0).Top = lblPeakML(1).Top + 90
'''pgLeft.Top = lblPeakML(0).Top + 60
'''For X = 0 To 3
'''   lblPeakMR(X).Top = lblPeakML(X).Top
'''Next X
'''lblOn(0).Top = pgLeft.Top + pgLeft.Height + 15
'''lblOn(1).Top = lblOn(0).Top


'sspVolumeTot.Width = 345

sspVolumeTot.BackColor = vbBlack

'Shape1.Left = 0
'Shape1.Top = 15
'Shape1.Height = 1225
'Shape1.Width = 320
'Shape1.BorderColor = &HE7DB49
'Shape1.ZOrder 0

'Change the progress bar's colors...
InitialisePeaks 0



''''Change_pb_ForeColor picLevelL(0).hWnd, vbGreen
''''Change_pb_Color picLevelL(0).hWnd, vbBlack
''''Change_pb_ForeColor picLevelR(0).hWnd, vbGreen
''''Change_pb_Color picLevelR(0).hWnd, vbBlack

Exit Sub
err1:

MsgBox "Error in Module : SetupMainVolumeLights " & Chr(13) & Chr(13) & Err.Description, vbExclamation
   

End Sub

Sub LoadButtonColors()

On Error GoTo err1

ButColors(0) = vbNDefault
ButForColors(0) = vbNDefaultFore
ButColors(1) = vbNColor1
ButForColors(1) = vbWhite 'vbBlack
ButColors(2) = vbNColor2
ButForColors(2) = vbWhite 'vbBlack
ButColors(3) = vbNColor3
ButForColors(3) = vbWhite 'vbCyan
ButColors(4) = vbNColor4
ButForColors(4) = vbWhite 'vbCyan
ButColors(5) = vbNColor5
ButForColors(5) = vbWhite 'vbGreen
ButColors(6) = vbNColor6
ButForColors(6) = vbWhite 'vbBlack
ButColors(7) = vbNColor7
ButForColors(7) = vbWhite 'vbBlack
ButColors(8) = vbNColor8
ButForColors(8) = vbWhite 'vbCyan
ButColors(9) = vbNColor9
ButForColors(9) = vbWhite 'vbGreen
ButColors(10) = vbNColor10
ButForColors(10) = vbWhite 'vbGreen
ButColors(11) = vbNColor11
ButForColors(11) = vbWhite 'vbBlack
ButColors(12) = vbNColor12
ButForColors(12) = vbWhite 'vbBlack
ButColors(13) = vbNColor13
ButForColors(13) = vbWhite 'vbGreen
ButColors(14) = vbNColor14
ButForColors(14) = vbWhite 'vbBlack
ButColors(15) = vbNColor15
ButForColors(15) = vbWhite 'vbBlack
ButColors(16) = vbNColor16
ButForColors(16) = vbWhite 'vbBlack

Exit Sub
err1:

MsgBox "Error in Module : LoadButtonColors " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Function GetDefaultSoundDevice() As Long
Dim sDevice As String
   On Error GoTo err1
   
   GetDefaultSoundDevice = -1
   sDevice = GetSetting(regMainKey, regSubKey, "Current Device")
   If sDevice <> "" Then
      lDeviceNo = CLng(sDevice)
   End If

   sspDevice.Caption = Trim(GetSetting(regMainKey, regSubKey, "Current Device Description"))
   sspDevice.Tag = CStr(lDeviceNo)
   
   If lDeviceNo <> -99 Then GetDefaultSoundDevice = lDeviceNo
   
   Exit Function
   
err1:
   MsgBox "Error in Module : GetDefaultSoundDevice " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Function

Sub LoadDefaultPalette()

   On Error GoTo err1
   
   PaletteName = UCase(GetSetting(regMainKey, regSubKey, "Palette Name"))
   If Trim(PaletteName) <> "" Then
     LoadPalette PaletteName, 1, 1   '1=ALL
     Me.Caption = PaletteName
     lblPaleteName.Caption = Me.Caption
   End If
   
   Exit Sub
   
err1:
   MsgBox "Error in Module : LoadDefaultPalette " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   On Error Resume Next
   
''''''   Dim c As Long
''''''   Dim I As BASS_DEVICEINFO
''''''   Dim iDevCnt As Long
''''''
''''''   c = 1      ' device 1 = 1st real device
''''''   While BASS_GetDeviceInfo(c, I)
''''''     If (I.Flags And BASS_DEVICE_ENABLED) Then
''''''       c = c + 1
''''''     End If
''''''   Wend
''''''
''''''   For iDevCnt = 1 To c - 1
''''''     Call BASS_SetDevice(iDevCnt)
''''''     Call BASS_Free
''''''   Next iDevCnt
''''''

'  Call BASS_PluginFree(0)
'  Call BASS_Free

   
   Kill App.Path & "\tmpClear" 'Clear
   Kill App.Path & "\tmpOpen"  'Open
   Kill App.Path & "\tmpColor" 'Color
   Kill App.Path & "\tmpSpkr"  'Color
   Kill App.Path & "\tmpDrive" 'Drive
   Kill App.Path & "\tmpclose" 'Drive
   Kill App.Path & "\tmpComplete" 'Completed
   Kill App.Path & "\tmpPlus"     'Plus
   Kill App.Path & "\tmpBanner"   'Banner
   
   Kill App.Path & "\tmpPause"    'Pause
   Kill App.Path & "\tmpInvPause" 'Inverted Pause
   Kill App.Path & "\tmpFade"     'fade out
   Kill App.Path & "\tmpInvFade"  'Inverted Fade out
   
   Kill App.Path & "\tmpSetup"    'Setup
   Kill App.Path & "\tmpInvSetup" 'Inverted Setup
''   Kill App.Path & "\tmpCLS"      'CLS
''   Kill App.Path & "\tmpInvtmpCLS"  'Inverted CLS
        
   SaveSetting regMainKey, regSubKey, "AdjustVolume", iAdjustVol
   SaveSetting regMainKey, regSubKey, "ButtonDirection", iButtonDirection
   SaveSetting regMainKey, regSubKey, "Max Buttons", iButtonMaxSelected
   SaveSetting regMainKey, regSubKey, "Remove Silence", iButtonRemoveSilence
   SaveSetting regMainKey, regSubKey, "Streams", iButtonStreams
   SaveSetting regMainKey, regSubKey, "PlayStopPause", iButtonPlayStopPause
   SaveSetting regMainKey, regSubKey, "AutoAdvance", iAutoAdvance
   SaveSetting regMainKey, regSubKey, "ButtonColor", iButtonDefaultColor
   SaveSetting regMainKey, regSubKey, "SecureMode", iSecureMode
   SaveSetting regMainKey, regSubKey, "SecureModePWD", sSecurePWD
    
   SaveSetting regMainKey, regSubKey, "Palette Name", Me.Caption
   
   SaveSetting regMainKey, regSubKey, "MoveTop", IIf(Me.Top < 0, 0, Me.Top)
   SaveSetting regMainKey, regSubKey, "MoveLeft", IIf(Me.Left < 0, 0, Me.Left)
   
 '''  Call UnSubClass(Me.hwnd)
   
End Sub

Private Sub Form_Resize()
'MsgBox "With : " & Me.Width & Chr(13) & "Height : " & Me.Height
On Error GoTo err1

   If Me.WindowState <> 1 Then
     ' Me.Caption = ""
'      Me.Top = 0
'      Me.Left = 0
      If ApplyStandardTheme Then
         Me.Width = 20445 + 30 + 115
         Me.Height = 11010 - 60
      Else
         Me.Width = 20445 + 30 + 115
         Me.Height = 11010 - 60
      End If
   ElseIf Me.WindowState = 1 Then
      'Me.Caption = "BackTrax Player"
   End If
   
Exit Sub
err1:

MsgBox "Error in Module : Form_Resize " & Chr(13) & Chr(13) & Err.Description, vbExclamation


End Sub

Private Sub Form_Unload(Cancel As Integer)

   End

End Sub

Sub SetButtonColor(Index As Integer, ColorIndex As Integer)

   On Error GoTo err1
   
 '  Set sspSongTitle(Index).Picture = Nothing
   
  ' sspProgress(Index).FloodColor = vbYellow
   lblStatus(Index).ForeColor = ButForColors(ColorIndex)  'vbWhite
   lblTimePlayed(Index).ForeColor = ButForColors(ColorIndex) 'vbWhite
   lblTimeLeft(Index).ForeColor = ButForColors(ColorIndex)           'vbBlack
   'Store the Button color variable for later use/retrieval
   
   sspProgress(Index).Tag = ColorIndex
   
  ' cpvVolume(Index).BackColor = ButColors(ColorIndex)
  ''' sspProgress(Index).BackColor = &HC0C0C0   ' vbBlack

   sspSongTitle(Index).ForeColor = ButForColors(ColorIndex)
   lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
  ' lblButtonCnt(Index).ForeColor = vbWhite
   sspSongTitle(Index).Font.Bold = True
   sspSongTitle(Index).BackColor = ButColors(ColorIndex)
   cmdSong(Index).ForeColor = ButForColors(ColorIndex)
   cmdSong(Index).BackColor = ButColors(ColorIndex)
   
   imgCompleted0(Index).Visible = False
   imgCompleted1(Index).Visible = False
   imgCompleted2(Index).Visible = False
   imgCompleted3(Index).Visible = False
      
'   sspSongTitle(Index).BevelOuter = ssInsetBevel    'ssNoneBevel
'   sspSongTitle(Index).BevelWidth = 1
Exit Sub
err1:

MsgBox "Error in Module : SetButtonColor " & Chr(13) & Chr(13) & Err.Description, vbExclamation


End Sub

Sub FormatScreen()
Dim i As Integer

'For i = 1 To 4
'   sspOption3(i).ForeColor = vbDirectionColor
'Next i

cmdPause.ForeColor = vbDirectionColor
cmdFadeOut.ForeColor = vbDirectionColor
cmdLoadPalette.ForeColor = vbDirectionColor
cmdSavePalette.ForeColor = vbDirectionColor
cmdClearPalette.ForeColor = vbDirectionColor
cmdSettings.ForeColor = vbDirectionColor
cmdExit.ForeColor = vbDirectionColor




End Sub

Sub ResetButton(Index As Integer, Optional Stopped As Boolean)
   Dim iColor As Integer
   On Error GoTo err1
   
   iColor = Val(sspProgress(Index).Tag) 'VAL()   Will also force to 0 if nothing setup
   SetButtonColor Index, iColor
    
   lblStatus(Index).Caption = "Ready"
   
   lblStream(Index).Caption = ""
   lblStream(Index).Tag = ""
   lblStream(Index).Visible = False
   
   imgVol(Index).Visible = sspVol.Left = 22000 'True
'   If imgVol(Index).Visible = False And sspVol.Left = 22000 Then Closevolume
   
'  If bDoEq Then imgEQ(Index).Visible = True
   
   imgDirection(Index).Visible = True
      
   imgCompleted0(Index).Visible = False
   imgCompleted1(Index).Visible = False
   imgCompleted2(Index).Visible = False
   imgCompleted3(Index).Visible = False

   sspSongTitle(Index).Width = sspSongTitle(0).Width   '- 50
 '  sspSongTitle(Index).Font3D = ssNoneFont3D
   
'   Set sspSongTitle(Index).Picture = Nothing
'   sspSongTitle(Index).BevelInner = ssNoneBevel
'   sspSongTitle(Index).BevelOuter = ssInsetBevel
'   sspSongTitle(Index).BorderWidth = 1
'   sspSongTitle(Index).BevelWidth = 1
   sspSongTitle(Index).BackStyle = 0
'   sspSongTitle(Index).Outline = False
   sspSongTitle(Index).ForeColor = ButForColors(iColor)  'vbWhite
   lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
   
   
'   lblTimeLeft(Index).Caption = "0:00"  'Trim((time \ 60) & ":" & Format(time Mod 60, "00"))
'   lblTimePlayed(Index).Tag = Trim((time \ 60) & ":" & Format(time Mod 60, "00"))
   lblTimeLeft(Index).Caption = "00:00"
   lblTimePlayed(Index).Caption = lblTimePlayed(Index).Tag
   
'   lblTimePlayed(Index).Caption = "0:00"
   lblTimePlayed(Index).Top = lblTimePlayed(0).Top
   lblTimePlayed(Index).Height = lblTimePlayed(0).Height
   lblTimePlayed(Index).FontSize = lblTimePlayed(0).FontSize
   lblTimePlayed(Index).FontBold = False
   
   sspSongTitle(Index).Visible = True
   lblTimePlayed(Index).Visible = True
   lblTimeLeft(Index).Visible = True
   
   
   sspProgress(Index).Visible = True
   sspProgress(Index).value = 0
   
'''   imgSetup(Index).Visible = True
   
  ' sspProgress(Index).FloodPercent = 0
  ' sspProgress(Index).BevelOuter = ssNoneBevel
  ' sspProgress(Index).BevelInner = ssNoneBevel
  ' sspProgress(Index).Outline = True
   
''''   picLevelL(Index).value = 0
''''   picLevelR(Index).value = 0
''''   sspLevel(Index).Visible = False
   
   If Val(cmdSong(Index).Tag) <> 0 Then
     If Val(cmdSong(Index).Tag) = 1 Then
       TimerP1.Enabled = False
       TimerF1.Enabled = False
       TimerP1Level.Enabled = False
'''       picChan1(0).Width = 0
'''       picChan1(1).Width = 0
       Left1Chan = 0
       Right1Chan = 0
       sspStream1.Caption = "Ready..."
     Else
       TimerP2.Enabled = False
       TimerF2.Enabled = False
       TimerP2Level.Enabled = False
'''       picChan2(0).Width = 0
'''       picChan2(1).Width = 0
       Left2Chan = 0
       Right2Chan = 0
       sspStream2.Caption = "Ready..."
     End If
     cmdSong(Index).Tag = ""
   End If
   
   TimerMainLevel.Enabled = TimerP1Level.Enabled Or TimerP2Level.Enabled
   
   Change_pb_Color sspProgress(Index).hWnd, &H260F35   'vbNCompleted
   
   If Stopped Then
      cmdSong(Index).BackColor = vbNCompleted
      'cmdSong(Index).Picture = cmdReset.Picture   '   LoadResPicture(113, vbResBitmap)
      
      imgCompleted0(Index).Visible = True
      imgCompleted1(Index).Visible = True
      imgCompleted2(Index).Visible = True
      imgCompleted3(Index).Visible = True
      lblTimeLeft(Index).ForeColor = ButForColors(iColor)
      If iButtonDirection = 0 Then 'Top To Bottom
        cmdSong(Index).PictureAlignment = ssRightTop
      Else
        cmdSong(Index).PictureAlignment = ssRightMiddle
      End If
     ' cpvVolume(Index).BackColor = cmdSong(Index).BackColor
      sspSongTitle(Index).ForeColor = vbWhite
      lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
      GetTotalTime
   End If
   
   sspSongTitle(Index).BackColor = cmdSong(Index).BackColor
   cmdSong(Index).ForeColor = ButForColors(iColor)
   
Exit Sub
err1:

MsgBox "Error in Module : ResetButton " & Chr(13) & Chr(13) & Err.Description, vbExclamation

Resume Next

End Sub

Sub ResetButtons()
   On Error Resume Next
   
   For i = 1 To iMaxBut
     If lblStatus(i).Caption = "Ready" Then
       ResetButton i
       sspSongTitle(i).Caption = ""
       sspSongTitle(i).Tag = ""
     End If
   Next i
   
End Sub

Private Sub imgCompleted0_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
   
End Sub

Private Sub imgCompleted1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
End Sub

Private Sub imgCompleted2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
   
End Sub

Private Sub imgCompleted3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
End Sub

Private Sub imgDirection_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
DoEvents
End Sub

Private Sub imgDirection_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   
   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If

End Sub

Private Sub imgEQ_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If iPlayingChan = 0 Then Exit Sub
'Set the Forms Index property...
frmEq.SongIndex = Index
Load frmEq

frmEq.Show vbModal

End Sub

Private Sub imgSetup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'DoEvents
'Screen.MousePointer = 13  '14
'DoEvents
End Sub

Private Sub imgSetup_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo err1
   Dim iSongs As Integer
   Dim i As Integer
   
'   Timer3.Enabled = False
'   If LongPress > 40 Then
'      LongPress = 0
'      Exit Sub
'   End If
   
   If iSecureMode = 2 Then Exit Sub
   
   If Button = 2 Then 'Right-clicked
''''     iMnuFlag = 2
''''     Setstate Index
   Else
      If sspSongTitle(Index).Caption = "" Then
''''         '==================================================================================
''''         'For Demo system, only allow load of 5 songs
''''         If DemoFlag Then
''''            If DetermineTotButtons(True) >= DemoMax Then
''''               MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3, vbExclamation, DemoHeading
''''               Exit Sub
''''            End If
''''         End If
''''         '==================================================================================
''''
''''         iMnuFlag = 2 'Make sure we only load a new song when button is not initialised
''''         Setstate Index
''''
''''         palletArr(0) = ""
''''         If Trim(Me.Caption) = "" Then
''''            Me.Caption = "tmp001"
''''            palletArr(0) = Trim(Me.Caption)
''''         End If
''''         SavePalete Trim(Me.Caption)
         LoadNewSong Index
      Else
         bTagEditMP3 = UCase(Right(sspSongTitle(Index).Tag, 3)) = "MP3"
         'FilenameToLoad = sspSongTitle(Index).Tag
         If lblStatus(Index).Caption = "Ready" Then ShowOptionScreen Index
         
         SavePalete Trim(Me.Caption), iPageno
         LoadPalette Trim(Me.Caption), iPageno, 3
         
      End If
   End If
   
   
   Exit Sub
   
err1:

MsgBox "Error in Module : cmdSong_MouseUp " & Chr(13) & Chr(13) & Err.Description, vbExclamation
End Sub

Private Sub imgVol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   
   If lblStatus(Index).Caption <> "Playing" Then
     imgVol(Index).Picture = LoadResPicture(117, vbResIcon)
     DoEvents
   End If
   
End Sub


Private Sub imgVol_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DoEvents
Screen.MousePointer = 14  '14
DoEvents
End Sub

Private Sub imgVol_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo err1

   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
   
   
   'imgVol(Index).Picture = LoadResPicture(118, vbResIcon)    'LoadPicture(App.Path & "\tmpSpkr")
   imgVol(Index).Picture = LoadResPicture(133, vbResIcon)    'LoadPicture(App.Path & "\tmpSpkr")
   DoEvents
   
   'Set Global variable
   iVolIndex = Index
   lVolume = Val(lblVol(Index).Caption)
   'Shift screen to correct position
   If ButLeft <> Screen.ActiveForm.Left + cmdSong(Index).Left + imgVol(Index).Left + imgVol(Index).Width Then
      ShowVolScreen Index
      cpvVol.value = lVolume
      cpvVol.ZOrder 0
      DoEvents
   Else
      Closevolume
      cpvVol.ZOrder 1
   End If
''   'Set the variables
''   SetEqLabels Index
''
''   If lblStatus(Index).Caption = "Playing" Then
''      'Load the handles here
''      SetEqHandles Val(cmdSong(Index).Tag)
''      'Update EQ with current values in the labels
''      UpdateAllFX Val(cmdSong(Index).Tag)
''   End If




'''frmVolume.Show vbModal
'''
'''lblVol(Index).Caption = lVolume
'''lblBass(Index).Caption = lBass
'''lblMid(Index).Caption = lMid
'''lblHigh(Index).Caption = lHigh


'''frmTestMp3Tags.Show vbModal
'''
'''If bTagsUpdated Then
'''   Screen.MousePointer = vbHourglass
'''   Me.Enabled = False
'''   If FilenameToLoad = "" Then
'''     Screen.MousePointer = vbDefault
'''     Me.Enabled = True
'''     Exit Sub
'''   End If
'''   ClearButton Index
'''   'Reset button to Loaded state
'''   ResetButton Index
'''   'Add caption to button
'''   lblStatus(Index).Caption = "Loading..."
'''   ExtractTagInfo FilenameToLoad
'''   SetupButton Index, Id3TagArr(1), Id3TagArr(2), FilenameToLoad
'''End If

Exit Sub
err1:
MsgBox "Error in Module : imgVol_MouseUp " & Chr(13) & Chr(13) & Err.Description, vbExclamation
  
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Closevolume
End Sub

Private Sub lblSelect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next

   If lblStatus(Index).Caption <> "Playing" Then
     imgVol(Index).Picture = LoadResPicture(117, vbResIcon)
     DoEvents
   End If
End Sub

Private Sub lblSelect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err1
   'imgVol(Index).Picture = imgVolSource.Picture
   'imgVol(Index).Picture = LoadPicture(App.Path & "\tmpSpkr")
   
   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
   
   imgVol(Index).Picture = LoadResPicture(133, vbResIcon)
   'Set Global variable
   iVolIndex = Index
   lVolume = Val(lblVol(Index).Caption)
   'Shift screen to correct position
   If ButLeft <> Screen.ActiveForm.Left + cmdSong(Index).Left + imgVol(Index).Left + imgVol(Index).Width Then
      ShowVolScreen Index
      cpvVol.value = lVolume
   Else
      Closevolume
   End If

Exit Sub
err1:
MsgBox "Error in Module : lblSelect_MouseUp " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub lblStatus_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If Button = 2 Then
     'ShowPopMenu Index
   Else
     iMnuFlag = 0
     Setstate Index
   End If
End Sub

Private Sub lblStream_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If lblTimeLeft(Index).Caption = "" Then
      sspSongTitle(Index).TabIndex = 0
      If Me.lblStatus(Index).Caption = "Playing" Then Exit Sub
      
      LongPress = 0
      SelPlayerIndex = Index
      Timer3.Enabled = True
      Exit Sub
   End If
End Sub

Private Sub lblStream_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
DoEvents
End Sub

Private Sub lblStream_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If lblTimeLeft(Index).Caption = "" Then
      On Error Resume Next
      
      Timer3.Enabled = False
      If LongPress > LongPressCnt Then
         LongPress = 0
         Exit Sub
      End If
      
      If Button = 2 Then
         iMnuFlag = 2
         Setstate Index
      Else
         iMnuFlag = 0
         Setstate Index
      End If
      
      Command2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub lblTimeLeft_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
   On Error Resume Next
'''   If lblTimeLeft(Index).Caption = "" Then
'''      sspSongTitle(Index).TabIndex = 0
'''      If Me.lblStatus(Index).Caption = "Playing" Then Exit Sub
'''
'''      LongPress = 0
'''      SelPlayerIndex = Index
'''      Timer3.Enabled = True
'''      Exit Sub
'''   End If

   If lblStatus(Index).Caption <> "Playing" Then
     imgVol(Index).Picture = LoadResPicture(117, vbResIcon)
     DoEvents
   End If
   
   Exit Sub
   
   
err1:
      MsgBox "Error in Module : sspSongTitle_MouseDown " & Chr(13) & Chr(13) & Err.Description, vbExclamation
      Exit Sub
   
End Sub

Private Sub lblTimeLeft_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
DoEvents
End Sub

Private Sub lblTimeLeft_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo err1

   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
   
   'imgVol(Index).Picture = imgVolSource.Picture
  ' imgVol(Index).Picture = LoadPicture(App.Path & "\tmpSpkr")
   imgVol(Index).Picture = LoadResPicture(133, vbResIcon)
   DoEvents
   'Set Global variable
   
   If iVolIndex <> 0 And iVolIndex <> Index Then
    Closevolume
   End If
   
   iVolIndex = Index
   lVolume = Val(lblVol(Index).Caption)
   'Shift screen to correct position
   If ButLeft <> Screen.ActiveForm.Left + cmdSong(Index).Left + imgVol(Index).Left + imgVol(Index).Width Then
      ShowVolScreen Index
      cpvVol.value = lVolume
      cpvVol.ZOrder 0
   Else
      Closevolume
      cpvVol.ZOrder 1
   End If
   
Exit Sub
err1:
MsgBox "Error in Module : lblTimeLeft_MouseUp " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub lblTimePlayed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'''   On Error Resume Next
'''   If lblTimeLeft(Index).Caption = "" Then
'''      sspSongTitle(Index).TabIndex = 0
'''      If Me.lblStatus(Index).Caption = "Playing" Then Exit Sub
'''
'''      LongPress = 0
'''      SelPlayerIndex = Index
'''      Timer3.Enabled = True
'''      Exit Sub
'''   End If
End Sub

Private Sub lblTimePlayed_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
DoEvents
End Sub

Private Sub lblTimePlayed_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
End Sub

Private Sub lblVolInd_Click()
Closevolume
End Sub

Private Sub lblVolTxt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Closevolume
End Sub

Private Sub mnuPop_Click(Index As Integer)
   On Error Resume Next
   
   bAsigning = False
   iMnuFlag = Index
   If Index = 2 Then bAsigning = True

End Sub

''''Private Sub picProgress_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
''''   On Error Resume Next
''''   If Button <> 2 Then Exit Sub
''''
''''   If Me.lblStatus(Index).Caption <> "Playing" Then Exit Sub
''''
''''   SetSongPos chan(CLng(cmdSong(Index).Tag)), Index, CInt(X)
''''
''''End Sub

Sub SetSongPos(PlayStreamHndle As Long, Index As Integer, sPos As Integer)
   Dim iColor As Integer
   On Error GoTo err1
   
   Dim pos As Single
   Dim NewPerc As Single
   Dim NewStart As Integer
   
   NewPerc = (sPos / MaxWidth) * 100  'Percentage of where I need to start
   NewStart = ((NewPerc * Duration(CLng(cmdSong(Index).Tag))) / 100)
   iColor = Val(sspProgress(Index).Tag)
   
      sspSongTitle(Index).ForeColor = vbBlack
      lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
      'sspSongTitle(Index).BackStyle = 1
      sspSongTitle(Index).BackColor = vbNDYellow
'      sspSongTitle(Index).BevelOuter = ssInsetBevel
      
   If Duration(CLng(cmdSong(Index).Tag)) - NewStart > 10 Then
      If iColor = 0 Then
        lblTimePlayed(Index).ForeColor = vbYellow '''vbBlack
      Else
        lblTimePlayed(Index).ForeColor = vbYellow   'sspSongTitle(Index).ForeColor
      End If
      If CLng(cmdSong(Index).Tag) = 1 Then
         TimerF1.Enabled = False
      Else
         TimerF2.Enabled = False
      End If
   End If
      
   Call BASS_ChannelSetPosition(PlayStreamHndle, BASS_ChannelSeconds2Bytes(PlayStreamHndle, CDbl(NewStart)), BASS_POS_BYTE)  ' set the position
      
   
Exit Sub
err1:
MsgBox "Error in Module : SetSongPos " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Sub SetSongPosFromDucking(PlayStreamHndle As Long, Index As Integer, sPos As Integer)
   Dim iColor As Integer
   On Error GoTo err1
   
   Dim pos As Single
   Dim NewPerc As Single
   Dim NewStart As Integer
   
   NewPerc = ((sPos * 13) / MaxWidth) * 100 'Percentage of where I need to start
   NewStart = ((NewPerc * Duration(CLng(cmdSong(Index).Tag))) / 100)
   iColor = Val(sspProgress(Index).Tag)
   
      sspSongTitle(Index).ForeColor = vbBlack
      sspSongTitle(Index).BackStyle = 1
      sspSongTitle(Index).ForeColor = vbBlack
      lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
      sspSongTitle(Index).BackColor = vbNDYellow
'      sspSongTitle(Index).BevelOuter = ssInsetBevel
      
   If Duration(CLng(cmdSong(Index).Tag)) - NewStart > 10 Then
      lblTimePlayed(Index).ForeColor = vbYellow   'sspSongTitle(Index).ForeColor
      If CLng(cmdSong(Index).Tag) = 1 Then
         TimerF1.Enabled = False
      Else
         TimerF2.Enabled = False
      End If
   End If
      
   Call BASS_ChannelSetPosition(PlayStreamHndle, BASS_ChannelSeconds2Bytes(PlayStreamHndle, CDbl(NewStart)), BASS_POS_BYTE)   ' set the position
      
Exit Sub
err1:
MsgBox "Error in Module : SetSongPosFromDucking " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub


Sub MainLevelColor()

On Error GoTo err1

iBarColor = iBarColor + 1
If iBarColor > 5 Then iBarColor = 1

Select Case iBarColor
   Case 1
      Change_pb_ForeColor pgLeft.hWnd, vbCyan
      Change_pb_ForeColor pgRight.hWnd, vbCyan
   Case 2
      Change_pb_ForeColor pgLeft.hWnd, vbGreen
      Change_pb_ForeColor pgRight.hWnd, vbGreen
   Case 3
      Change_pb_ForeColor pgLeft.hWnd, vbOrange
      Change_pb_ForeColor pgRight.hWnd, vbOrange
   Case 4
      Change_pb_ForeColor pgLeft.hWnd, vbMagenta
      Change_pb_ForeColor pgRight.hWnd, vbMagenta
   Case 5
      Change_pb_ForeColor pgLeft.hWnd, vbBlue
      Change_pb_ForeColor pgRight.hWnd, vbBlue
End Select

Change_pb_Color pgLeft.hWnd, vbBlack
Change_pb_Color pgRight.hWnd, vbBlack

Exit Sub
err1:
MsgBox "Error in Module : MainLevelColor " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub optLayout_Click(Index As Integer, value As Integer)

On Error GoTo err1

If bLoading Then Exit Sub

If InStr(1, sspStream1.Caption, "Ready") = 0 Then Exit Sub
If InStr(1, sspStream2.Caption, "Ready") = 0 Then Exit Sub

iButtonDirection = Index
frmPlayer.Enabled = False
Screen.MousePointer = vbHourglass
DoEvents
RemoveButtons
SetButtonsLayout
ClearButtons
DoEvents
LoadPalette PaletteName, iPageno, 1

frmPlayer.Enabled = True
Screen.MousePointer = vbDefault
DoEvents

Exit Sub
err1:
MsgBox "Error in Module : optLayout_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Sub RemoveButtons()
Dim iBut As Integer
Dim MaxButs As Integer

On Error GoTo err1

MaxButs = cmdSong.Count - 1
'Skip 0 because this is our source...
For iBut = 1 To MaxButs
      Unload lblTimePlayed(iBut)
      Unload lblTimeLeft(iBut)
      Unload lblSelect(iBut)
      Unload lblStatus(iBut)
      Unload imgVol(iBut)
      
      Unload lblButtonCnt(iBut)
      
   ' If bDoEq Then Unload imgEQ(iBut)
      
      Unload imgCompleted0(iBut)
      Unload imgCompleted1(iBut)
      Unload imgCompleted2(iBut)
      Unload imgCompleted3(iBut)
      
      Unload sspProgress(iBut)
''''      Unload picLevelL(iBut)
''''      Unload picLevelR(iBut)
      Unload lblVol(iBut)
      
      Unload lblStream(iBut)
''''      Unload lblPeakL(iBut)
''''      Unload lblPeakR(iBut)
''''      Unload lblMidHL(iBut)
''''      Unload lblMidLL(iBut)
''''      Unload lblMidHR(iBut)
''''      Unload lblMidLR(iBut)
      Unload imgDirection(iBut)
'''      Unload imgSetup(iBut)
      Unload sspSongTitle(iBut)
''''      Unload sspLevel(iBut)
   ''''   Unload lnProgress(iBut)
      'All the sub-controls are removed, Remove button control
      Unload cmdSong(iBut)
      
     
Next iBut
   
Exit Sub
err1:
MsgBox "Error in Module : RemoveButtons " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub PanMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
End Sub

Private Sub pgRight_Click()
MainLevelColor
End Sub

Private Sub imgColor_Click(Index As Integer)
Dim iColor As Integer
On Error Resume Next
         
gColor = -1
 frmSelectButtonColor.Show vbModal

If gColor <> -1 Then
   SetButtonColor Index, gColor  'Set the color of the button accordingly...
   Exit Sub
End If

If bClearButton Then
   ClearButton (Index)
   Exit Sub
End If

If bLoadButton Then
   If lblStatus(Index).Caption <> "Playing" Then
     iMnuFlag = 2
     Setstate Index
   End If
End If

End Sub

''''
''''Private Sub Picture1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''''
'''''Code to set slider position when arrow keys are used
''''
''''    Dim t!
''''
''''    Select Case KeyCode
''''       'the end(35), home(36), PgUp(33) or PgDown(34) could also be used
''''       Case 37, 38, 39, 40
''''          Select Case KeyCode
''''             Case 38 'Up button
''''                If Slider(Index).sldOrient = 1 Then   '1-Horizontal
''''                   Slider(Index).sldCrntVal = Slider(Index).sldMaxVal
''''                Else                                  '2-Vertical
''''                   t = (Slider(Index).sldMaxVal - Slider(Index).sldMinVal) / Slider(Index).sldNumDiv
''''                   If Slider(Index).sldCrntVal + t > Slider(Index).sldMaxVal Then
''''                      Slider(Index).sldCrntVal = Slider(Index).sldMaxVal
''''                   Else
''''                      Slider(Index).sldCrntVal = Slider(Index).sldCrntVal + t
''''                   End If
''''                End If
''''             Case 40 'Down button
''''                If Slider(Index).sldOrient = 1 Then   '1-Horizontal
''''                   Slider(Index).sldCrntVal = Slider(Index).sldMinVal
''''                Else
''''                   t = (Slider(Index).sldMaxVal - Slider(Index).sldMinVal) / Slider(Index).sldNumDiv
''''                   If Slider(Index).sldCrntVal - t < Slider(Index).sldMinVal Then
''''                      Slider(Index).sldCrntVal = Slider(Index).sldMinVal
''''                   Else
''''                      Slider(Index).sldCrntVal = Slider(Index).sldCrntVal - t
''''                   End If
''''                End If
''''             Case 37 'Left button
''''                If Slider(Index).sldOrient = 1 Then   '1-Horizontal
''''                   t = (Slider(Index).sldMaxVal - Slider(Index).sldMinVal) / Slider(Index).sldNumDiv
''''                   If Slider(Index).sldCrntVal - t < Slider(Index).sldMinVal Then
''''                      Slider(Index).sldCrntVal = Slider(Index).sldMinVal
''''                   Else
''''                      Slider(Index).sldCrntVal = Slider(Index).sldCrntVal - t
''''                   End If
''''                Else
''''                   Slider(Index).sldCrntVal = Slider(Index).sldMinVal
''''                End If
''''             Case 39 'Right button
''''                If Slider(Index).sldOrient = 1 Then   '1-Horizontal
''''                   t = (Slider(Index).sldMaxVal - Slider(Index).sldMinVal) / Slider(Index).sldNumDiv
''''                   If Slider(Index).sldCrntVal + t > Slider(Index).sldMaxVal Then
''''                      Slider(Index).sldCrntVal = Slider(Index).sldMaxVal
''''                   Else
''''                      Slider(Index).sldCrntVal = Slider(Index).sldCrntVal + t
''''                   End If
''''                Else
''''                   Slider(Index).sldCrntVal = Slider(Index).sldMaxVal
''''                End If
''''          End Select
''''
''''          If Slider(Index).sldOrient = 1 Then   '1-Horizontal
''''             Slider(Index).sldCrntPos = (Slider(Index).sldCrntVal - Slider(Index).sldMinVal) / ((Slider(Index).sldMaxVal - Slider(Index).sldMinVal) / Slider(Index).sldNumDiv)
''''          Else                                  '2-Vertical
''''             Slider(Index).sldCrntPos = (Slider(Index).sldMaxVal - Slider(Index).sldCrntVal) / ((Slider(Index).sldMaxVal - Slider(Index).sldMinVal) / Slider(Index).sldNumDiv)
''''          End If
''''          DrawButton Slider(Index), Picture1(Index)
''''    End Select
''''
''''End Sub
''''
''''Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
''''
'''''Change button position if mouse is clicked on a valid area (either mouse)
''''    ClickSlider Slider(Index), Picture1(Index), X, Y
''''
''''End Sub
''''
''''Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
''''
'''''Reset the button based on mouse position
''''    If Slider(Index).sldCrntMove Then   'MouseDown over button will make this true
''''       Picture1(Index).Refresh      'Takes out some of the choppiness of the movement
''''       AdjustButton Slider(Index), Picture1(Index), X, Y
''''    End If
''''
''''    lVolume = Format(Slider(Index).sldCrntVal, "0.0")
''''    VolumeChanged iVolIndex
''''
''''
''''
'''''''       Label6(31) = "Current Value : " & Format(Slider(Index).sldCrntVal, "0.0")
'''''''       Label6(32) = "Minimum Value : " & Slider(Index).sldMinVal
'''''''       Label6(30) = "Maximum Value : " & Slider(Index).sldMaxVal
''''
''''End Sub
''''
''''Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
''''
'''''It is remotely possible for a mouse up to occur without the mouse move event
'''''being invoked.  The code here is the same as in the mouse move to catch this.
'''''Move the button based on mouse position.
''''    If Slider(Index).sldCrntMove Then
''''       Picture1(Index).Refresh      'Takes out some of the choppiness of the movement
''''       AdjustButton Slider(Index), Picture1(Index), X, Y
''''    End If
''''
'''''Turn off mouse tracking
''''    Slider(Index).sldCrntMove = False
''''
''''
'''''****Used for demo only - show some current Slider values in labels
''''       Label6(31) = "Current Value : " & Format(Slider(Index).sldCrntVal, "0.0")
''''       Label6(32) = "Minimum Value : " & Slider(Index).sldMinVal
''''       Label6(30) = "Maximum Value : " & Slider(Index).sldMaxVal
''''
''''End Sub

Private Sub SSPanel4_Click()
MainLevelColor
End Sub



Private Sub SSPanel9_Click()

If lstSystem.Visible = True Then
   lstSystem.Visible = False
 '  sspDevice.Visible = False
Else
   lstSystem.Visible = True
 '  sspDevice.Visible = True
End If

End Sub

'''''Private Sub sspButtonDirection_Click(Index As Integer)
'''''Dim bChange As Boolean
'''''
'''''On Error GoTo err1
'''''
'''''bChange = True  'Change the colors of the buttons
'''''If bLoading Then Exit Sub
'''''
'''''If (cmdSong.Count - 1) <> iMaxBut Then
'''''   bChange = False
'''''ElseIf sspButtonDirection(Index).BackColor = vbDirectionColor Then
'''''   Exit Sub
'''''End If
'''''
'''''If InStr(1, sspStream1.Caption, "Ready") = 0 Then Exit Sub
'''''If InStr(1, sspStream2.Caption, "Ready") = 0 Then Exit Sub
'''''
''''''''''''Change the button colors...
'''''''''''If bChange Then
'''''''''''   If sspButtonDirection(1).BackColor = vbBlack Then
'''''''''''      sspButtonDirection(2).BackColor = vbBlack
'''''''''''      sspButtonDirection(1).BackColor = vbDirectionColor
'''''''''''      sspButtonDirection(2).ForeColor = vbDirectionColor
'''''''''''      sspButtonDirection(1).ForeColor = vbBlack
'''''''''''   Else
'''''''''''      sspButtonDirection(1).BackColor = vbBlack
'''''''''''      sspButtonDirection(2).BackColor = vbDirectionColor
'''''''''''      sspButtonDirection(1).ForeColor = vbDirectionColor
'''''''''''      sspButtonDirection(2).ForeColor = vbBlack
'''''''''''   End If
'''''''''''End If
'''''''''''
''''''''''''Get the correct layout indicator
'''''''''''iButtonDirection = Index
''''''Disable the form
'''''frmPlayer.Enabled = False
'''''Screen.MousePointer = vbHourglass
'''''DoEvents
''''''Clear and load buttons in new layout
'''''RemoveButtons
'''''SetButtonsLayout
'''''ClearButtons
'''''DoEvents
'''''LoadPalette PaletteName
''''''Enable the screen again...
'''''frmPlayer.Enabled = True
'''''Screen.MousePointer = vbDefault
'''''DoEvents
'''''
'''''Exit Sub
'''''err1:
'''''MsgBox "Error in Module : sspButtonDirection_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation
'''''

'''''End Sub

'''Private Sub sspButtonPlayStop_Click(Index As Integer)
'''
'''On Error GoTo err1
'''
'''If sspButtonPlayStop(Index).BackColor = vbDirectionColor Then Exit Sub
'''
'''If sspButtonPlayStop(0).BackColor = vbBlack Then
'''   sspButtonPlayStop(1).BackColor = vbBlack
'''   sspButtonPlayStop(0).BackColor = vbDirectionColor
'''   sspButtonPlayStop(1).ForeColor = vbDirectionColor
'''   sspButtonPlayStop(0).ForeColor = vbBlack
'''Else
'''   sspButtonPlayStop(0).BackColor = vbBlack
'''   sspButtonPlayStop(1).BackColor = vbDirectionColor
'''   sspButtonPlayStop(0).ForeColor = vbDirectionColor
'''   sspButtonPlayStop(1).ForeColor = vbBlack
'''End If
'''iButtonPlayStopPause = Index
'''
'''Exit Sub
'''err1:
'''MsgBox "Error in Module : sspButtonPlayStop_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation
'''
'''End Sub


Private Sub sspProgress_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
'''   On Error Resume Next
'''   If lblTimeLeft(Index).Caption = "" Then
'''      sspSongTitle(Index).TabIndex = 0
'''      If Me.lblStatus(Index).Caption = "Playing" Then Exit Sub
'''
'''      LongPress = 0
'''      SelPlayerIndex = Index
'''      Timer3.Enabled = True
'''      Exit Sub
'''   End If
   
End Sub

Private Sub sspProgress_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
DoEvents
End Sub

'''Private Sub sspOption3_Click(Index As Integer)

'On Error GoTo err1
'
'If sspOption3(Index).BackColor = vbDirectionColor Then Exit Sub
'
'iMaxBut = sspOption3(Index).Caption
'iButtonMaxSelected = Index
'
'For i = 1 To 4
'   If i = Index Then
'      sspOption3(i).BackColor = vbDirectionColor
'      sspOption3(i).ForeColor = vbBlack
'   Else
'      sspOption3(i).BackColor = vbBlack
'      sspOption3(i).ForeColor = vbDirectionColor
'   End If
'Next i
'
''Hide the volume screen, just in case...
'ButLeft = 22000
'sspVol.Left = ButLeft
'
''Disable the form
'frmPlayer.Enabled = False
'Screen.MousePointer = vbHourglass
'DoEvents
''Clear and load buttons in new layout
'RemoveButtons
'SetButtonsLayout
'ClearButtons
'DoEvents
'LoadPalette PaletteName
''Enable the screen again...
'frmPlayer.Enabled = True
'Screen.MousePointer = vbDefault
'DoEvents
'
'Exit Sub
'err1:
'MsgBox "Error in Module : sspOption3_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation
'
'''End Sub

Private Sub sspProgress_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


   On Error Resume Next
   
   If sspSongTitle(Index).Caption = "" Then
      LoadNewSong Index
      Exit Sub
   End If
   
   
   If Me.lblStatus(Index).Caption <> "Playing" Then Exit Sub
   
   SetSongPos chan(CLng(cmdSong(Index).Tag)), Index, CInt(X)
   
  ' Debug.Print sspProgress(Index).Width
      
End Sub

'''Private Sub sspProgress_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'''   On Error Resume Next
'''
'''   If Me.lblStatus(Index).Caption <> "Playing" Then Exit Sub
'''
'''   SetSongPos chan(CLng(cmdSong(Index).Tag)), Index, CInt(x)
'''
'''
'''End Sub

Private Sub sspSongTitle_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

'If TypeOf Source Is SSPanel Then
'  If Source.Index = Index Then
'    Exit Sub
'  End If
'  'If Trim(sspSongTitle(Source.Index).Caption) <> "" Then
'    'If Trim(sspSongTitle(Index).Caption) <> "" Then Exit Sub 'Make sure we dont overtype the current button where we drop to...
'    ClearButton Index
'    SetupButton Index, sspSongTitle(Source.Index).Caption, "", sspSongTitle(Source.Index).Tag
'    ClearButton Source.Index
'    Source.Caption = ""
'    'sspSongTitle(Source.Index).Caption = ""
'    Source.Tag = ""
'  'End If
'End If
'Call cmdSong_DragDrop(Index, Source, X, Y)
'Timer3.Enabled = False
'sspSongTitle(Index).DragMode = False
'Source.DragMode = False
'DragingButton = False
'LongPress = 0

End Sub

Private Sub sspSongTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo err1

sspSongTitle(Index).TabIndex = 0
If Me.lblStatus(Index).Caption = "Playing" Then Exit Sub

'''DragingButton = True

MousePointer = vbHourglass
DoEvents
LongPress = 0
SelPlayerIndex = Index
Timer3.Enabled = True

Exit Sub
err1:
MsgBox "Error in Module : sspSongTitle_MouseDown " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub sspSongTitle_MouseExit(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'MsgBox "Exit pos : " & X
End Sub

Private Sub sspSongTitle_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Screen.MousePointer = vbDefault
DoEvents

End Sub

Private Sub sspSongTitle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   
   Timer3.Enabled = False
    MousePointer = vbDefault
    DoEvents
   If LongPress >= LongPressCnt Then
      LongPress = 0
      Exit Sub
   End If
   
'''   DragingButton = False
'''   sspSongTitle(Index).DragMode = 0
   
   If Button = 2 Then
      If iSecureMode = 2 Then Exit Sub
      'LoadNewSong Index
      ShowOptionScreen Index
   Else
'      If iButtonPlayStopPause = 1 Then    'Pause
'         ProcessPause
'         Exit Sub
'      Else
         iMnuFlag = 0
       '  PlayingCurrently = Index
         Setstate Index
      'End If
   End If
   
   Command2.SetFocus

End Sub


Private Sub sspSongTitle_OLEDragDrop(Index As Integer, Data As Threed.SSDataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error Resume Next
   If Me.lblStatus(Index).Caption = "Playing" Then Exit Sub

   If Effect = 7 Then
     FilenameToLoad = Data.Files(1)
     iMnuFlag = 4
     Setstate Index
     SavePalete Trim(Me.Caption), iPageno
     LoadPalette Trim(Me.Caption), iPageno, 3
   End If

End Sub

'Private Sub sspSongTitle_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
''   On Error Resume Next
''   If Me.lblStatus(Index).Caption = "Playing" Then Exit Sub
''
''   If Effect = 7 Then
''     FilenameToLoad = Data.Files(1)
''     iMnuFlag = 4
''     Setstate Index
''     SavePalete Trim(Me.Caption)
''   End If
'End Sub



Private Sub sspVol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Closevolume
End Sub

Sub Closevolume(Optional SkipSave As Boolean = False)

On Error GoTo err1

  imgVol(iVolIndex).Visible = True

  ButLeft = 22000
  sspVol.Left = ButLeft
  cpvVol.ZOrder 1
   
  If SkipSave Then
    If Not sspSongTitle(iVolIndex).Visible Then
      imgVol(iVolIndex).Visible = False
    End If
  Else
    SavePalete Trim(Me.Caption), iPageno
    DoEvents
    LoadPalette Trim(Me.Caption), iPageno, 3
  End If
   
Exit Sub
err1:
MsgBox "Error in Module : Closevolume " & Chr(13) & Chr(13) & Err.Description, vbExclamation
   
End Sub

Private Sub SSScroll1_Click()

End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   sspTime.Caption = Format(Now, "HH:MM")
   lblDate.Caption = Format(Now, "DD MMMM YYYY")
   
   If sspTimeInd.ForeColor = vbCyan Then
      sspTimeInd.ForeColor = &H808000
   Else
      sspTimeInd.ForeColor = vbCyan
   End If

End Sub

Sub ClearArray()

End Sub

Sub LoadPalette(pHeading As String, PageNo As Integer, LoadOption As Integer)
  Dim FileToOpen As String
  
  If Not TimerMainLevel.Enabled Then
     InitialisePeaks 0
     DoEvents
  End If
   
  FileToOpen = App.Path & "\Palets\" & pHeading & ".dat"
  
  
  
  Select Case LoadOption
    Case 1  'Load ALL
      LoadPaletteArray FileToOpen
      LoadPalettePage PageNo
      WriteNewPaletteFile FileToOpen 'This is only done to list the file on the top when we search for Palette file...
      
    Case 2  'Only load the next page from array
      LoadPalettePage PageNo
      
    Case 3 'When we dragged, I need to reload the array from SAVED file
      LoadPaletteArray FileToOpen
          
    Case 4  'Only load current page from array
      LoadPalettePage PageNo
      
    Case 5 'Clear palette array
      ReDim PlArr(10, 181)  'Start at 1 with both dimensions
      
    Case Else
      MsgBox "ELSE"
  End Select

  GetTotalTime

End Sub

Sub LoadPaletteArray(FileToOpen As String, Optional PlayingIndex As Integer)
  Dim iPos As Integer
  Dim FD
  Dim sInput As String
  Dim sStr As String
  Dim sArr() As String
  Dim sVol As Single
  Dim sNow As String
  Dim iRow As Integer
  Dim time As Long
  Dim Bytes As Long

   sNow = Format(Now, "YYYYMMDDHHmmSS")
   
   On Error Resume Next

   ReDim PlArr(10, 181)  'Start at 1 with both dimensions
  
   FD = FreeFile
   
   i = -1
   
   If Not FSO.FileExists(FileToOpen) Then Exit Sub

   iRow = 0
   
   Open FileToOpen For Input As FD
   Do Until (EOF(FD) = True)
      Line Input #FD, sInput
      'This is to skip the first line containing the create date/time
      If Left(sInput, 3) = "000" Then
        Line Input #FD, sInput
      End If
      
      'From here on, we are dealing with the actual data
      iRow = iRow + 1 'This will be the iteration from 1 to 180 - Song entries, we will take care of which button in seperate procedure
      iPos = InStr(3, sInput, "|")
      If iPos > 0 Then
         '==================================================================================
         'For Demo system, only allow load of 5 songs
         If DemoFlag Then
            DemoCnt = DemoCnt + 1
            If DemoCnt > DemoMax Then
               'MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3, vbExclamation, DemoHeading
               'Load the DEMO image and skip to loop end
               'cmdSong(iButNum).picture
               
               'GoTo ReadNext
               Exit Do
            End If
         End If
         '==================================================================================
            
         'Exception takes place here. If the current song is playing, do not stop or load new info here
        ' If lblStatus(iRow).Caption <> "Playing" Then
            sArr = Split(sInput, "|")
            'Check if file exists on the HD still...
            'Skip any entries that does NOT exist on the HD anymore
            If Not FSO.FileExists(sArr(1)) Then GoTo ReadNext
            'Get the info from line and populate the Array...
            sArr(0) = Mid(sArr(0), 5)
            PlArr(PLA.efTtle, iRow) = sArr(0)               'Keep Full original title here
            PlArr(PLA.eTtle, iRow) = FixSongTitle(sArr(0))  'Fix the above to show nice title ('Determine if there are "-" in the title array, if so, split the 2)
            PlArr(PLA.eFN, iRow) = sArr(1)                  'Keep the filename here
            sVol = 70                                       'Set Default value for Volume = 70%
            If Val(sArr(2)) > 0 Then sVol = Val(sArr(2))    'Get the volume
            If sVol > 100 Then sVol = 99.999
            PlArr(PLA.eVol, iRow) = sVol
            PlArr(PLA.eAve, iRow) = sArr(3)                 'Use this to keep the Average when song is loaded...
            PlArr(PLA.eClr, iRow) = Val(sArr(4))            'Color
            'Get the time for song
            Call BASS_StreamFree(chan(3))        'free the old stream
            chan(3) = BASS_StreamCreateFile(BASSFALSE, StrPtr(PlArr(PLA.eFN, iRow)), 0, 0, 0)
            Bytes = BASS_ChannelGetLength(chan(3), BASS_POS_BYTE)
            time = BASS_ChannelBytes2Seconds(chan(3), Bytes)
           
            'Get the starting position by finding the silence in the front and set start after silence
            PlArr(PLA.eLs, iRow) = CStr(ScanForLeadingSilences(CStr(PlArr(PLA.eFN, iRow)), iRow))
            PlArr(PLA.eTp, iRow) = Trim(Format((time \ 60), "00") & ":" & Format(time Mod 60, "00"))
            'Reset the Bass channel for next check
            chan(3) = 0
       '  End If
      End If
ReadNext:
  Loop
  
Close1:
  Close FD
  
  Call BASS_StreamFree(chan(3))  ' free the stream

  Exit Sub
err1:
MsgBox "Error in Module : LoadPaletteArray " & Chr(13) & Chr(13) & Err.Description, vbExclamation
   Resume Next
   Close FD

End Sub

Sub WriteNewPaletteFile(FileToOpen As String)
Dim sArr(181)
Dim i As Integer
Dim sNow As String
Dim FD
Dim FW

 'Update the time we accessed it, so this list will always be on top
   'Open for input, so we can load the array with all the entries.
   i = -1
   
   'FileToOpen = App.Path & "\Palets\" & pHeading & ".dat"
   FD = FreeFile
   sNow = Format(Now, "YYYYMMDDHHmmSS")
   Open FileToOpen For Input As FD
   FW = FreeFile
   Open App.Path & "\Palets\TmpOutput.txt" For Output As FW
   
   Do Until (EOF(FD) = True)
      i = i + 1
      If i > 180 Then Exit Do
      Line Input #FD, sArr(i)
      'Now test for NOW line...  AND OVERWRITE with true NOW
      If Left(sArr(i), 4) = "000:" Then
         sArr(i) = "000:" & sNow
      End If
      'Print the data to a new file...
      Print #FW, sArr(i)
   Loop
   Close FD
   Close FW
   
    'Remove the Tmp001,dat file, sice we just re-created it
    If FSO.FileExists(FileToOpen) Then
    'If FSO.FileExists(App.Path & "\Palets\TmpOutput.txt") Then
       FSO.DeleteFile FileToOpen
       FSO.CopyFile App.Path & "\Palets\TmpOutput.txt", FileToOpen
       FSO.DeleteFile App.Path & "\Palets\TmpOutput.txt"
       DoEvents
    End If

Exit Sub
err1:
MsgBox "Error in Module : WriteNewPaletteFile " & Chr(13) & Chr(13) & Err.Description, vbExclamation
   Resume Next
   Close FD
   Close FW
   
End Sub

Sub LoadPalettePage(PageNo As Integer)
   Dim tmpIMaxBut As Integer
   Dim tmpIbutStart As Integer
   Dim tmpButNo As Integer
   Dim iButIncrease As Integer
   Dim iInner As Integer
   Dim iButNum As Integer
'   Dim PlayingIndex As Integer
   
'   PlayingIndex = 0
'   If Player1Index <> 0 Then
'      PlayingIndex = Player1Index
'   ElseIf Player2Index <> 0 Then
'      PlayingIndex = Player2Index
'   End If

  iPageno = PageNo
  Select Case iMaxBut
      Case 9
         iButIncrease = 8
      Case 16
         iButIncrease = 15
      Case 20
         iButIncrease = 19
      Case 30
         iButIncrease = 29
  End Select
   
   Select Case PageNo
     Case 1 '1-16
       tmpIbutStart = 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
     Case 2 '17-32
       'tmpIbutStart = 17
       tmpIbutStart = iMaxBut + 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 32
     Case 3 '33-48
       tmpIbutStart = iMaxBut + iMaxBut + 1
       'tmpIbutStart = 33
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 48
     Case 4 '49-64
       'tmpIbutStart = 49
       tmpIbutStart = iMaxBut + iMaxBut + iMaxBut + 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 64
       
     Case 5 '65-80
       'tmpIbutStart = 65
       tmpIbutStart = iMaxBut + iMaxBut + iMaxBut + iMaxBut + 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 80
       
     Case 6 '81-96
       'tmpIbutStart = 81
       tmpIbutStart = iMaxBut + iMaxBut + iMaxBut + iMaxBut + iMaxBut + 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 96
       
   End Select
   iInner = 0
   
   'Now load the buttons from array
   For iButNum = 1 To 180

      If iButNum < tmpIbutStart Then GoTo ReadNext
      If iButNum > tmpIMaxBut Then Exit For   ' GoTo ReadNext
      If iButNum = 0 Then GoTo ReadNext

'''      If tmpIbutStart = 1 Then
'''         iInner = iButNum
'''      Else
'''         iInner = iButNum - (iMaxBut * (iPageno - 1))
'''         iInner = tmpIMaxBut - iMaxBut
'''      End If  'iButNum
'''
'''
'''      Select Case iButNum
'''         Case 1 To tmpIMaxBut  '16
'''            iInner = iButNum
'''         Case 17 To 32
'''            iInner = iButNum - 16
'''         Case 33 To 48
'''            iInner = iButNum - 32
'''         Case 49 To 64
'''            iInner = iButNum - 48
'''         Case 65 To 80
'''            iInner = iButNum - 64
'''         Case 81 To 100
'''
'''         Case Is > 80
'''            iInner = iButNum - 80
'''      End Select
      
    iInner = iButNum - GetNextButtonNumber
   ' If lblStatus(iButNum).Caption <> "Playing" Then
      If PlArr(PLA.eTtle, iButNum) <> "" Then
        sspSongTitle(iInner).Caption = PlArr(PLA.eTtle, iButNum)       'Fixed SongTitle
        sspSongTitle(iInner).TagVariant = PlArr(PLA.efTtle, iButNum)   'Keep Full title here
        cmdSong(iInner).TagVariant = PlArr(PLA.eAve, iButNum)          'Use this to keep the Average when song is loaded...
        lblVol(iInner).Caption = PlArr(PLA.eVol, iButNum)              'Keep Volume here
        cmdSong(iInner).Tag = 3                                     'To RESET button
        sspSongTitle(iInner).Tag = PlArr(PLA.eFN, iButNum)             'Keep the filename here
        lblTimeLeft(iInner).Tag = PlArr(PLA.eLs, iButNum)              'Leading Silences
        lblTimeLeft(iInner).Caption = "00:00"                       'Sets the Left lable to 0:0
        lblTimePlayed(iInner).Tag = PlArr(PLA.eTp, iButNum)            'Sets Time Played (Right label) Put in tag for later use
        lblTimePlayed(iInner).Caption = lblTimePlayed(iInner).Tag   'Sets Time Played (Right label)
        cmdSong(iInner).Tag = ""
        ResetButton iInner
        SetButtonColor iInner, CInt(PlArr(PLA.eClr, iButNum))          'Sets Button Color
        'Determine volume, if saved with it...
        lblVol(iInner).Caption = PlArr(PLA.eVol, iButNum)              'Sets the Song Volume display - Set to 75% of max  (75000)
        imgVol(iInner).Visible = True
        imgDirection(iInner).Visible = True
      End If
   ' End If
    DoEvents
ReadNext:
   Next iButNum
   
   
   
End Sub

'''public Function FixSongTitle(sData As String) As String
'''Dim sTitle() As String
'''
'''Dim sTemp As String
'''Dim MaxChars As Integer
'''
'''Const sNameChrs As String = "**"
'''
'''Select Case iMaxBut
'''   Case 9, 16
'''      'iChr13 = 2
'''      MaxChars = 40
'''   Case 20
'''      'iChr13 = 1
'''      MaxChars = 30
'''   Case 30
'''      'iChr13 = 1
'''      MaxChars = 28
'''
'''End Select
'''
'''sTemp = sData
'''
'''If InStr(1, sTemp, "-") = 0 Then
'''   If InStr(1, sTemp, "_") > 0 Then
'''      sTemp = Replace(sTemp, "_", "-")
'''   End If
'''End If
'''sTemp = Replace(sTemp, "/", " / ")
'''sTitle = Split(Replace(sTemp, "&", "&&"), "-")
'''
'''
'''
''''Select Case UBound(sTitle)
''''   Case 1
''''      If iMaxBut = 30 And Len(Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs) > 70 Then
''''         FixSongTitle = Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
''''      Else
''''         FixSongTitle = String(iChr13, Chr(13)) & Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
''''      End If
''''   Case 2
''''      FixSongTitle = String(iChr13, Chr(13)) & Trim(sTitle(1)) & " " & Trim(sTitle(2)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
''''   Case 3
''''      FixSongTitle = String(1, Chr(13)) & Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "]" & Chr(13) & Trim(sTitle(2)) & Chr(13) & Trim(sTitle(3))
''''   Case 4
''''      FixSongTitle = sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
''''   Case Else
''''      If iMaxBut = 9 Or iMaxBut = 25 Then
''''         FixSongTitle = String(iChr13 + 1, Chr(13)) & Trim(sTitle(0))
''''      Else
''''         FixSongTitle = String(iChr13, Chr(13)) & Trim(sTitle(0))
''''      End If
''''End Select
'''
'''Select Case UBound(sTitle)
'''   Case Is > 1
'''      For i = 1 To UBound(sTitle)
'''         FixSongTitle = FixSongTitle & Trim(sTitle(i)) & "-"
'''      Next i
'''      If Right(FixSongTitle, 1) = "-" Then FixSongTitle = Mid(FixSongTitle, 1, Len(FixSongTitle) - 1)
'''      FixSongTitle = FixSongTitle & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'''   Case 1
'''      If Len(Trim(sTitle(1))) > MaxChars Then sTitle(1) = FixTitle(sTitle(1), MaxChars)
'''      If Len(Trim(sTitle(0))) > MaxChars Then sTitle(0) = FixTitle(sTitle(0), MaxChars)
'''
'''      FixSongTitle = Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'''   Case Else
'''      If Len(Trim(sTitle(0))) > MaxChars Then sTitle(0) = FixTitle(sTitle(0), MaxChars)
'''      FixSongTitle = sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'''
'''End Select
''''If Trim(sTitle(1)) <> "" And Trim(sTitle(0)) <> "" Then  'Both present
''''   FixSongTitle = Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
''''ElseIf Trim(sTitle(1)) = "" Then
''''   FixSongTitle = sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
''''Else
''''   FixSongTitle = Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
''''End If
'''
'''
''''Debug.Print "sData : " & sData & "  (" & Len(sData) & ")" & Chr(13) & "sTitle(0) : " & sTitle(0) & "  (" & Len(sTitle(0)) & ")" & Chr(13) & "sTitle(1) : " & sTitle(1) & "  (" & Len(sTitle(1)) & ")" & Chr(13) & "FIXED : " & FixSongTitle & "  (" & Len(FixSongTitle) & ")"
'''
'''
'''End Function
'''
'''Function FixTitle(sTitle As String, MaxChars As Integer) As String
'''Dim iChr13 As Integer
'''Dim iCnt As Integer
'''
'''   For iCnt = MaxChars To 1 Step -1
'''      If Mid(sTitle, iCnt, 1) = " " Then
'''         iChr13 = iCnt  'Get last position before end where we have a space so I can wrap it there
'''         Exit For
'''      End If
'''   Next iCnt
'''   FixTitle = Left(sTitle, iChr13) & Chr(13) & Mid(sTitle, iChr13 + 1)
'''
'''End Function

Private Sub Timer2_Timer()

iFlood = iFlood + 10
If iFlood > 100 Then iFlood = 0

sspFlood.FloodPercent = iFlood

End Sub

Private Sub Timer3_Timer()

   On Error Resume Next
   
   LongPress = LongPress + 1
   
'''   If DragingButton Then
'''      sspSongTitle(SelPlayerIndex).DragMode = 1
'''      Exit Sub
'''   End If
   
   If LongPress > LongPressCnt Then
      Timer3.Enabled = False
      MousePointer = vbDefault
      DoEvents
      bTagEditMP3 = UCase(Right(sspSongTitle(SelPlayerIndex).Tag, 3)) = "MP3"
      XPos = sspSongTitle(SelPlayerIndex).Parent.X
      If lblStatus(SelPlayerIndex).Caption = "Ready" Then ShowOptionScreen SelPlayerIndex
      'ShowOptionScreen SelPlayerIndex

      Exit Sub
      
   End If
   
End Sub

Sub ShowVolScreen(Index As Integer)

Dim iSpace As Integer
Dim iCnt As Integer
Dim iWidth As Integer
On Error GoTo err1

  ' cmdCloseVol.Picture = LoadResPicture(134, vbResBitmap)
  ' cmdCloseVol.Picture = LoadResPicture(134, vbResIcon)
   cmdCloseVol.Picture = LoadPicture(App.Path & "\tmpclose")
   imgVol(Index).Visible = False
   
   ButLeft = cmdSong(Index).Left + 15  '+ 30
   
   lblVolInd.Font.Size = 7
   lblVolInd.Top = 0
   cmdCloseVol.Height = 375
   
   
   Select Case iMaxBut
      Case 30
         ButTop = PanMain.Top + cmdSong(Index).Top + (sspProgress(Index).Top - 320) + 150
         iWidth = cmdSong(Index).Width - 140   '- sspLevel(Index).Width - 140
         sspVol.Height = 425
         lblVolInd.Font.Size = 6
         lblVolInd.Top = 30
         cmdCloseVol.Height = 345
         cmdCloseVol.Top = 60
      Case 20
         ButTop = PanMain.Top + cmdSong(Index).Top + (sspProgress(Index).Top - 270) + 105
         iWidth = cmdSong(Index).Width - 150  'sspLevel(Index).Width - 150
         sspVol.Height = 425
         cmdCloseVol.Top = 60
         cmdCloseVol.Height = 360
         lblVolInd.Font.Size = 7
      Case 16
         ButTop = PanMain.Top + cmdSong(Index).Top + (sspProgress(Index).Top - 200)
         ButTop = ButTop + 60
         iWidth = cmdSong(Index).Width - 160   'sspLevel(Index).Width - 160
         sspVol.Height = 420   '525
         cmdCloseVol.Top = 60 '120
      Case 9
         ButTop = PanMain.Top + cmdSong(Index).Top + (sspProgress(Index).Top - 170)
         iWidth = sspSongTitle(Index).Width - 1  '   cmdSong(Index).Width - 160   'sspLevel(Index).Width - 160
         sspVol.Height = 525
         cmdCloseVol.Top = 120
   End Select
   
   cmdCloseVol.Width = cmdCloseVol.Height
    
   sspVol.Left = ButLeft
   sspVol.Top = ButTop
  ' sspVol.Height = 525   '600  '690  '560
   'Change the size accoring to the layout
   'sspVol.Height = cmdSong(Index).Height
   
   iWidth = (sspSongTitle(Index).Width + sspSongTitle(Index).Left) + 15  '- 15
   sspVol.Width = iWidth - 90 ' + 60 + 60
   
   cpvVol.Left = (cmdCloseVol.Left + cmdCloseVol.Width) + 115
   cpvVol.Width = sspVol.Width - cpvVol.Left - 50
   cpvVol.Height = sspProgress(Index).Height
   
   lblVolInd.Left = cpvVol.Left
   lblVolInd.Width = cpvVol.Width
   
   cpvVol.Refresh
   DoEvents
   
   sspVol.BackColor = cmdSong(Index).BackColor  '   vbNCompleted  'vbDefaultBack    'cmdSong(Index).BackColor
   sspVol.BorderWidth = 0
   sspVol.BevelInner = ssNoneBevel
   sspVol.BevelOuter = ssNoneBevel
   sspVol.Outline = False
  
Exit Sub
err1:
MsgBox "Error in Module : ShowVolScreen " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Sub InitFX(Index As Integer)
'''   Dim p As BASS_DX8_PARAMEQ
'''
'''   If cmdSong(Index).Tag = "" Then
'''      Exit Sub
'''   End If
'''
'''   'Load the handles here
'''    fx(0) = BASS_ChannelSetFX(chan(CInt(cmdSong(Index).Tag)), BASS_FX_DX8_PARAMEQ, 0) ' bass
'''    fx(1) = BASS_ChannelSetFX(chan(CInt(cmdSong(Index).Tag)), BASS_FX_DX8_PARAMEQ, 0) ' mid
'''    fx(2) = BASS_ChannelSetFX(chan(CInt(cmdSong(Index).Tag)), BASS_FX_DX8_PARAMEQ, 0) ' treble
'''  '  fx(3) = BASS_ChannelSetFX(chan(CInt(cmdSong(Index).Tag)), BASS_FX_DX8_REVERB, 0)  ' reverb
'''
'''
'''   'Bass
'''   Call BASS_FXGetParameters(fx(0), p)
'''   p.fBandwidth = 18
'''   p.fCenter = iBassFreq '[83 hz]
'''   p.fGain = 10 - lBass
'''   Call BASS_FXSetParameters(fx(0), p)
'''
'''   'Mids
'''   Call BASS_FXGetParameters(fx(1), p)
'''   p.fBandwidth = 18
'''   p.fCenter = iMidFreq '[1 Khz]
'''   p.fGain = 10 - lMid
'''   Call BASS_FXSetParameters(fx(1), p)
'''
'''   'Highs
'''   Call BASS_FXGetParameters(fx(2), p)
'''   p.fBandwidth = 18
'''   p.fCenter = iHighFreq '[8 Khz]
'''   p.fGain = 10 - lHigh
'''   Call BASS_FXSetParameters(fx(2), p)


End Sub

Sub ShowOptionScreen(Index As Integer)
Dim ErrDesc As String
On Error GoTo err1
   
   ErrDesc = "Set Buttons position with INDEX=" & Index & Chr(13) & _
             "cmdSong.LEFT:" & cmdSong(Index).Left & Chr(13) & _
             "cmdSong.Width:" & cmdSong(Index).Width & Chr(13) & _
             "Screen.Activeform.Left:" & Screen.ActiveForm.Left
   
'   ButLeft = Screen.ActiveForm.Left + cmdSong(Index).Left + (cmdSong(Index).Width - 2835)
'   ButTop = Screen.ActiveForm.Top + PanMain.Top + 1000
   
   ButLeft = Screen.ActiveForm.Left + cmdSong(Index).Left + (cmdSong(Index).Width - 2835)
   ButTop = Screen.ActiveForm.Top + PanMain.Top + 1000
   
   
   
   gColor = -1
    ErrDesc = "frmSelectButtonColor.show"
    frmSelectButtonColor.Show vbModal
   
   If bExitSetup Then Exit Sub
   
   If gColor <> -1 Then
      ErrDesc = "SetButtonColor Index=" & Index & " , gColor=" & gColor
      SetButtonColor Index, gColor  'Set the color of the button accordingly...
      ErrDesc = "palletArr(0) = "
      palletArr(0) = ""
      If Trim(Me.Caption) = "" Then
         Me.Caption = "tmp001"
         ErrDesc = "palletArr(0) = Trim(Me.Caption)"
         palletArr(0) = Trim(Me.Caption)
      End If
      ErrDesc = "SavePalete Trim(Me.Caption)"
      FixPaletteButtonArray PLA.eClr, Index, CStr(gColor)
      SavePalete Trim(Me.Caption), iPageno
      LoadPalette Trim(Me.Caption), iPageno, 3
      Exit Sub
   End If
   
   If bClearButton Then
      ErrDesc = "ClearButton (Index=" & Index & ")"
      ClearButton (Index)
      ErrDesc = "palletArr(0) ="
      palletArr(0) = ""
      If Trim(Me.Caption) = "" Then
         Me.Caption = "tmp001"
         palletArr(0) = Trim(Me.Caption)
      End If
      ErrDesc = "SavePalete Trim(Me.Caption)"
      SavePalete Trim(Me.Caption), iPageno
      LoadPalette Trim(Me.Caption), iPageno, 3
      Exit Sub
   End If
   
   If bLoadButton Then
      If lblStatus(Index).Caption <> "Playing" Then
        iMnuFlag = 2
        ErrDesc = "Setstate Index=" & Index
        Setstate Index
      End If
   End If
   
   If bTagEdit Then
      If lblStatus(Index).Caption = "Playing" Then Exit Sub
      ErrDesc = "FilenameToLoad = sspSongTitle(Index).Tag"
      FilenameToLoad = sspSongTitle(Index).Tag
      
      bTagsUpdated = False
       ErrDesc = "frmTestMp3Tags.Show vbModal"
       frmTestMp3Tags.Show vbModal
      
      If bTagsUpdated Then
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         If FilenameToLoad = "" Then
           Screen.MousePointer = vbDefault
           Me.Enabled = True
           Exit Sub
         End If
         vbKeepColor = Val(sspProgress(Index).Tag)
         ClearButton Index
         'Reset button to Loaded state
         ResetButton Index
         'Add caption to button
         lblStatus(Index).Caption = "Loading..."
         ExtractTagInfo FilenameToLoad
         SetupButton Index, Id3TagArr(1), Id3TagArr(2), FilenameToLoad
         vbKeepColor = 0
         If Not TimerMainLevel.Enabled Then
            InitialisePeaks 0
         End If
      End If
   End If
   
'''   If bDucking Then
'''      If lblStatus(Index).Caption = "Playing" Then Exit Sub
'''      FilenameToLoad = sspSongTitle(Index).Tag
'''
'''       frmDucking.Show vbModal
'''
'''      If ReturnPos > 0 Then
'''         'Set the STARTING position according to what was set on the Ducking form
'''         'SetSongPosFromDucking chan(CLng(cmdSong(Index).Tag)), Index, ReturnPos
'''      End If
'''
'''   End If
   Me.Enabled = True
   
Exit Sub
err1:
MsgBox "Error in Module : ShowOptionScreen (" & ErrDesc & ")" & Chr(13) & Chr(13) & Err.Description, vbExclamation
      
End Sub

Private Sub Timer4_Timer()

Timer4.Enabled = False
If DemoFlag Then ShowPlayerDemoMsg

End Sub

Sub ShowPlayerDemoMsg()

   'If DemoFlag Then
      If 5 - iCntDemo = 1 Then
         SSPanel9.ForeColor = vbYellow
'         'MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3 & Chr(13) & Chr(13) & "The DEMO system may only be used 1 more time.", vbExclamation, DemoHeading
'         If MsgBox(vbTab & DemoMsg1 & Chr(13) & vbTab & "==============" & Chr(13) & Chr(13) & vbTab & "-   " & DemoMsg3 & Chr(13) & vbTab & "-   " & "The DEMO system may be used 1 more times." & Chr(13) & Chr(13) & vbTab & "Would you like to Register now??", vbQuestion + vbYesNo, DemoHeading) = vbYes Then
'            GoTo EnterSerial
'         End If
      ElseIf 5 - iCntDemo = 0 Then
         SSPanel9.ForeColor = vbRed
         'MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3 & Chr(13) & Chr(13) & "The DEMO system may be used for the LAST time.", vbExclamation, DemoHeading
'         If MsgBox(vbTab & DemoMsg1 & Chr(13) & vbTab & "==============" & Chr(13) & Chr(13) & vbTab & "-   " & DemoMsg3 & Chr(13) & vbTab & "-   " & "The DEMO system may be used for the LAST time." & Chr(13) & Chr(13) & vbTab & "Would you like to Register now??", vbQuestion + vbYesNo, DemoHeading) = vbYes Then
'            GoTo EnterSerial
'         End If
      Else
         SSPanel9.ForeColor = vbGreen
'         If MsgBox(vbTab & DemoMsg1 & Chr(13) & vbTab & "==============" & Chr(13) & Chr(13) & vbTab & "-   " & DemoMsg3 & Chr(13) & vbTab & "-   " & "The DEMO system may be used " & 5 - iCntDemo & " more times." & Chr(13) & Chr(13) & vbTab & "Would you like to Register now??", vbQuestion + vbYesNo, DemoHeading) = vbYes Then
'            GoTo EnterSerial
'         End If
      End If
   'End If

   On Error Resume Next
   Me.SetFocus
   Command2.SetFocus

   Exit Sub

'EnterSerial:
'   frmSerial.Show vbModal

End Sub

Private Sub TimerF1_Timer()

   On Error Resume Next
   Stop1 = True
   SetBlinkColor Player1Index

End Sub

Private Sub TimerF2_Timer()
   
   On Error Resume Next
   Stop2 = True
   SetBlinkColor Player2Index
   

End Sub

Sub SetBlinkColor(Index As Integer)
   Dim iColor As Integer
   On Error GoTo err1
   
   iColor = Val(sspProgress(Index).Tag) 'VAL()   Will also force to 0 if nothing setup
   If lblStatus(Index).ForeColor = vbBlack Then
      lblStatus(Index).ForeColor = vbWhite
      sspSongTitle(Index).ForeColor = vbBlack
'      Set sspSongTitle(Index).Picture = Nothing
'      sspSongTitle(Index).BevelOuter = ssInsetBevel
      sspSongTitle(Index).BackColor = vbNDYellow
   Else
      lblStatus(Index).ForeColor = vbBlack
      sspSongTitle(Index).ForeColor = vbWhite
'      sspSongTitle(Index).BevelOuter = ssNoneBevel
      sspSongTitle(Index).BackColor = ButColors(iColor)
   End If
   lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
   
Exit Sub
err1:
MsgBox "Error in Module : SetBlinkColor " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Sub SetFadeColor(Index As Integer)
   Dim iColor As Integer
   On Error GoTo err1
   
   iColor = Val(sspProgress(Index).Tag) 'VAL()   Will also force to 0 if nothing setup
   If lblStatus(Index).ForeColor = vbBlack Then
      lblStatus(Index).ForeColor = vbWhite
      sspSongTitle(Index).ForeColor = vbBlack
'      Set sspSongTitle(Index).Picture = Nothing
'      sspSongTitle(Index).BevelOuter = ssInsetBevel
      sspSongTitle(Index).BackColor = vbWhite   'vbFadeOut
   Else
      lblStatus(Index).ForeColor = vbBlack
      sspSongTitle(Index).ForeColor = vbWhite
'      sspSongTitle(Index).BevelOuter = ssNoneBevel
      sspSongTitle(Index).BackColor = ButColors(iColor)
   End If
   lblButtonCnt(Index).ForeColor = sspSongTitle(Index).ForeColor
   
Exit Sub
err1:
MsgBox "Error in Module : SetBlinkColor " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub TimerMainLevel_Timer()
Dim TotL As Long
Dim TotR As Long
On Error GoTo err1

   If Left1Chan > Left2Chan Then
      LeftMChan = Left1Chan
      RightMChan = Right1Chan
   ElseIf Left2Chan > Left1Chan Then
      LeftMChan = Left2Chan
      RightMChan = Right2Chan
   Else
      LeftMChan = Left1Chan Xor Left2Chan
      RightMChan = Right1Chan Xor Right2Chan
   End If

   SetMainPeakLevel LeftMChan, RightMChan
   
Exit Sub
err1:
MsgBox "Error in Module : TimerMainLevel_Timer " & Chr(13) & Chr(13) & Err.Description, vbExclamation

   
End Sub

Private Sub TimerP1_Timer()
   On Error Resume Next
   Dim pos As Single
   Dim iSecs As Integer
   Dim iMins As Integer
   Dim iPos As Integer
   'Dim TimeLeft As Long
   Dim TimeElapsedPerc As Long
   Dim STime As String
   
   On Error Resume Next
   
   If BASS_ChannelIsActive(chan(1)) = 0 Then
     pos = -1 ' reached the end
   Else
     'Get current possition of playing...
     pos = Format(bassTime.GetPlayingPos(chan(1)), "0")
   End If
   
   'Check if END Reached. Stop timer, stop playing and reset button...
   If pos = -1 Then
     TimerP1.Enabled = False
     StopSong Player1Index
     ResetButton Player1Index, True
     'Set Player to black to indicate, it did already play...
      If iAutoAdvance = 2 Then
         StartNextAvalableSong (Player1Index)
'      ElseIf iAutoAdvance = 3 Then
'
'        StartNextAvalableSong (Player1Index)
      End If
     Exit Sub
   End If
   
   'Calculate the progress bar's position, as well as the time left to display
   TimeElapsedPerc = Val(pos) * 100 / Val(Duration(1))
''''   '==================================================================================
''''   'For Demo system, only allow load of 5 songs
''''   If DemoFlag Then
''''     ' DemoCnt = DemoCnt + 1
''''      If Round(pos) > DemoTime Then
''''         TimerP1.Enabled = False
''''         StopSong Player1Index
''''         ResetButton Player1Index, True
''''         MsgBox DemoMsg2, vbExclamation, DemoHeading
''''         Exit Sub
''''      End If
''''   End If
''''   '==================================================================================
   
   'TimeElapsedPerc = Val(pos) * 100 / Val(Duration(1))
   'TimeLeft = (TimeElapsedPerc / 100) * MaxWidth
   'sspProgress(Player1Index).ProValue = TimeElapsedPerc
   sspProgress(Player1Index).value = TimeElapsedPerc
 '  SetProgress Player1Index, TimeElapsedPerc
   'sspProgress(Player1Index).FloodPercent = TimeElapsedPerc
   
   iSecs = Round(pos)  'Get the seconds rounded
   STime = ConvertSecondsToTime(iSecs)
   lblTimeLeft(Player1Index).Caption = STime
   
   STime = ""
   STime = Right(bassTime.GetTime(Duration(1) - pos), 5)
   STime = Format(Left(STime, 2), "00") & Right(STime, 3)
   
   'Show Time left
   lblTimePlayed(Player1Index).Caption = STime
   If Val(Mid(STime, 1, InStr(1, STime, ":") - 1)) = 0 Then
   'If Left(sTime, 2) = "00" Then
     If Val(Right(STime, 2)) < 10 Then
       TimerF1.Enabled = True
     End If
   End If
   
Exit Sub
err1:
MsgBox "Error in Module : TimerP1_Timer " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Function ConvertSecondsToTime(inSeconds As Integer) As String
   Dim seconds As Integer
   seconds = inSeconds Mod 60
   inSeconds = (inSeconds - seconds) / 60
   Dim minutes As Integer
   minutes = inSeconds Mod 60
   inSeconds = (inSeconds - minutes) / 60
   
   ConvertSecondsToTime = Format(minutes, "00") & ":" & Format(seconds, "00")
            
End Function


Private Sub TimerP1Level_Timer()
   Dim LevelInd As Long
   Dim LevelAve As Single
   Dim LeftChan As Integer
   Dim RightChan As Integer
   Dim Levels As Single
   Dim dB As Integer
   
   On Error Resume Next
   
  ' Call BASS_ChannelGetLevelEx(chan(1), Levels, 0.015, BASS_LEVEL_STEREO)
 ' LevelInd = BASS_ChannelGetLevelEx(Chan(1), levels(), 0.02, BASS_LEVEL_STEREO)
'  Call BASS_ChannelGetLevelEx(chan(1), Levels, 0.02, BASS_LEVEL_STEREO)
  'Call BASS_ChannelGetLevelEx(chan(1), Levels, 0.02, BASS_LEVEL_STEREO)
   

LevelInd = BASS_ChannelGetLevel(chan(1))
LeftChan = Round(LoWord(LevelInd) * iVol1)
RightChan = Round(HiWord(LevelInd) * iVol1)

'LeftChan = Round(((LoWord(LevelInd) / 32767) * 100) * iVol1)
'RightChan = Round(((HiWord(LevelInd) / 32767) * 100) * iVol1)

'Levels = BASS_ChannelGetLevel(chan(1))   '/ 32767
'db = 10 * Math.Log(Level)


'Call BASS_ChannelGetData(Chan(1), levels(1), 0.02)

'peak = BASS_ChannelGetLevelEx(Chan(1), levels, 0.01, BASS_LEVEL_STEREO)    'All Chanels
'LeftChan = levels(0)
'RightChan = levels(1)

   'Call BASS_ChannelGetLevelEx(chan(1), Levels, 0.075, BASS_LEVEL_RMS)
   
   
   'Call BASS_ChannelGetLevelEx(chan(1), levels, 0.015, BASS_LEVEL_MONO)
'   LeftChan = (Levels * 100) * iVol1
'   RightChan = (Levels * 100) * iVol1
'   LeftChan = (((Levels + LevelAve) / 2) * 100) * iVol1
'   RightChan = (((Levels + LevelAve) / 2) * 100) * iVol1
   
   
'   LevelInd = BASS_ChannelGetLevel(chan(1)) BASS_LEVEL_STEREO
'   LeftChan = Round(LoWord(Levels) * iVol1)
'   RightChan = Round(HiWord(Levels) * iVol1)
     
   Left1Chan = LeftChan
   Right1Chan = RightChan
   'Set the levels for this button...
   SetPeakLevel LeftChan, RightChan, Player1Index

'   If LeftChan / 100 / 100 * 30 > 97 Then
'      picLevelL(Player1Index).value = 100
'      lblPeakL(Player1Index).Visible = True
'   Else
'      picLevelL(Player1Index).value = LeftChan / 100 / 100 * 30
'
'   End If

'   If RightChan / 100 / 100 * 30 > 97 Then
'      picLevelR(Player1Index).value = 100
'      lblPeakR(Player1Index).Visible = True
'   Else
'      picLevelR(Player1Index).value = RightChan / 100 / 100 * 30
'   End If
   
'         picLevelL(Player1Index).value = 100
'         picLevelR(Player1Index).value = 100

   
'''   picLevelL(Player1Index).Height = (LeftChan / 8)
'''   picLevelR(Player1Index).Height = (RightChan / 8)
   
''If (LeftChan / 8) > 4180 Then
''   picChan2(0).Width = 4180
''Else
''   picChan2(0).Width = (LeftChan / 8)
''End If
''
''If (RightChan / 8) > 4180 Then
''   picChan2(1).Width = 4180
''Else
''   picChan2(1).Width = (RightChan / 8)
''End If

     

'''''picChan2(0).Width = (LeftChan / 8)
'''''picChan2(1).Width = (RightChan / 8)

   
   

End Sub

Sub SetEqHandles(channel As Integer)

   'If cmdSong(Index).Tag <> "" Then
      'Load the handles here
      If fxBass(channel) = 0 Then fxBass(channel) = BASS_ChannelSetFX(chan(channel), BASS_FX_DX8_PARAMEQ, 0)  ' bass
      If fxMid(channel) = 0 Then fxMid(channel) = BASS_ChannelSetFX(chan(channel), BASS_FX_DX8_PARAMEQ, 0)   ' mid
      If fxHigh(channel) = 0 Then fxHigh(channel) = BASS_ChannelSetFX(chan(channel), BASS_FX_DX8_PARAMEQ, 0) ' treble
      '  fx(3) = BASS_ChannelSetFX(chan(CInt(cmdSong(Index).Tag)), BASS_FX_DX8_REVERB, 0)  ' reverb
   'End If
   
End Sub

Sub SetEq(Index As Integer)

''''Dim d(1023) As Single
'Dim wBass As Integer
'Dim wMids As Integer
'Dim wHighs As Integer
'Dim p As BASS_DX8_PARAMEQ
'
'
'   SetEqHandles Index
'
'   'Bass
'   Call BASS_FXGetParameters(fx(0), p)
'   p.fBandwidth = 18
'   p.fCenter = 83 '[83 hz]
'   p.fGain = 10 - lBass
'   Call BASS_FXSetParameters(fx(0), p)
'
'   'Mids
'   Call BASS_FXGetParameters(fx(1), p)
'   p.fBandwidth = 18
'   p.fCenter = 1000 '[1 Khz]
'   p.fGain = 10 - lMid
'   Call BASS_FXSetParameters(fx(1), p)
'
'   'Highs
'   Call BASS_FXGetParameters(fx(2), p)
'   p.fBandwidth = 18
'   p.fCenter = 8000 '[8 Khz]
'   p.fGain = 10 - lHigh
'   Call BASS_FXSetParameters(fx(2), p)
'
'
'   'Initialise the label values also
'   lblBass(Index).Caption = lBass
'   lblMid(Index).Caption = lMid
'   lblHigh(Index).Caption = lHigh
    
    
    
    
    




'''   wBass = 0
'''   wMids = 0
'''   wHighs = 0

'''
'''   Label1.Text = "Bass Level: " & wBass
'''   Label30.Text = "Mid Level: " & wMids
'''   Label47.Text = "High Level: " & wHighs



End Sub

Sub SetPeakLevel(LeftChanLevel As Integer, RightChanLevel As Integer, PlayerIndex As Integer)

'Set the peak labels to invisible
''''   lblPeakL(PlayerIndex).Visible = False
''''   lblPeakR(PlayerIndex).Visible = False
''''   lblMidHL(PlayerIndex).Visible = False
''''   lblMidHR(PlayerIndex).Visible = False
''''   lblMidLL(PlayerIndex).Visible = False
''''   lblMidLR(PlayerIndex).Visible = False
''''  ' picChanL.Height = 0
'''''Left Channel
''''   Select Case LeftChanLevel
''''   'Select Case LeftChanLevel / 100 / 100 * 30
''''      Case Is < PeakNormal  'Normal green
''''         picLevelL(PlayerIndex).value = LeftChanLevel
''''         'picLevelL(PlayerIndex).value = LeftChanLevel / 100 / 100 * 30
''''         'picChanL.Height = -LeftChanLevel
''''      Case (PeakNormal + 1) To PeakMidLow
''''         picLevelL(PlayerIndex).value = PeakDispMax
''''         lblMidLL(PlayerIndex).Visible = True
''''      Case (PeakMidLow + 1) To PeakMidHigh
''''         picLevelL(PlayerIndex).value = PeakDispMax
''''         lblMidLL(PlayerIndex).Visible = True
''''         lblMidHL(PlayerIndex).Visible = True
''''      Case Is > PeakMidHigh
''''         picLevelL(PlayerIndex).value = PeakDispMax
''''         lblMidLL(PlayerIndex).Visible = True    '&H00319DFF&
''''         lblMidHL(PlayerIndex).Visible = True   '&H000D50FF&
''''         lblPeakL(PlayerIndex).Visible = True   '&H000000FF&
''''   End Select
''''
'''''Right Channel
''''   Select Case RightChanLevel
''''   'Select Case LeftChanLevel / 100 / 100 * 30
''''      Case Is < PeakNormal  'Normal green
''''         picLevelR(PlayerIndex).value = LeftChanLevel
''''         'picLevelL(PlayerIndex).value = LeftChanLevel / 100 / 100 * 30
''''      Case (PeakNormal + 1) To PeakMidLow
''''         picLevelR(PlayerIndex).value = PeakDispMax
''''         lblMidLR(PlayerIndex).Visible = True
''''      Case (PeakMidLow + 1) To PeakMidHigh
''''         picLevelR(PlayerIndex).value = PeakDispMax
''''         lblMidLR(PlayerIndex).Visible = True
''''         lblMidHR(PlayerIndex).Visible = True
''''      Case Is > PeakMidHigh
''''         picLevelR(PlayerIndex).value = PeakDispMax
''''         lblMidLR(PlayerIndex).Visible = True
''''         lblMidHR(PlayerIndex).Visible = True
''''         lblPeakR(PlayerIndex).Visible = True
''''   End Select
      
End Sub

Sub SetMainPeakLevel(LeftChanLevel As Integer, RightChanLevel As Integer)
Dim il As Integer

On Error GoTo err1

'Set the peak labels to invisible
For il = 0 To 3
   lblPeakML(il).Visible = False
   lblPeakMR(il).Visible = False
Next il
   
   
   'Left Channel
   Select Case LeftChanLevel
      Case 0
         'InitialisePeaks 0
         pgLeft.value = 0   '30473
         lblPeakML(0).Visible = False
         lblPeakML(1).Visible = False
         lblPeakML(2).Visible = False
         lblPeakML(3).Visible = False
      Case Is < PeakNormal  'Normal green
         pgLeft.value = LeftChanLevel
      Case Is = PeakNormal  'Normal green
         pgLeft.value = MaxLevelVal   '30473
         lblPeakML(0).Visible = True
      Case (PeakNormal + 1) To PeakMidLow
         pgLeft.value = MaxLevelVal   '30473
         lblPeakML(0).Visible = True
         lblPeakML(1).Visible = True
      Case (PeakMidLow + 1) To PeakMidHigh
         pgLeft.value = MaxLevelVal   '30473
         lblPeakML(0).Visible = True
         lblPeakML(1).Visible = True
         lblPeakML(2).Visible = True
      Case Is > PeakMidHigh
         pgLeft.value = MaxLevelVal   '30473
         lblPeakML(0).Visible = True
         lblPeakML(1).Visible = True
         lblPeakML(2).Visible = True
         lblPeakML(3).Visible = True
   End Select
   
   'Righ Channel
   Select Case RightChanLevel
      Case 0
         'InitialisePeaks 0
         pgRight.value = 0   '30473
         lblPeakMR(0).Visible = False
         lblPeakMR(1).Visible = False
         lblPeakMR(2).Visible = False
         lblPeakMR(3).Visible = False
      Case Is < PeakNormal  'Normal green
         pgRight.value = RightChanLevel
      Case Is = PeakNormal  'Normal green
         pgRight.value = MaxLevelVal  '30473
         lblPeakMR(0).Visible = True
      Case (PeakNormal + 1) To PeakMidLow
         pgRight.value = MaxLevelVal  '30473
         lblPeakMR(0).Visible = True
         lblPeakMR(1).Visible = True
      Case (PeakMidLow + 1) To PeakMidHigh
         pgRight.value = MaxLevelVal  '30473
         lblPeakMR(0).Visible = True
         lblPeakMR(1).Visible = True
         lblPeakMR(2).Visible = True
      Case Is > PeakMidHigh
         pgRight.value = MaxLevelVal  '30473
         lblPeakMR(0).Visible = True
         lblPeakMR(1).Visible = True
         lblPeakMR(2).Visible = True
         lblPeakMR(3).Visible = True
   End Select

Exit Sub
err1:
MsgBox "Error in Module : SetMainPeakLevel " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub TimerP2_Timer()
   On Error Resume Next
   Dim pos As Single
   Dim iSecs As Integer
   Dim iMins As Integer
   Dim iPos As Integer
   Dim TimeElapsedPerc As Long
   Dim STime As String
   
   On Error Resume Next
      
   If (BASS_ChannelIsActive(chan(2)) = 0) Then
     pos = -1 ' reached the end
   Else
     'Get current possition of playing...
     pos = Format(bassTime.GetPlayingPos(chan(2)), "0")
   End If
   
   
   'Check if END Reached. Stop timer, stop playing and reset button...
   If pos = -1 Then
     TimerP2.Enabled = False
     StopSong Player2Index
     ResetButton Player2Index, True
      If iAutoAdvance = 2 Then
         StartNextAvalableSong (Player2Index)
      ElseIf iAutoAdvance = 3 Then
      
      End If
     Exit Sub
   End If
   
   'Calculate the progress bar's position, as well as the time left to display
   TimeElapsedPerc = Val(pos) * 100 / Val(Duration(2))
'''   '==================================================================================
'''   'For Demo system, only allow load of 5 songs
'''   If DemoFlag Then
'''     ' DemoCnt = DemoCnt + 1
'''      If Round(pos) > DemoTime Then
'''         TimerP2.Enabled = False
'''         StopSong Player2Index
'''         ResetButton Player2Index, True
'''         MsgBox DemoMsg2, vbExclamation, DemoHeading
'''         Exit Sub
'''      End If
'''   End If
'''   '==================================================================================
   'TimeElapsedPerc = Val(pos) * 100 / Val(Duration(2))
   'Show progress bar value
   'sspProgress(Player2Index).ProValue = TimeElapsedPerc
   sspProgress(Player2Index).value = TimeElapsedPerc
   'sspProgress(Player2Index).FloodPercent = TimeElapsedPerc

   iSecs = Round(pos)  'Get the seconds rounded
   STime = ConvertSecondsToTime(iSecs)
   lblTimeLeft(Player2Index).Caption = STime
   STime = ""
   STime = Right(bassTime.GetTime(Duration(2) - pos), 5)
   STime = Format(Left(STime, 2), "00") & Right(STime, 3)
   
   'Show Time left
   lblTimePlayed(Player2Index).Caption = STime
   
   If Left(STime, 2) = "00" Then
     If Val(Right(STime, 2)) < 10 Then
       TimerF2.Enabled = True
     End If
   End If
   
Exit Sub
err1:
MsgBox "Error in Module : TimerP2_Timer " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub TimerP2Level_Timer()

   Dim LevelInd As Long
   Dim LevelAve As Long
   Dim LeftChan As Integer
   Dim RightChan As Integer
   Dim StereoLevel As Single
   
   On Error Resume Next
   
   Dim Levels As Single
   
   LevelInd = BASS_ChannelGetLevel(chan(2))
   LeftChan = Round(LoWord(LevelInd) * iVol2)
   RightChan = Round(HiWord(LevelInd) * iVol2)
   Left2Chan = LeftChan
   Right2Chan = RightChan
   
   SetPeakLevel LeftChan, RightChan, Player2Index
   
End Sub

Private Sub timFade_Timer()
   
   On Error Resume Next
   Stop1 = True
   SetFadeColor LastPlaying
   
End Sub

Private Sub tmFade_Timer()

On Error Resume Next

If fVolume < 0.1 Then
   tmFade.Enabled = False
   timFade.Enabled = False
   StopSong LastPlaying
   ResetButton LastPlaying, True
   cmdFadeOut.Enabled = True
   cmdFadeOut.BackColor = vbBlack
   cmdFadeOut.ForeColor = vbDirectionColor
   cmdFadeOut.Picture = LoadPicture(App.Path & "\tmpFade")
   'InitialisePeaks 0
   Command2.SetFocus
   Exit Sub
End If

Select Case fVolume
   Case Is < 0.5
      fVolume = fVolume - (Fadeout / 50)
   Case Is < 5
      fVolume = fVolume - (Fadeout / 25)
'   Case Is < 10
'      fVolume = fVolume - (Fadeout / 8)
   Case Is < 20
      fVolume = fVolume - (Fadeout / 7)
'   Case Is < 30
'      fVolume = fVolume - (Fadeout / 6)
   Case Is < 40
      fVolume = fVolume - (Fadeout / 5)
'   Case Is < 50
'      fVolume = fVolume - (Fadeout / 4)
'   Case Is < 60
'      fVolume = fVolume - (Fadeout / 3)
'   Case Is < 70
'      fVolume = fVolume - (Fadeout / 2)
   Case Else
      fVolume = fVolume - Fadeout
End Select

'If fVolume < 0.5 Then
'   fVolume = fVolume - (Fadeout / 60)
'ElseIf fVolume < 5 Then
'   fVolume = fVolume - (Fadeout / 30)
''''ElseIf fVolume < 10 Then
''''   fVolume = fVolume - (Fadeout / 10)
'ElseIf fVolume < 20 Then
'   fVolume = fVolume - (Fadeout / 5)
'Else
'   fVolume = fVolume - Fadeout
'End If

SetSongVolume LastPlaying

End Sub

Private Sub tmrBass_Timer()
''''   On Error Resume Next
''''   'pgCpu.value = Format(BASS_GetCPU, "0.00") * 28
''''   sspCpu.FloodPercent = Format(BASS_GetCPU, "0.00") * 28
''''   'sspCpu.FloodType = ssLeftToRight
''''
''''   If sspCpu.FloodPercent > 95 Then
''''      sspCpu.FloodColor = vbRed
''''     ' Change_pb_ForeColor pgCpu.hWnd, vbRed
''''   ElseIf sspCpu.FloodPercent > 80 Then
''''      sspCpu.FloodColor = vbYellow
''''      'Change_pb_ForeColor pgCpu.hWnd, vbYellow 'vbOrange
''''   ElseIf sspCpu.FloodPercent > 60 Then
''''      sspCpu.FloodColor = vbGreen
''''      'Change_pb_ForeColor pgCpu.hWnd, vbGreen  'vbYellow
''''   Else
''''      sspCpu.FloodColor = &HE8926C
''''      'Change_pb_ForeColor pgCpu.hWnd, &HE8926C      'vbBlue
''''   End If

End Sub

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  ProcessClickOK
End If
End Sub
