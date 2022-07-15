VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetupSoundCards 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   10980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetupSoundCards.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetupSoundCards.frx":000C
   ScaleHeight     =   10980
   ScaleWidth      =   20445
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvSoundCards 
      Height          =   2010
      Left            =   3330
      TabIndex        =   2
      Top             =   2820
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   3545
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   405
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1365
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   714
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   7104768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "                                                       Configuration Settings"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   17190
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetupSoundCards.frx":3551
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetupSoundCards.frx":A785
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetupSoundCards.frx":13BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetupSoundCards.frx":144CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetupSoundCards.frx":14DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetupSoundCards.frx":15709
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   8730
      Left            =   30
      TabIndex        =   1
      Top             =   1770
      Width           =   20310
      _ExtentX        =   35825
      _ExtentY        =   15399
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
      Caption         =   "SSPanel2"
      BevelOuter      =   0
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   6090
         MaxLength       =   20
         TabIndex        =   42
         Top             =   7785
         Width           =   2880
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   615
         Left            =   3390
         TabIndex        =   4
         Top             =   4305
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   1085
         _Version        =   131074
         BackColor       =   14993249
         BevelOuter      =   0
         Begin Threed.SSPanel sspAutoAdvance 
            Height          =   555
            Index           =   2
            Left            =   1305
            TabIndex        =   5
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Auto  Advance"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspAutoAdvance 
            Height          =   555
            Index           =   1
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Play And Stop"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspAutoAdvance 
            Height          =   555
            Index           =   3
            Left            =   2580
            TabIndex        =   49
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "STOP Current  PLAY NEXT"
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   615
         Left            =   3390
         TabIndex        =   7
         Top             =   5400
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1085
         _Version        =   131074
         BackColor       =   14993249
         BorderWidth     =   1
         BevelOuter      =   0
         Begin Threed.SSPanel sspAutoRemoveSilence 
            Height          =   555
            Index           =   1
            Left            =   30
            TabIndex        =   8
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Remove  Silences"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspAutoRemoveSilence 
            Height          =   555
            Index           =   2
            Left            =   1305
            TabIndex        =   9
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Restore  Silences"
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   615
         Left            =   9345
         TabIndex        =   10
         Top             =   5400
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1085
         _Version        =   131074
         BackColor       =   14993249
         BevelOuter      =   0
         Begin Threed.SSPanel sspButtonShowPlayArea 
            Height          =   555
            Index           =   2
            Left            =   1305
            TabIndex        =   11
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Dual  Stream"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspButtonShowPlayArea 
            Height          =   555
            Index           =   1
            Left            =   30
            TabIndex        =   12
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Single Stream"
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel sspButtonDirMain 
         Height          =   615
         Left            =   9345
         TabIndex        =   13
         Top             =   4305
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1085
         _Version        =   131074
         BackColor       =   14993249
         BevelOuter      =   0
         Begin Threed.SSPanel sspButtonDirection 
            Height          =   555
            Index           =   1
            Left            =   1305
            TabIndex        =   14
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "TOP to BOTTOM"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspButtonDirection 
            Height          =   555
            Index           =   2
            Left            =   30
            TabIndex        =   15
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "LEFT to RIGHT"
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   615
         Left            =   9360
         TabIndex        =   20
         Top             =   7695
         Visible         =   0   'False
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1085
         _Version        =   131074
         BackColor       =   14993249
         BevelOuter      =   0
         Begin Threed.SSPanel sspOption3 
            Height          =   555
            Index           =   3
            Left            =   1680
            TabIndex        =   21
            Top             =   30
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   979
            _Version        =   131074
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "25"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspOption3 
            Height          =   555
            Index           =   2
            Left            =   855
            TabIndex        =   22
            Top             =   30
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   979
            _Version        =   131074
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "16"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspOption3 
            Height          =   555
            Index           =   1
            Left            =   30
            TabIndex        =   23
            Top             =   30
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   979
            _Version        =   131074
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "9"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspOption3 
            Height          =   555
            Index           =   4
            Left            =   2505
            TabIndex        =   24
            Top             =   30
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   979
            _Version        =   131074
            ForeColor       =   15194953
            BackColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "30"
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   9345
         TabIndex        =   33
         Top             =   6600
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1085
         _Version        =   131074
         BackColor       =   14993249
         BorderWidth     =   1
         BevelOuter      =   0
         Begin Threed.SSPanel sspDefaultColors 
            Height          =   555
            Index           =   2
            Left            =   30
            TabIndex        =   34
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "  Random   Colours"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspDefaultColors 
            Height          =   555
            Index           =   1
            Left            =   1305
            TabIndex        =   35
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   " Default   Colours"
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   615
         Left            =   3390
         TabIndex        =   37
         Top             =   7665
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1085
         _Version        =   131074
         BackColor       =   14993249
         BorderWidth     =   1
         BevelOuter      =   0
         Begin Threed.SSPanel sspDefaultSecure 
            Height          =   555
            Index           =   1
            Left            =   30
            TabIndex        =   38
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Non Secure"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspDefaultSecure 
            Height          =   555
            Index           =   2
            Left            =   1305
            TabIndex        =   39
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "SECURE"
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   615
         Left            =   3420
         TabIndex        =   50
         Top             =   6600
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1085
         _Version        =   131074
         BackColor       =   14993249
         BevelOuter      =   0
         Begin Threed.SSPanel sspButtonAdjustVol 
            Height          =   555
            Index           =   2
            Left            =   1305
            TabIndex        =   51
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Set Volume to 100%"
            BevelOuter      =   0
         End
         Begin Threed.SSPanel sspButtonAdjustVol 
            Height          =   555
            Index           =   1
            Left            =   30
            TabIndex        =   52
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            _Version        =   131074
            CaptionStyle    =   1
            ForeColor       =   15194953
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
            Caption         =   "Normalize Volume to 75%"
            BevelOuter      =   0
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Set Initial Volume"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   8
         Left            =   3420
         TabIndex        =   53
         Top             =   6315
         Width           =   4740
      End
      Begin Threed.SSCommand sspDefCol 
         Height          =   300
         Index           =   5
         Left            =   12780
         TabIndex        =   48
         Top             =   6915
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         _Version        =   131074
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "6"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand sspDefCol 
         Height          =   300
         Index           =   4
         Left            =   12420
         TabIndex        =   47
         Top             =   6915
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         _Version        =   131074
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "5"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand sspDefCol 
         Height          =   300
         Index           =   3
         Left            =   12060
         TabIndex        =   46
         Top             =   6915
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         _Version        =   131074
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "4"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand sspDefCol 
         Height          =   300
         Index           =   1
         Left            =   12420
         TabIndex        =   45
         Top             =   6600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         _Version        =   131074
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "2"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand sspDefCol 
         Height          =   300
         Index           =   2
         Left            =   12780
         TabIndex        =   44
         Top             =   6600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "3"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand sspDefCol 
         Height          =   300
         Index           =   0
         Left            =   12060
         TabIndex        =   43
         Top             =   6600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         _Version        =   131074
         BackColor       =   16768505
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin VB.Label lblPlugins 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   7620
         Left            =   15660
         TabIndex        =   32
         Top             =   450
         Width           =   4290
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Password..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Index           =   6
         Left            =   6090
         TabIndex        =   41
         Top             =   7530
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Secure Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   7
         Left            =   3420
         TabIndex        =   40
         Top             =   7380
         Width           =   3240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Default / Random Button Colours"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   6
         Left            =   9345
         TabIndex        =   36
         Top             =   6315
         Width           =   3240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Plug-ins (Music Formats) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   5
         Left            =   15390
         TabIndex        =   31
         Top             =   180
         Width           =   3900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   3300
         X2              =   15015
         Y1              =   3225
         Y2              =   3225
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Player configuration Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Left            =   3330
         TabIndex        =   26
         Top             =   3405
         Width           =   9690
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Player button layout"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   4
         Left            =   9360
         TabIndex        =   25
         Top             =   7410
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Remove leading silences"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   3
         Left            =   3390
         TabIndex        =   19
         Top             =   5115
         Width           =   3240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Player options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   2
         Left            =   3390
         TabIndex        =   18
         Top             =   4020
         Width           =   3240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Allow 1 or 2 streams of music to play simultaniously"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   1
         Left            =   9345
         TabIndex        =   17
         Top             =   5115
         Width           =   4740
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Player button layout Direction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   0
         Left            =   9345
         TabIndex        =   16
         Top             =   4020
         Width           =   3240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   3240
         X2              =   14955
         Y1              =   8715
         Y2              =   8715
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSetupSoundCards.frx":15EAF
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   600
         Left            =   3345
         TabIndex        =   3
         Top             =   255
         Width           =   11460
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   3255
         X2              =   14970
         Y1              =   195
         Y2              =   195
      End
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   885
      Left            =   18150
      TabIndex        =   27
      Top             =   180
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1561
      _Version        =   131074
      BackColor       =   3092271
      BevelOuter      =   0
      Begin Threed.SSCommand cmdOk 
         Height          =   810
         Left            =   45
         TabIndex        =   29
         Top             =   45
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1429
         _Version        =   131074
         ForeColor       =   15194953
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
         Caption         =   "OK"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   810
         Left            =   975
         TabIndex        =   28
         Top             =   45
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1429
         _Version        =   131074
         ForeColor       =   15194953
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
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
      Height          =   135
      Left            =   990
      TabIndex        =   30
      Top             =   735
      Width           =   2205
   End
End
Attribute VB_Name = "frmSetupSoundCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iPreviousSelection As Integer
Const SelectedChar As String = "R"

Private Sub CancelButton_Click()
If bSinglePlayer Then
  lDeviceSingle = -99
Else
  lDeviceNo = -99
End If
Unload Me

End Sub

Private Sub cmdExit_Click()

'Get the defaults, in case we cancel
iSecureMode = Val(GetSetting(regMainKey, regSubKey, "SecureMode"))
If iSecureMode = 0 Then iSecureMode = 1   '1 is the default (OFF)
If iSecureMode = 1 Then
  sSecurePWD = ""
Else
  sSecurePWD = GetSetting(regMainKey, regSubKey, "SecureModePWD")
End If


If bSinglePlayer Then
  lDeviceSingle = -99
Else
  lDeviceNo = -99
End If
Unload Me
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.BackColor = &HE7DB49
cmdExit.ForeColor = vbBlack
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.ForeColor = &HE7DB49
cmdExit.BackColor = vbBlack
End Sub

Function ValidatePassword() As Boolean
Dim bNumber As Boolean
Dim bAlpha As Boolean
Dim bSpecial As Boolean
Dim bUpper As Boolean
Dim bLower As Boolean
Dim bLength As Boolean
Dim sPWD As String
sPWD = Trim(txtPassword.text)

If Len(sPWD) >= 5 Then bLength = True

bSpecial = True

For i = 1 To Len(sPWD)
  If IsNumeric(Mid(sPWD, i, 1)) Then bNumber = True
'  If Mid(sPWD, I, 1) = "!" Or _
'     Mid(sPWD, I, 1) = "@" Or _
'     Mid(sPWD, I, 1) = "#" Or _
'     Mid(sPWD, I, 1) = "$" Or _
'     Mid(sPWD, I, 1) = "%" Or _
'     Mid(sPWD, I, 1) = "^" Or _
'     Mid(sPWD, I, 1) = "&" Or _
'     Mid(sPWD, I, 1) = "*" Or _
'     Mid(sPWD, I, 1) = "(" Or _
'     Mid(sPWD, I, 1) = ")" Then bSpecial = True
  
  Select Case Asc(Mid(sPWD, i, 1))
    Case 65 To 90 'Uppercase
      bUpper = True
    Case 97 To 122 'Lowercase
      bLower = True
    Case 48 To 57 'Numbers
      bNumber = True
  End Select
Next i

If bNumber And bUpper And bLower And bSpecial And bLength Then ValidatePassword = True

End Function

Private Sub cmdOK_Click()

If iSecureMode = 2 Then
  If Trim(txtPassword.text) <> "" Then
    If ValidatePassword Then
      sSecurePWD = Trim(txtPassword.text)
      SaveSetting regMainKey, regSubKey, "SecureMode", iSecureMode
      SaveSetting regMainKey, regSubKey, "SecureModePWD", sSecurePWD
    Else
      MsgBox "Password Does not conform to Validation Rules." & Chr(13) & Chr(13) & "Password must contain at least 1 of each of the following:" & Chr(13) & vbTab & "At least 5 characters" & Chr(13) & vbTab & "At lease 1 Uppercase and 1 Lowercase alpha Character" & Chr(13) & vbTab & "NUMBER Character", vbInformation, "Password ERROR"
      'MsgBox "Password Does not conform to Validation Rules." & Chr(13) & Chr(13) & "Password must contain at least 1 of each of the following:" & Chr(13) & vbTab & "At least 5 characters" & Chr(13) & vbTab & "At lease 1 Uppercase and 1 Lowercase alpha Character" & Chr(13) & vbTab & "NUMBER Character" & Chr(13) & vbTab & "SPECIAL Character", vbInformation, "Password ERROR"
      txtPassword.SetFocus
      Exit Sub
    End If
  Else
    MsgBox "Please enter a PASSWORD for SECURE MODE.", vbExclamation, "Password Error"
    txtPassword.SetFocus
    Exit Sub
  End If
Else
  sSecurePWD = ""
  SaveSetting regMainKey, regSubKey, "SecureMode", iSecureMode
  SaveSetting regMainKey, regSubKey, "SecureModePWD", sSecurePWD
End If

'Save all other settings...
SaveSetting regMainKey, regSubKey, "AdjustVolume", iAdjustVol
SaveSetting regMainKey, regSubKey, "ButtonDirection", iButtonDirection
SaveSetting regMainKey, regSubKey, "PlayStopPause", iButtonPlayStopPause
SaveSetting regMainKey, regSubKey, "AutoAdvance", iAutoAdvance
SaveSetting regMainKey, regSubKey, "Remove Silence", iButtonRemoveSilence
SaveSetting regMainKey, regSubKey, "Streams", iButtonStreams
SaveSetting regMainKey, regSubKey, "ButtonColor", iButtonDefaultColor
'SaveSetting regMainKey, regSubKey, "SecureMode", iSecureMode
'SaveSetting regMainKey, regSubKey, "SecureModePWD", sSecurePWD
SaveSetting regMainKey, regSubKey, "Max Buttons", iButtonMaxSelected

'iPreviousSelection = lvSoundCards.SelectedItem.Index
   
If bSinglePlayer Then
  If lDeviceSingle <> 0 Then Unload Me
Else
  If lDeviceNo <> 0 Then Unload Me
End If

End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOk.BackColor = &HE7DB49
cmdOk.ForeColor = vbBlack
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOk.ForeColor = &HE7DB49
cmdOk.BackColor = vbBlack
End Sub

Private Sub Form_Activate()

'EnableCloseButton Me.hWnd, False
SetDefaultDevice
'If lDeviceNo <> -99 Then
'  lvSoundCards.ListItems(lDeviceNo).Selected = True
'Else
'  lvSoundCards.ListItems(1).Selected = True
'End If



'If lvSoundCards.ListItems.Count > 0 Then
''   Call SetLVSubImages(lvSoundCards, iPreviousSelection, 2, 0, True)
'   Call SetLVSubImages(lvSoundCards, lvSoundCards.SelectedItem.Index, 2, 1, True)
'   iPreviousSelection = lvSoundCards.SelectedItem.Index
'
'  If bSinglePlayer Then
'    lDeviceSingle = Val(Mid(lvSoundCards.SelectedItem.key, 2))
'  Else
'    lDeviceNo = Val(Mid(lvSoundCards.SelectedItem.key, 2))
'  End If
'
'End If

If BusyPlaying Then
   lvSoundCards.Enabled = False
   SSPanel5.Enabled = False
   sspButtonDirMain.Enabled = False
Else
   'lvSoundCards.SetFocus
   cmdExit.SetFocus
End If

If iSecureMode = 1 Then
  Label3(6).Visible = False
  txtPassword.Visible = False
Else 'Secure Mode active
  Label3(6).Visible = True
  txtPassword.Visible = True
End If




End Sub

Sub SetMaxButtonLayout()

   For i = 1 To 4
      If i = iButtonMaxSelected Then
         sspOption3(i).BackColor = vbDirectionColor
         sspOption3(i).ForeColor = vbBlack
         iMaxBut = Val(sspOption3(i).Caption)
      Else
         sspOption3(i).BackColor = vbBlack
         sspOption3(i).ForeColor = vbDirectionColor
      End If
   Next i
   
End Sub

Sub SetButtonStreams(Index As Integer)

   If Index = 1 Then
      sspButtonShowPlayArea(2).BackColor = vbBlack
      sspButtonShowPlayArea(1).BackColor = vbDirectionColor
      sspButtonShowPlayArea(2).ForeColor = vbDirectionColor
      sspButtonShowPlayArea(1).ForeColor = vbBlack
   Else
      sspButtonShowPlayArea(1).BackColor = vbBlack
      sspButtonShowPlayArea(2).BackColor = vbDirectionColor
      sspButtonShowPlayArea(1).ForeColor = vbDirectionColor
      sspButtonShowPlayArea(2).ForeColor = vbBlack
   End If
   
End Sub

Sub SetAutoRemoveSilences(Index As Integer)

   If Index = 1 Then
      sspAutoRemoveSilence(2).BackColor = vbBlack
      sspAutoRemoveSilence(1).BackColor = vbDirectionColor
      sspAutoRemoveSilence(2).ForeColor = vbDirectionColor
      sspAutoRemoveSilence(1).ForeColor = vbBlack
   Else
      sspAutoRemoveSilence(1).BackColor = vbBlack
      sspAutoRemoveSilence(2).BackColor = vbDirectionColor
      sspAutoRemoveSilence(1).ForeColor = vbDirectionColor
      sspAutoRemoveSilence(2).ForeColor = vbBlack
   End If
   
End Sub

Sub SetAdjustVolume(Index As Integer)

   If Index = 1 Then
      sspButtonAdjustVol(2).BackColor = vbBlack
      sspButtonAdjustVol(1).BackColor = vbDirectionColor
      sspButtonAdjustVol(2).ForeColor = vbDirectionColor
      sspButtonAdjustVol(1).ForeColor = vbBlack
   Else
      sspButtonAdjustVol(1).BackColor = vbBlack
      sspButtonAdjustVol(2).BackColor = vbDirectionColor
      sspButtonAdjustVol(1).ForeColor = vbDirectionColor
      sspButtonAdjustVol(2).ForeColor = vbBlack
   End If
   
End Sub

Sub SetAutoAdvance(Index As Integer)

Dim i As Integer

For i = 1 To 3
  sspAutoAdvance(i).ForeColor = vbDirectionColor
  sspAutoAdvance(i).BackColor = vbBlack
Next i

sspAutoAdvance(Index).BackColor = vbDirectionColor
sspAutoAdvance(Index).ForeColor = vbBlack
   
End Sub

Sub SetDefaultColors(Index As Integer)

   If Index = 1 Then
      sspDefaultColors(2).BackColor = vbBlack
      sspDefaultColors(1).BackColor = vbDirectionColor
      sspDefaultColors(2).ForeColor = vbDirectionColor
      sspDefaultColors(1).ForeColor = vbBlack
   Else
      sspDefaultColors(1).BackColor = vbBlack
      sspDefaultColors(2).BackColor = vbDirectionColor
      sspDefaultColors(1).ForeColor = vbDirectionColor
      sspDefaultColors(2).ForeColor = vbBlack
   End If
   
End Sub

Sub SetSecure(Index As Integer)

   If Index = 1 Then
      sspDefaultSecure(2).BackColor = vbBlack
      sspDefaultSecure(1).BackColor = vbDirectionColor
      sspDefaultSecure(2).ForeColor = vbDirectionColor
      sspDefaultSecure(1).ForeColor = vbBlack
   Else
      sspDefaultSecure(1).BackColor = vbBlack
      sspDefaultSecure(2).BackColor = vbDirectionColor
      sspDefaultSecure(1).ForeColor = vbDirectionColor
      sspDefaultSecure(2).ForeColor = vbBlack
   End If
   
End Sub

Private Sub sspAutoAdvance_Click(Index As Integer)

On Error GoTo err1
Dim i As Integer

If sspAutoAdvance(Index).BackColor = vbDirectionColor Then Exit Sub

For i = 1 To 3
  sspAutoAdvance(i).ForeColor = vbDirectionColor
  sspAutoAdvance(i).BackColor = vbBlack
Next i

sspAutoAdvance(Index).BackColor = vbDirectionColor
sspAutoAdvance(Index).ForeColor = vbBlack


iAutoAdvance = Index

Exit Sub
err1:
MsgBox "Error in Module : sspAutoAdvance_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub sspAutoRemoveSilence_Click(Index As Integer)

On Error GoTo err1

If sspAutoRemoveSilence(Index).BackColor = vbDirectionColor Then Exit Sub

If sspAutoRemoveSilence(1).BackColor = vbBlack Then
   sspAutoRemoveSilence(2).BackColor = vbBlack
   sspAutoRemoveSilence(1).BackColor = vbDirectionColor
   sspAutoRemoveSilence(2).ForeColor = vbDirectionColor
   sspAutoRemoveSilence(1).ForeColor = vbBlack
Else
   sspAutoRemoveSilence(1).BackColor = vbBlack
   sspAutoRemoveSilence(2).BackColor = vbDirectionColor
   sspAutoRemoveSilence(1).ForeColor = vbDirectionColor
   sspAutoRemoveSilence(2).ForeColor = vbBlack
End If

iButtonRemoveSilence = Index

Exit Sub
err1:
MsgBox "Error in Module : sspAutoRemoveSilence_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub sspButtonDirection_Click(Index As Integer)
On Error GoTo err1

'Change the button Direction ...
If sspButtonDirection(1).BackColor = vbBlack Then
   sspButtonDirection(2).BackColor = vbBlack
   sspButtonDirection(1).BackColor = vbDirectionColor
   sspButtonDirection(2).ForeColor = vbDirectionColor
   sspButtonDirection(1).ForeColor = vbBlack
Else
   sspButtonDirection(1).BackColor = vbBlack
   sspButtonDirection(2).BackColor = vbDirectionColor
   sspButtonDirection(1).ForeColor = vbDirectionColor
   sspButtonDirection(2).ForeColor = vbBlack
End If

'Get the correct layout indicator
iButtonDirection = Index

Exit Sub
err1:
MsgBox "Error in Module : sspButtonShowPlayArea_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub sspButtonShowPlayArea_Click(Index As Integer)

On Error GoTo err1

If sspButtonShowPlayArea(Index).BackColor = vbDirectionColor Then Exit Sub

If sspButtonShowPlayArea(1).BackColor = vbBlack Then
   sspButtonShowPlayArea(2).BackColor = vbBlack
   sspButtonShowPlayArea(1).BackColor = vbDirectionColor
   sspButtonShowPlayArea(2).ForeColor = vbDirectionColor
   sspButtonShowPlayArea(1).ForeColor = vbBlack
   'ShowButtonStreams True
Else
   sspButtonShowPlayArea(1).BackColor = vbBlack
   sspButtonShowPlayArea(2).BackColor = vbDirectionColor
   sspButtonShowPlayArea(1).ForeColor = vbDirectionColor
   sspButtonShowPlayArea(2).ForeColor = vbBlack
   'ShowButtonStreams False
End If
iButtonStreams = Index

Exit Sub
err1:
MsgBox "Error in Module : sspButtonShowPlayArea_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub


Private Sub Form_Load()

lvSoundCards.View = lvwReport
lvSoundCards.ColumnHeaders.Add , , "Available Devices ", 10000
lvSoundCards.ColumnHeaders.Add , , "Spare ", 500

''Me.Picture = LoadPicture(App.Path & "\tmpBanner")

lblVersion.Caption = "Version   :    " & App.Major & "." & App.Minor & "." & App.Revision
   lblVersion.Top = 750
   lblVersion.Left = 255
   lblVersion.Width = 2940
   lblVersion.FontSize = 7

'''HelpContextID = hlpSetupsound

If bSinglePlayer Then
  lDeviceSingle = -99
Else
  lDeviceNo = -99
End If

LoadDevices

lblPlugins.Caption = ListOfPlugins

For i = 1 To 4
   sspOption3(i).ForeColor = vbDirectionColor
   sspOption3(i).Enabled = False
Next i

SetCorrectLayoutButton iButtonDirection
SetAutoAdvance iAutoAdvance
SetAutoRemoveSilences iButtonRemoveSilence
SetButtonStreams iButtonStreams
SetButtonsCount iButtonMaxSelected
SetDefaultColors iButtonDefaultColor
SetSecure iSecureMode
SetDefaultColorsTxt vbNDefault
SetDefaultColorsVisible
SetAdjustVolume iAdjustVol

txtPassword.text = sSecurePWD

'Label3(0).ForeColor = IIf(BusyPlaying, &H808080, &HFFFF80)
'Label3(4).ForeColor = IIf(BusyPlaying, &H808080, &HFFFF80)

''Me.Width = 20445
''Me.Height = 11010
''Me.Top = 0
''Me.Left = 0

Label2(4).Top = 15000
SSPanel5.Top = 15000

Me.Width = frmPlayer.Width
Me.Height = frmPlayer.Height
Me.Top = frmPlayer.Top
Me.Left = frmPlayer.Left

SSPanel2.Height = Me.Height
SSPanel2.Width = Me.Width + 300

'SetDefaultDevice

End Sub

Sub SetDefaultColorsTxt(DefaultColor As Long)
Dim i As Integer

    
For i = 0 To 5
  If sspDefCol(i).BackColor = vbNDefault Then
    sspDefCol(i).Caption = "X"
    sspDefCol(i).Font.Bold = True
    sspDefCol(i).Font.Size = 12
  Else
    sspDefCol(i).Caption = CStr(i + 1) 'Chr(i + 106)  '106
    sspDefCol(i).Font.Bold = False
    sspDefCol(i).Font.Size = 10
  End If
Next i

End Sub

Sub SetButtonsCount(Index As Integer)
Dim i As Integer

For i = 1 To 4
   If i = Index Then
      sspOption3(i).BackColor = IIf(BusyPlaying, &H808080, vbDirectionColor)
      sspOption3(i).ForeColor = vbBlack
   Else
      sspOption3(i).BackColor = vbBlack
      sspOption3(i).ForeColor = IIf(BusyPlaying, &H808080, vbDirectionColor)
   End If
Next i


End Sub

Sub SetCorrectLayoutButton(iLayout As Integer)

On Error GoTo err1
'Set the correct button back color to the inverse so the next code will fix it...
sspButtonDirection(iLayout).BackColor = vbBlack

'Change the button colors...
If iLayout = 1 Then
'If sspButtonDirection(0).BackColor = vbBlack Then
   sspButtonDirection(2).BackColor = vbBlack
   sspButtonDirection(1).BackColor = IIf(BusyPlaying, &H808080, vbDirectionColor)
   sspButtonDirection(2).ForeColor = IIf(BusyPlaying, &H808080, vbDirectionColor)
   sspButtonDirection(1).ForeColor = vbBlack
Else
   sspButtonDirection(1).BackColor = vbBlack
   sspButtonDirection(2).BackColor = IIf(BusyPlaying, &H808080, vbDirectionColor)
   sspButtonDirection(1).ForeColor = IIf(BusyPlaying, &H808080, vbDirectionColor)
   sspButtonDirection(2).ForeColor = vbBlack
End If

Exit Sub
err1:

MsgBox "Error in Module : SetCorrectLayoutButton " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Sub SetDefaultDevice()
Dim iRow As Integer
On Error Resume Next

If bSinglePlayer Then
  lDeviceSingle = GetSetting(regMainKey, regSubKey, "SinglePlay Device")
  If lDeviceNo = 0 Then lDeviceNo = -99
  If lDeviceSingle <> -99 Then
    For iRow = 1 To lvSoundCards.ListItems.Count
      If Val(Mid(lvSoundCards.ListItems(iRow).key, 2)) = lDeviceSingle Then
        'lvSoundCards.ListItems(iRow).Selected = True
        Call SetLVSubImages(lvSoundCards, iRow, 1, 1, True)
        Exit For
      End If
    Next iRow
  End If
Else
  lDeviceNo = GetSetting(regMainKey, regSubKey, "Current Device")
  If lDeviceNo = 0 Then lDeviceNo = -99
  If lDeviceNo <> -99 Then
    For iRow = 1 To lvSoundCards.ListItems.Count
      If Val(Mid(lvSoundCards.ListItems(iRow).key, 2)) = lDeviceNo Then
        'lvSoundCards.ListItems(iRow).Selected = True
         Call SetLVSubImages(lvSoundCards, iRow, 1, 1, True)
         Exit For
      End If
    Next iRow
  End If
End If


If lDeviceNo <> -99 Then
  lvSoundCards.ListItems(lDeviceNo).Selected = True
Else
  lvSoundCards.ListItems(1).Selected = True
End If

lvSoundCards.SelectedItem.EnsureVisible


End Sub

Sub LoadDevices()
Dim iDeviceCnt As Long

On Error Resume Next

'For iDeviceCnt = 0 To BASS_GetDeviceCount - 1
'  'sDevices = BASS_GetDeviceDescriptionString(iDeviceCnt)
'  Set mItem = lvSoundCards.ListItems.Add(, "~" & CStr(iDeviceCnt), BASS_GetDeviceDescriptionString(iDeviceCnt), 0, 0)
'  mItem.SubItems(1) = ""
'Next iDeviceCnt

Dim c As Integer
Dim i As BASS_DEVICEINFO
Dim iItem As Integer
Dim iDev As Integer

c = 1      ' device 1 = 1st real device
iItem = 0
iDev = 2
While BASS_GetDeviceInfo(c, i)
  If (i.Flags And BASS_DEVICE_ENABLED) Then  ' enabled, so add it...
      iDev = iDev + 1
      If iDev > 6 Then iDev = 6
      'lstDevices.AddItem VBStrFromAnsiPtr(i.name)
      Set mItem = lvSoundCards.ListItems.Add(, "~" & CStr(c), "   " & VBStrFromAnsiPtr(i.name), 0, iDev)
      iItem = iItem + 1
      Call SetLVSubImages(lvSoundCards, iItem, 1, 0, True)
  End If
  c = c + 1
Wend

'      iDev = iDev + 1
'      iItem = iItem + 1
'      Set mItem = lvSoundCards.ListItems.Add(, "~" & "1", "   Test Device 1", 0, iDev)
'      Call SetLVSubImages(lvSoundCards, iItem, 1, 0, True)
'
'      iDev = iDev + 1
'      iItem = iItem + 1
'      Set mItem = lvSoundCards.ListItems.Add(, "~" & "2", "   Test Device 2", 0, iDev)
'      Call SetLVSubImages(lvSoundCards, iItem, 1, 0, True)
'
'      iDev = iDev + 1
'      iItem = iItem + 1
'      Set mItem = lvSoundCards.ListItems.Add(, "~" & "3", "   Test Device 3", 0, iDev)
'      Call SetLVSubImages(lvSoundCards, iItem, 1, 0, True)
'
'      iDev = 6
'      iItem = iItem + 1
'      Set mItem = lvSoundCards.ListItems.Add(, "~" & "4", "   Test Device 4", 0, iDev)
'      Call SetLVSubImages(lvSoundCards, iItem, 1, 0, True)
'
'      iDev = 6
'      iItem = iItem + 1
'      Set mItem = lvSoundCards.ListItems.Add(, "~" & "5", "   Test Device 5", 0, iDev)
'      Call SetLVSubImages(lvSoundCards, iItem, 1, 0, True)
   
End Sub

Private Sub Form_Resize()

If Me.WindowState <> 1 Then
   If ApplyStandardTheme Then
      Me.Width = frmPlayer.Width - 120
      Me.Height = frmPlayer.Height - 120
      Me.Top = frmPlayer.Top + 60
      Me.Left = frmPlayer.Left + 60
   Else
      Me.Width = frmPlayer.Width - 180
      Me.Height = frmPlayer.Height - 180
      Me.Top = frmPlayer.Top + 90
      Me.Left = frmPlayer.Left + 90
   End If
End If

End Sub

Private Sub lvSoundCards_Click()

On Error Resume Next

If lvSoundCards.ListItems.Count > 0 Then
  If bSinglePlayer Then
    lDeviceSingle = Val(Mid(lvSoundCards.SelectedItem.key, 2))
  Else
    lDeviceNo = Val(Mid(lvSoundCards.SelectedItem.key, 2))
  End If
    
  ResetSoundCardSelection
  Call SetLVSubImages(lvSoundCards, lvSoundCards.SelectedItem.Index, 1, 1, True)
  iPreviousSelection = lvSoundCards.SelectedItem.Index
  
  If bSinglePlayer Then
    SaveSetting regMainKey, regSubKey, "SinglePlay Device", lDeviceSingle
    SaveSetting regMainKey, regSubKey, "SinglePlay Device Description", lvSoundCards.SelectedItem
  Else
    SaveSetting regMainKey, regSubKey, "Current Device", lDeviceNo
    SaveSetting regMainKey, regSubKey, "Current Device Description", lvSoundCards.SelectedItem
  End If
End If

End Sub

Sub ResetSoundCardSelection()
Dim iRow As Long

For iRow = 1 To lvSoundCards.ListItems.Count
  Call SetLVSubImages(lvSoundCards, iRow, 1, 0, True)
Next iRow

End Sub

Private Sub sspDefaultColors_Click(Index As Integer)
On Error GoTo err1
Dim i As Integer

If sspDefaultColors(Index).BackColor = vbDirectionColor Then Exit Sub

If sspDefaultColors(1).BackColor = vbBlack Then
   sspDefaultColors(2).BackColor = vbBlack
   sspDefaultColors(1).BackColor = vbDirectionColor
   sspDefaultColors(2).ForeColor = vbDirectionColor
   sspDefaultColors(1).ForeColor = vbBlack
Else
   sspDefaultColors(1).BackColor = vbBlack
   sspDefaultColors(2).BackColor = vbDirectionColor
   sspDefaultColors(1).ForeColor = vbDirectionColor
   sspDefaultColors(2).ForeColor = vbBlack
End If

iButtonDefaultColor = Index

SetDefaultColorsVisible

Exit Sub
err1:
MsgBox "Error in Module : sspDefaultColors_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation
End Sub

Private Sub sspButtonAdjustVol_Click(Index As Integer)
On Error GoTo err1
Dim i As Integer

If sspButtonAdjustVol(Index).BackColor = vbDirectionColor Then Exit Sub

If sspButtonAdjustVol(1).BackColor = vbBlack Then
   sspButtonAdjustVol(2).BackColor = vbBlack
   sspButtonAdjustVol(1).BackColor = vbDirectionColor
   sspButtonAdjustVol(2).ForeColor = vbDirectionColor
   sspButtonAdjustVol(1).ForeColor = vbBlack
Else
   sspButtonAdjustVol(1).BackColor = vbBlack
   sspButtonAdjustVol(2).BackColor = vbDirectionColor
   sspButtonAdjustVol(1).ForeColor = vbDirectionColor
   sspButtonAdjustVol(2).ForeColor = vbBlack
End If

iAdjustVol = Index

Exit Sub
err1:
MsgBox "Error in Module : sspButtonAdjustVol_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation
End Sub

Sub SetDefaultColorsVisible()

On Error GoTo err1
Dim i As Integer

'Show the Default colors
For i = 0 To 5
 sspDefCol(i).Visible = iButtonDefaultColor = 1
Next i

Exit Sub

err1:
MsgBox "Error in Module : SetDefaultColorsVisible " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub sspDefaultSecure_Click(Index As Integer)

On Error GoTo err1

If Index = 1 Then
  Label3(6).Visible = False
  txtPassword.Visible = False
Else 'Secure Mode active
  Label3(6).Visible = True
  txtPassword.Visible = True
  txtPassword.SetFocus
End If

If sspDefaultSecure(Index).BackColor = vbDirectionColor Then Exit Sub

If sspDefaultSecure(1).BackColor = vbBlack Then
   sspDefaultSecure(2).BackColor = vbBlack
   sspDefaultSecure(1).BackColor = vbDirectionColor
   sspDefaultSecure(2).ForeColor = vbDirectionColor
   sspDefaultSecure(1).ForeColor = vbBlack
Else
   sspDefaultSecure(1).BackColor = vbBlack
   sspDefaultSecure(2).BackColor = vbDirectionColor
   sspDefaultSecure(1).ForeColor = vbDirectionColor
   sspDefaultSecure(2).ForeColor = vbBlack
End If

iSecureMode = Index  '2=Secure

Exit Sub
err1:
MsgBox "Error in Module : sspDefaultSecure_Click " & Chr(13) & Chr(13) & Err.Description, vbExclamation

End Sub

Private Sub sspDefCol_Click(Index As Integer)
  Dim i As Integer
  For i = 0 To 5
    sspDefCol(i).Font.Bold = False
    sspDefCol(i).Font.Size = 10
    sspDefCol(i).Caption = CStr(i + 1)   'Chr(i + 106)  '"Sample"
  Next i
  sspDefCol(Index).Caption = "X"   'SelectedChar
  sspDefCol(Index).Font.Bold = True
  sspDefCol(Index).Font.Size = 12
  
  'Set the variables
  vbNDefault = sspDefCol(Index).BackColor
  vbNDefaultFore = sspDefCol(Index).ForeColor
  'Save the settings
  SaveSetting regMainKey, regSubKey, "ButtonDefColor", vbNDefault
  SaveSetting regMainKey, regSubKey, "ButtonDefForColor", vbNDefaultFore
   
End Sub

Private Sub sspOption3_Click(Index As Integer)

For i = 1 To 4
   If i = Index Then
      sspOption3(i).BackColor = vbDirectionColor
      sspOption3(i).ForeColor = vbBlack
   Else
      sspOption3(i).BackColor = vbBlack
      sspOption3(i).ForeColor = vbDirectionColor
   End If
Next i

iButtonMaxSelected = Index

End Sub

