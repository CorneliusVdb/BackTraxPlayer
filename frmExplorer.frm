VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{3CAE3B0A-2E39-407A-A22F-585643D71497}#1.0#0"; "vbalIml240_10Tec.ocx"
Begin VB.Form frmExplorer 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   11790
   ClientLeft      =   15
   ClientTop       =   15
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
   Icon            =   "frmExplorer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmExplorer.frx":000C
   ScaleHeight     =   11790
   ScaleWidth      =   20445
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   7965
      Left            =   18795
      TabIndex        =   27
      Top             =   2175
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   14049
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
      BorderWidth     =   1
      BevelOuter      =   0
      Begin Threed.SSPanel SSPanel3 
         Height          =   8940
         Left            =   2040
         TabIndex        =   28
         Top             =   60
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   15769
         _Version        =   131074
         BackColor       =   16711680
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
      End
      Begin Threed.SSCommand cmdDown 
         Height          =   945
         Left            =   135
         TabIndex        =   17
         Top             =   5625
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1667
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   0
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmExplorer.frx":3551
         AutoSize        =   1
         Alignment       =   8
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdBottom 
         Height          =   945
         Left            =   150
         TabIndex        =   18
         Top             =   6840
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1667
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   0
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmExplorer.frx":3D6F
         AutoSize        =   1
         Alignment       =   8
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdTop 
         Height          =   945
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1667
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   0
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmExplorer.frx":4AA4
         AutoSize        =   1
         Alignment       =   8
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdUp 
         Height          =   945
         Left            =   120
         TabIndex        =   16
         Top             =   1425
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1667
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   0
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmExplorer.frx":57A5
         AutoSize        =   1
         Alignment       =   8
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   17295
      MaxLength       =   50
      TabIndex        =   12
      Top             =   1245
      Width           =   2370
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   315
      Left            =   17235
      TabIndex        =   30
      Top             =   1215
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   131074
      BackColor       =   16761024
      BorderWidth     =   1
      BevelOuter      =   0
      Begin Threed.SSCommand cmdSearch 
         Height          =   345
         Left            =   2430
         TabIndex        =   13
         Top             =   -15
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         Picture         =   "frmExplorer.frx":5FC7
         AutoSize        =   1
         BevelWidth      =   0
         Outline         =   0   'False
      End
   End
   Begin vbalIml240_10Tec.vbalImageList ImageList2 
      Left            =   20145
      Top             =   120
      _ExtentX        =   1905
      _ExtentY        =   1323
      ImageWidth      =   24
      ImageHeight     =   24
      ColorDepth      =   32
      DesignTimeLastFileAddPath=   "C:\Development\BacktraxPlayer\AAA_Icon\"
      Size            =   108028
      Images          =   "frmExplorer.frx":6361
      Keys            =   $"frmExplorer.frx":2097D
   End
   Begin vbalIml240_10Tec.vbalImageList ImageList2a 
      Left            =   20175
      Top             =   870
      _ExtentX        =   1905
      _ExtentY        =   1323
      ColorDepth      =   32
      DesignTimeLastFileAddPath=   "C:\Development\BacktraxPlayer\AAA_Icon\"
      Size            =   10036
      Images          =   "frmExplorer.frx":20B2C
      Keys            =   "SB 19	SB 24	"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   405
      Index           =   0
      Left            =   -30
      TabIndex        =   29
      Top             =   1680
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
      Caption         =   "                                                                   Available Music Files"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel sspLoading 
      Height          =   975
      Left            =   8100
      TabIndex        =   25
      Top             =   5580
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   1720
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   0
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
      Caption         =   "Loading..."
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.Label lblCnt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   2250
         TabIndex        =   26
         Top             =   345
         Width           =   1695
      End
      Begin VB.Image imgPlus 
         Height          =   300
         Left            =   4725
         Picture         =   "frmExplorer.frx":23280
         Stretch         =   -1  'True
         Top             =   45
         Width           =   300
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1680
      Index           =   1
      Left            =   -45
      TabIndex        =   0
      Top             =   -45
      Width           =   17325
      _ExtentX        =   30559
      _ExtentY        =   2963
      _Version        =   131074
      BackColor       =   12632256
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   0
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   11655
         Top             =   705
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   12180
         Top             =   720
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   19635
         Top             =   1140
      End
      Begin Threed.SSPanel sspDir 
         Height          =   690
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   1095
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   1217
         _Version        =   131074
         CaptionStyle    =   1
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmExplorer.frx":484E6
         BevelOuter      =   0
         Alignment       =   4
         PictureAlignment=   7
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   1
         Left            =   1065
         TabIndex        =   2
         Top             =   1260
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C:\"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   2
         Left            =   1635
         TabIndex        =   3
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   3
         Left            =   2955
         TabIndex        =   4
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   4
         Left            =   4110
         TabIndex        =   5
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   5
         Left            =   5580
         TabIndex        =   6
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   6
         Left            =   6675
         TabIndex        =   7
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   7
         Left            =   7890
         TabIndex        =   8
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   8
         Left            =   9225
         TabIndex        =   9
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   9
         Left            =   10590
         TabIndex        =   10
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel sspDir 
         Height          =   375
         Index           =   10
         Left            =   12195
         TabIndex        =   11
         Top             =   1260
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sub Dir"
         BevelOuter      =   0
         AutoSize        =   1
         PictureAlignment=   8
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1065
         Left            =   4425
         TabIndex        =   23
         Top             =   75
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   1879
         _Version        =   131074
         BackColor       =   0
         BackStyle       =   1
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
         BevelWidth      =   2
         BevelOuter      =   0
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
         TabIndex        =   32
         Top             =   780
         Width           =   2205
      End
      Begin VB.Label lblFilename 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Test Files egljew gwf pergj "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   6780
         TabIndex        =   24
         Top             =   270
         Width           =   10500
      End
      Begin VB.Label lblDir 
         AutoSize        =   -1  'True
         Caption         =   "test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12810
         TabIndex        =   22
         Top             =   1035
         Visible         =   0   'False
         Width           =   255
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
         Left            =   12720
         MouseIcon       =   "frmExplorer.frx":48ACB
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   750
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   10
         Left            =   12030
         Picture         =   "frmExplorer.frx":49395
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   9
         Left            =   10425
         Picture         =   "frmExplorer.frx":49739
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   8
         Left            =   9060
         Picture         =   "frmExplorer.frx":49ADD
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   7
         Left            =   7725
         Picture         =   "frmExplorer.frx":49E81
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   6
         Left            =   6525
         Picture         =   "frmExplorer.frx":4A225
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   5
         Left            =   5430
         Picture         =   "frmExplorer.frx":4A5C9
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   4
         Left            =   3945
         Picture         =   "frmExplorer.frx":4A96D
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   3
         Left            =   2790
         Picture         =   "frmExplorer.frx":4AD11
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   2
         Left            =   1485
         Picture         =   "frmExplorer.frx":4B0B5
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
      Begin VB.Image imgNext 
         Height          =   345
         Index           =   1
         Left            =   855
         Picture         =   "frmExplorer.frx":4B459
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   135
      End
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   885
      Left            =   18150
      TabIndex        =   31
      Top             =   180
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1561
      _Version        =   131074
      BackColor       =   3092271
      BevelOuter      =   0
      Begin Threed.SSCommand cmdExit 
         Height          =   810
         Left            =   975
         TabIndex        =   20
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
      Begin Threed.SSCommand cmdOk 
         Height          =   810
         Left            =   45
         TabIndex        =   19
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
   End
   Begin VSFlex8LCtl.VSFlexGrid grdFiles 
      Height          =   7800
      Left            =   3945
      TabIndex        =   14
      Top             =   2220
      Width           =   14895
      _cx             =   26273
      _cy             =   13758
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   16777215
      BackColorFixed  =   0
      ForeColorFixed  =   0
      BackColorSel    =   0
      ForeColorSel    =   16777215
      BackColorBkg    =   0
      BackColorAlternate=   0
      GridColor       =   8421504
      GridColorFixed  =   0
      TreeColor       =   0
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   1
      MultiTotals     =   0   'False
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   390
      X2              =   19935
      Y1              =   10290
      Y2              =   10290
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim iExplorePos As Integer
Dim m_StartLocation As Single
Dim m_ScrollClicked As Boolean
Dim iRowSelected As Integer

Dim i As Integer
Dim sExploreCap As String
Dim sExploreTag As String
Dim iLeft As Integer
Dim iWidth As Integer
Dim iImgLeft As Integer
Dim iRow As Integer
Dim bGoingDown As Boolean
Dim bItemMoved As Boolean
Dim iPreviousSelection As Integer
Dim iBlink As Integer

Private Sub cmdBottom_Click()

Timer1.Enabled = False
Timer2.Enabled = False

grdFiles.TopRow = grdFiles.Rows - 1
grdFiles.Row = grdFiles.TopRow

grdFiles.Col = 1
grdFiles.SetFocus

   
End Sub

Private Sub cmdBottom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub cmdDown_Click()
'SetItemFocusA lvFiles, iRow + 18
End Sub

Private Sub cmdDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

bGoingDown = True
bItemMoved = False
Timer2.Enabled = True

End Sub

Private Sub cmdDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim b1 As Boolean

Timer2.Enabled = False
Timer1.Enabled = False

'grdFiles.ScrollBars = flexScrollBarVertical
If grdFiles.Row + 14 >= grdFiles.Rows Then
   grdFiles.Row = grdFiles.Rows - 1
Else
   grdFiles.Row = grdFiles.Row + 14
End If

'If Not grdFiles.RowIsVisible(grdFiles.Row) Then
   grdFiles.TopRow = grdFiles.Row
'End If
'grdFiles.ScrollTrack = True
'
'grdFiles.ScrollBars = flexScrollBarNone
grdFiles.Col = 1
grdFiles.SetFocus


End Sub

'Dim WithEvents objExt As VBControlExtender

Private Sub cmdExit_Click()
FilenameToLoad = ""
DoEvents
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

Private Sub cmdOK_Click()

If grdFiles.TextMatrix(grdFiles.Row, 2) <> "2" Then 'Valid music file
   Exit Sub
End If
If FilenameToLoad = "" Then Exit Sub
DoEvents
Unload Me

End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOk.BackColor = &HE7DB49
cmdOk.ForeColor = vbBlack
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOk.ForeColor = &HE7DB49
cmdOk.BackColor = vbBlack
End Sub

Private Sub cmdSearch_Click()
Dim iRow As Long
Dim iMaxRows As Long
Dim iSelRow As Long
'On Error Resume Next

If Trim(txtSearch.text) = "" Then Exit Sub

'iRow = grdFiles.FindRow("grl", 0, 1, False)
'iRow = grdFiles.FindRow(txtSearch.text, , 1, , False)
'iRow = grdFiles.FindRow(txtSearch.text, , 1)

iMaxRows = grdFiles.Rows - 1
iSelRow = 0
iRow = grdFiles.FindRow(txtSearch.text, 1, 1, False, False)

If iRow = -1 Then 'Exit Sub
   'Run through rows and search for text inside each row
   For iRow = 1 To iMaxRows
      If InStr(UCase(grdFiles.RowData(iRow)), UCase(Trim(txtSearch.text))) > 0 Then
         iSelRow = iRow
         Exit For
      End If
   Next iRow
Else
   iSelRow = iRow
End If

If iSelRow > 0 Then 'Valid row found
   txtSearch.text = ""
Else     'Not FOUND !!!!
   iSelRow = 1
End If

grdFiles.Row = iSelRow

grdFiles.TopRow = grdFiles.Row

grdFiles.Col = 1
grdFiles.SetFocus

End Sub

Private Sub cmdTop_Click()

SetGridTop

End Sub

Sub SetGridTop()
On Error Resume Next

'SetItemFocusA lvFiles, 1
'grdFiles.ScrollBars = flexScrollBarVertical

'If Not grdFiles.RowIsVisible(grdFiles.Row) Then
   grdFiles.TopRow = 1
   grdFiles.Row = 1
'End If
'grdFiles.ScrollBars = flexScrollBarNone
   grdFiles.Col = 1
   grdFiles.SetFocus
   
End Sub

Private Sub cmdTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub cmdUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bGoingDown = False
bItemMoved = False
'Timer2.Enabled = True

End Sub

Private Sub cmdUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer2.Enabled = False
Timer1.Enabled = False

If grdFiles.Row - 14 < 1 Then
   grdFiles.Row = 1
Else
   grdFiles.Row = grdFiles.Row - 14
End If

grdFiles.TopRow = grdFiles.Row

grdFiles.Col = 1
grdFiles.SetFocus


End Sub

Private Sub Form_Activate()

On Error Resume Next
DoEvents
DoEvents

'EnableCloseButton Me.hWnd, False

'ucLoading.LoadAnimatedGIF_File (App.Path & "\RedWait.gif") '

iExplorePos = 0
Me.Enabled = False

For i = 1 To 10
  sExploreCap = GetSetting(regMainKey, regSubKey, "Caption" & i)
  sExploreTag = GetSetting(regMainKey, regSubKey, "Tag" & i)
  iLeft = GetSetting(regMainKey, regSubKey, "Left" & i)
  iWidth = GetSetting(regMainKey, regSubKey, "Width" & i)
  iImgLeft = GetSetting(regMainKey, regSubKey, "ImgLeft" & i)
  If sExploreCap <> "" Then
    sspDir(i).Caption = sExploreCap
    sspDir(i).Tag = sExploreTag
    'Set the image's left property
    imgNext(i).Left = iImgLeft
    sspDir(i).Left = iLeft
    sspDir(i).Width = iWidth

    imgNext(i).Visible = True
    sspDir(i).Visible = True
    'Debug.Print "Index=" & i & " : BC=" & sspDir(i).Caption & " : BL=" & sspDir(i).Left & " : BW=" & sspDir(i).Width & " : IL=" & imgNext(i).Left

    iExplorePos = i
  End If
Next i

Screen.MousePointer = vbCustom
Screen.MouseIcon = lblCursorPlaceHolder.MouseIcon
'Me.Enabled = False
sspLoading.Left = vbLoadingLeft  '3855
sspLoading.Left = vbLoadingLeft  '3855

'sspLoading.ZOrder 0
Me.Refresh
DoEvents
'Timer3.Enabled = True
If iExplorePos = 0 Then
  LoadFileList grdFiles, "", False
Else
  LoadFileList grdFiles, sspDir(iExplorePos).Tag, False, 1
End If
Me.Refresh
DoEvents
'Timer3.Enabled = False
''Call cmdTop_Click  'Moves the list to top - ALWAYS !!!!!

'Timer3.Enabled = True

'SortListview
sspLoading.Left = 30000

'Screen.MousePointer = vbCustom
'Screen.MouseIcon = lblCursorPlaceHolder.MouseIcon

SetGridTop

Screen.MousePointer = Default
Me.Enabled = True
Me.Refresh
DoEvents

  
End Sub

Private Sub Form_Load()
Dim sPosition As String

'''Me.Width = 20445
'''Me.Height = 11010
'''Me.Top = 0
'''Me.Left = 0

'Me.Width = frmPlayer.Width - 120
'Me.Height = frmPlayer.Height - 120
'Me.Top = frmPlayer.Top + 45
'Me.Left = frmPlayer.Left + 45

   
'Me.Picture = LoadPicture(App.Path & "\tmpBanner")
Screen.MousePointer = vbHourglass
DoEvents

iDrawImage = 1
Timer1.Enabled = False

lblVersion.Caption = "Version   :    " & App.Major & "." & App.Minor & "." & App.Revision
lblVersion.Top = 750
   lblVersion.Left = 255
   lblVersion.Width = 2940
   lblVersion.FontSize = 7

grdFiles.FixedCols = 0
grdFiles.FixedRows = 0
grdFiles.Cols = 7
grdFiles.Rows = 1
grdFiles.ColWidth(0) = 600
grdFiles.ColWidth(1) = 11000
grdFiles.ColAlignment(1) = 1
grdFiles.ColWidth(2) = 0
grdFiles.ColWidth(3) = 0
grdFiles.ColWidth(4) = 600
grdFiles.RowHeight(0) = 600

grdFiles.ColWidth(5) = 0 'Number holder for image
grdFiles.ColWidth(6) = 0 'Number holder for image

grdFiles.GridLines = flexGridNone
grdFiles.CellAlignment = flexAlignLeftCenter


lblFilename.Caption = ""
sspLoading.Left = 30000

HelpContextID = hlpLoadSong

''''''''''''''''''''''''''''''''''SSPanel4.BackColor = vbBlack

' Hooking the form for mouse wheel scroll
'Call WheelHook(Me.hWnd)

'LoadFiles
For i = 1 To 10
  imgNext(i).Visible = False
  sspDir(i).Visible = False
  sspDir(i).Caption = ""
  sspDir(i).Tag = ""
  'sspDir(i).AutoSize = ssNoneAutoSize
  sspDir(i).Width = 100
Next i
Screen.MousePointer = vbHourglass
DoEvents

LoadDataIntoFile 119, App.Path & "\tmpPlus"   'Color
imgPlus.Picture = LoadPicture(App.Path & "\tmpPlus")

''''''''''''''SSPanel4.Width = Me.Width + 500
''''''''''''''SSPanel4.Height = SSPanel4.Height + 500


'==================================================================================
''''''Load the Grids with valid data
'''''On Error Resume Next
'''''
'''''iExplorePos = 0
'''''For i = 1 To 10
'''''  sExploreCap = GetSetting(regMainKey, regSubKey, "Caption" & i)
'''''  sExploreTag = GetSetting(regMainKey, regSubKey, "Tag" & i)
'''''  iLeft = GetSetting(regMainKey, regSubKey, "Left" & i)
'''''  iWidth = GetSetting(regMainKey, regSubKey, "Width" & i)
'''''  iImgLeft = GetSetting(regMainKey, regSubKey, "ImgLeft" & i)
'''''  If sExploreCap <> "" Then
'''''    sspDir(i).Caption = sExploreCap
'''''    sspDir(i).Tag = sExploreTag
'''''    'Set the image's left property
'''''    imgNext(i).Left = iImgLeft
'''''    sspDir(i).Left = iLeft
'''''    sspDir(i).Width = iWidth
'''''
'''''    imgNext(i).Visible = True
'''''    sspDir(i).Visible = True
'''''    'Debug.Print "Index=" & i & " : BC=" & sspDir(i).Caption & " : BL=" & sspDir(i).Left & " : BW=" & sspDir(i).Width & " : IL=" & imgNext(i).Left
'''''    iExplorePos = i
'''''  End If
'''''Next i
'''''Screen.MousePointer = vbHourglass
'''''DoEvents
''''''Screen.MousePointer = vbCustom
''''''Screen.MouseIcon = lblCursorPlaceHolder.MouseIcon
''''''Me.Enabled = False
'''''sspLoading.Left = vbLoadingLeft  '3855
''''''sspLoading.ZOrder 0
''''''Timer3.Enabled = True
'''''If iExplorePos = 0 Then
'''''  LoadFileList grdFiles, "", False
'''''Else
'''''  LoadFileList grdFiles, sspDir(iExplorePos).Tag, False, 1
'''''End If

sspLoading.Left = 30000

End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
   If ApplyStandardTheme Then
'      Me.Width = 20445  '20500  '17925
'      Me.Height = 11010 - 60  '11070
      
      Me.Width = frmPlayer.Width - 120
      Me.Height = frmPlayer.Height - 120
      Me.Top = frmPlayer.Top + 60
      Me.Left = frmPlayer.Left + 60

   Else
   '   Me.Width = 20445  '20500  '20395  '20370   '18030   '17925
   '   Me.Height = 11010 - 60  '11070
      Me.Width = frmPlayer.Width - 180
      Me.Height = frmPlayer.Height - 180
      Me.Top = frmPlayer.Top + 90
      Me.Left = frmPlayer.Left + 90
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
'Call WheelUnHook(Me.hWnd)
End Sub

Sub ItemSelected()
Dim sExt As String
Dim iImgLeft As Integer
Dim iCapLeft As Integer
Dim iCapWidth As Integer
Dim iCurRow As Integer
On Error Resume Next

grdFiles.Enabled = False

Screen.MousePointer = vbCustom
Screen.MouseIcon = lblCursorPlaceHolder.MouseIcon

lblDir.Caption = ""
lblFilename.Caption = ""
lblCnt.Visible = False

iCurRow = grdFiles.Row
If iCurRow = 0 Then
   grdFiles.Enabled = True
   grdFiles.SetFocus
   Screen.MousePointer = Default
   DoEvents
   Exit Sub
End If

'If a valid file was clicked on, show in heading and exit routine here

If grdFiles.TextMatrix(grdFiles.Row, 2) = "2" Then 'Valid music file
  sExt = LCase(Right(grdFiles.TextMatrix(grdFiles.Row, 1), 4))
  If InStr(1, Filter, sExt) > 0 Then
      FilenameToLoad = grdFiles.TextMatrix(grdFiles.Row, 3) & "\" & Trim(grdFiles.TextMatrix(grdFiles.Row, 1)) 'Path
      lblFilename.Caption = Trim(grdFiles.TextMatrix(grdFiles.Row, 1))
'''      grdFiles.Row = iPreviousSelection
'''      grdFiles.Col = 4
'''      Set grdFiles.CellPicture = ImageList2a.ItemPicture(1)
'''      grdFiles.Row = iCurRow
'''      Set grdFiles.CellPicture = ImageList2a.ItemPicture(2)
'''      iPreviousSelection = grdFiles.Row
      grdFiles.Col = 1
      ' Debug.Print FilenameToLoad
      grdFiles.Enabled = True
      grdFiles.SetFocus
      Screen.MousePointer = Default
      DoEvents
      Exit Sub
  End If
End If

'ELSE

iExplorePos = iExplorePos + 1
'Clear out all dir images from current one forawrd...
For i = iExplorePos To 10
    sspDir(i).Caption = ""
    sspDir(i).Tag = ""
    'Set the image's left property
    imgNext(i).Left = -1000
    sspDir(i).Left = -1000
    sspDir(i).Width = 0
    imgNext(i).Visible = False
    sspDir(i).Visible = False
Next i
'Sets the current dirList image true
imgNext(iExplorePos).Visible = True
sspDir(iExplorePos).Visible = True
'Sets the current DirList text
lblDir.Caption = Trim(grdFiles.TextMatrix(grdFiles.Row, 1))
sspDir(iExplorePos).Caption = lblDir.Caption
sspDir(iExplorePos).Tag = grdFiles.TextMatrix(grdFiles.Row, 3)
'Save Current settings
SaveSetting regMainKey, regSubKey, "Caption" & iExplorePos, sspDir(iExplorePos).Caption
SaveSetting regMainKey, regSubKey, "Tag" & iExplorePos, sspDir(iExplorePos).Tag

sspLoading.Left = vbLoadingLeft  '3855
DoEvents

'Load list with new selection data
If grdFiles.TextMatrix(grdFiles.Row, 2) = "0" Then
  LoadFileList grdFiles, Trim(grdFiles.TextMatrix(grdFiles.Row, 1)), 0
ElseIf grdFiles.TextMatrix(grdFiles.Row, 2) = "1" Then
  LoadFileList grdFiles, Trim(grdFiles.TextMatrix(grdFiles.Row, 3)), 0
End If
'Timer3.Enabled = False
grdFiles.Col = 1
grdFiles.Row = 1
grdFiles.SetFocus

'Set the image's left property
If iExplorePos = 1 Then
  iImgLeft = 855
Else
  iImgLeft = (sspDir(iExplorePos - 1).Left + sspDir(iExplorePos - 1).Width) + 125
End If

imgNext(iExplorePos).Left = iImgLeft
DoEvents
imgNext(iExplorePos).Left = iImgLeft

iCapLeft = iImgLeft + 260
iCapWidth = lblDir.Width  '+ 30

sspDir(iExplorePos).Left = iCapLeft
DoEvents
sspDir(iExplorePos).Left = iCapLeft


SaveSetting regMainKey, regSubKey, "Left" & iExplorePos, iCapLeft
SaveSetting regMainKey, regSubKey, "Width" & iExplorePos, iCapWidth
SaveSetting regMainKey, regSubKey, "ImgLeft" & iExplorePos, iImgLeft
'Debug.Print "Index=" & iExplorePos & " : BC=" & sspDir(iExplorePos).Caption & " : BL=" & sspDir(iExplorePos).Left & " : BW=" & sspDir(iExplorePos).Width & " : IL=" & imgNext(iExplorePos).Left

sspDir(iExplorePos).ToolTipText = "L=" & iCapLeft & " : W=" & iCapWidth
imgNext(iExplorePos).ToolTipText = "L=" & iImgLeft & " : W=135"

grdFiles.Enabled = True

'SortListview
sspLoading.Left = 30000  '3855
grdFiles.SetFocus

Screen.MousePointer = Default
DoEvents

End Sub

Private Sub grdFiles_DblClick()

If grdFiles.TextMatrix(grdFiles.Row, 2) <> "2" Then 'Valid music file
   Exit Sub
End If

If FilenameToLoad = "" Then Exit Sub
DoEvents
Unload Me

End Sub

Private Sub grdFiles_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Or KeyCode = 38 Then
   FilenameToLoad = grdFiles.TextMatrix(grdFiles.Row, 3) & "\" & Trim(grdFiles.TextMatrix(grdFiles.Row, 1)) 'Path
   lblFilename.Caption = Trim(grdFiles.TextMatrix(grdFiles.Row, 1))
   txtSearch.text = ""
   Exit Sub
End If

If KeyCode = 13 Then
   If FilenameToLoad = "" Then Exit Sub
   DoEvents
   Unload Me
End If

End Sub

Private Sub grdFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
InitY = Y
Timer1.Enabled = False
Timer1.Interval = 10
StartIndex = grdFiles.Row    '  .SelectedItem.Index

VelosityY = 0
Timer2.Enabled = True

End Sub

Private Sub grdFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Timer2.Enabled = False

EndY = Y
EndIndex = grdFiles.TopRow

txtSearch.text = ""

If Y > InitY Then
   If (Abs(Y) - Abs(InitY)) < 500 Then
      ItemSelected
      grdFiles.SetFocus
      Exit Sub
   End If
Else
   If Abs(InitY) - Abs(Y) < 500 Then
      ItemSelected
      grdFiles.SetFocus
      Exit Sub
   End If
End If


If Timer1.Enabled = False Then
   Timer1.Interval = VelosityY
   Timer1.Enabled = True
End If


End Sub
''
''Private Sub lvFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
''
''InitY = y
''Timer1.Enabled = False
''Timer1.Interval = 10
''StartIndex = lvFiles.SelectedItem.Index
''
''VelosityY = 0
''Timer2.Enabled = True
''
''End Sub
''
''Private Sub lvFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
''On Error Resume Next
''
''Timer2.Enabled = False
''
''EndY = y
''EndIndex = lvFiles.GetFirstVisible.Index
''
''If y > InitY Then
''   If (Abs(y) - Abs(InitY)) < 500 Then
''      ItemSelected
''      lvFiles.SetFocus
''      Exit Sub
''   End If
''Else
''   If Abs(InitY) - Abs(y) < 500 Then
''      ItemSelected
''      lvFiles.SetFocus
''      Exit Sub
''   End If
''End If
''
''
''If Timer1.Enabled = False Then
''   Timer1.Interval = VelosityY
''   Timer1.Enabled = True
''End If
''
''End Sub

Private Sub grdFiles_RowColChange()
   
   FilenameToLoad = grdFiles.TextMatrix(grdFiles.Row, 3) & "\" & Trim(grdFiles.TextMatrix(grdFiles.Row, 1)) 'Path
   lblFilename.Caption = Trim(grdFiles.TextMatrix(grdFiles.Row, 1))

End Sub

Private Sub sspDir_Click(Index As Integer)


On Error Resume Next

Screen.MousePointer = vbCustom
Screen.MouseIcon = lblCursorPlaceHolder.MouseIcon
Me.Enabled = False
DoEvents

If Index = 10 Then Exit Sub

If sspDir(Index + 1).Visible = False Then
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   DoEvents
   Exit Sub
End If

Me.lblFilename.Caption = ""

'Make the images of this index invisible and clear the captions...
For i = 10 To (Index + 1) Step -1
  imgNext(i).Visible = False
  sspDir(i).Visible = False
  sspDir(i).Caption = ""
  sspDir(i).Tag = ""
  SaveSetting regMainKey, regSubKey, "Caption" & i, ""  'Om die buttons se posisie te stoor
  SaveSetting regMainKey, regSubKey, "Tag" & i, ""  'Om die buttons se posisie te stoor
  SaveSetting regMainKey, regSubKey, "Left" & i, ""  'Om die buttons se posisie te stoor
  SaveSetting regMainKey, regSubKey, "Width" & i, ""  'Om die buttons se posisie te stoor
  SaveSetting regMainKey, regSubKey, "ImgLeft" & i, ""  'Om die buttons se posisie te stoor
Next i

iExplorePos = Index

sspLoading.Left = vbLoadingLeft  '3855
'DoEvents

'Timer3.Enabled = True
If Index = 0 Then
  sspDir(1).Caption = ""
  sspDir(1).Tag = ""
  imgNext(1).Visible = False
  sspDir(1).Visible = False
  LoadFileList grdFiles, "", False
Else
  LoadFileList grdFiles, sspDir(Index).Tag, False, 1
End If
'Timer3.Enabled = False
Call cmdTop_Click  'Moves the list to top - ALWAYS !!!!!

grdFiles.Col = 1
grdFiles.Row = 1
grdFiles.SetFocus

sspLoading.Left = 30000
FilenameToLoad = ""
'SortListview

Me.Enabled = True

grdFiles.SetFocus

Screen.MousePointer = vbDefault
DoEvents

End Sub

Sub SortListview()
'''lvFiles.SortKey = 1  'ColumnHeader.Index - 1
'''lvFiles.SortOrder = 0 ' lvFiles.SortOrder Xor 1
'''' Set Sorted to True to sort the list.
'''lvFiles.Sorted = True

End Sub

Private Sub sspDir_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'sspDir(Index).BackColor = &HD78C94      '&H808080
sspDir(Index).ForeColor = &H80FF&
'sspDir(Index).BevelOuter = ssInsetBevel
End Sub

Private Sub sspDir_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'MsgBox sspDir(index).Left
End Sub

Private Sub sspDir_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
sspDir(Index).ForeColor = vbWhite
End Sub

Private Sub Timer1_Timer()

If EndY < InitY Then 'Mouse moved from bottom to top, so the items needs to move up to bottom...
   If grdFiles.TopRow + 14 >= grdFiles.Rows Then
      Timer1.Enabled = False
      Timer1.Interval = 10
      Exit Sub
   End If
   '
   bGoingDown = False
Else
   If grdFiles.TopRow = 1 Then
      Timer1.Enabled = False
      Timer1.Interval = 10
      Exit Sub
   End If
   '
   bGoingDown = True
End If

Timer1.Interval = Timer1.Interval + Round((VelosityY * 1.5))  '5
If Timer1.Interval >= 200 Then
   Timer1.Enabled = False
   Timer1.Interval = 10
 '  SetItemFocusA grdFiles, grdFiles.TopRow
   Exit Sub
End If

iRow = grdFiles.TopRow

If bGoingDown Then
   If iRow = 1 Then
    '  SetItemFocusA grdFiles, 1
      grdFiles.TopRow = 1
      Timer1.Enabled = False
   Else
      grdFiles.TopRow = iRow - 1
      'SetItemFocusA grdFiles, iRow - 1
   End If
Else
   If iRow + 1 >= grdFiles.Rows Then
      grdFiles.TopRow = grdFiles.Rows - 14
     ' SetItemFocusA grdFiles, grdFiles.Rows, 1
      Timer1.Enabled = False
   Else
      'SetItemFocusA grdFiles, iRow + 1, 0
      grdFiles.TopRow = iRow + 1
   End If
End If

grdFiles.Row = grdFiles.TopRow
grdFiles.SetFocus
DoEvents

bItemMoved = True

End Sub

Private Sub Timer2_Timer()

VelosityY = VelosityY + 1

End Sub

Private Sub Timer3_Timer()

Timer3.Enabled = False

SetGridTop

'iBlink = iBlink + 1
'If iBlink > 9 Then iBlink = 0
'
'Select Case iBlink
'   Case 0
'      sspLoading.ForeColor = vbBlue
'   Case 1
'      sspLoading.ForeColor = &HF534C6  'Pers
'   Case 2
'      sspLoading.ForeColor = &H8C44E6  'Persrooi
'   Case 3
'      sspLoading.ForeColor = &H302BFF  'Rooi
'   Case 4
'      sspLoading.ForeColor = &H3181F9  'Oranje
'   Case 5
'      sspLoading.ForeColor = &H2BF4FF  'Geel
'   Case 6
'      sspLoading.ForeColor = &H2BF4FF  'Geel
'   Case 7
'      sspLoading.ForeColor = &H3181F9  'Oranje
'   Case 8
'      sspLoading.ForeColor = &H302BFF  'Rooi
'   Case 9
'      sspLoading.ForeColor = &H8C44E6  'Persrooi
'End Select

End Sub

Private Sub VSFlexGrid1_Click()

End Sub

Private Sub txtSearch_GotFocus()
txtSearch.SelStart = 0
txtSearch.SelLength = Len(txtSearch.text)
DoEvents
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cmdSearch_Click
End If
End Sub
