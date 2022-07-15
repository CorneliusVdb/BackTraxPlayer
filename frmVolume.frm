VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{7053654B-A6C9-4C60-B4AA-CB8D1BCFC2C0}#1.0#0"; "cpvslider.ocx"
Begin VB.Form frmVolume 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Volume and EQ"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVolume.frx":0000
   ScaleHeight     =   4710
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Slider2.cpvSlider cpvVol 
      Height          =   1800
      Left            =   4785
      Top             =   1995
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   3175
      BackColor       =   0
      SliderIcon      =   "frmVolume.frx":220D6
      RailPicture     =   "frmVolume.frx":222B0
      RailStyle       =   1
      Max             =   100
      Value           =   1
   End
   Begin Slider2.cpvSlider cpvBass 
      Height          =   1800
      Left            =   1665
      Top             =   1995
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   3175
      BackColor       =   0
      SliderIcon      =   "frmVolume.frx":222CC
      RailPicture     =   "frmVolume.frx":224A6
      RailStyle       =   1
   End
   Begin Slider2.cpvSlider cpvHigh 
      Height          =   1800
      Left            =   3345
      Top             =   1995
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   3175
      BackColor       =   0
      SliderIcon      =   "frmVolume.frx":224C2
      RailPicture     =   "frmVolume.frx":2269C
      RailStyle       =   1
      Value           =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   405
      Index           =   2
      Left            =   -105
      TabIndex        =   0
      Top             =   1305
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
      Caption         =   "                    Volume and Equalization"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Slider2.cpvSlider cpvMid 
      Height          =   1800
      Left            =   2475
      Top             =   1995
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   3175
      BackColor       =   0
      SliderIcon      =   "frmVolume.frx":226B8
      RailPicture     =   "frmVolume.frx":22892
      RailStyle       =   1
      Value           =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H002F2F2F&
      Index           =   4
      X1              =   1425
      X2              =   5265
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H002F2F2F&
      Index           =   3
      X1              =   1425
      X2              =   5265
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "-12 db"
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   2
      Left            =   3810
      TabIndex        =   8
      Top             =   3525
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "+12 db"
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   1
      Left            =   3810
      TabIndex        =   7
      Top             =   2025
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H005A5A5F&
      Index           =   2
      X1              =   1425
      X2              =   5265
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H005A5A5F&
      Index           =   1
      X1              =   1425
      X2              =   5265
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0 db"
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   0
      Left            =   3810
      TabIndex        =   6
      Top             =   2805
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H005A5A5F&
      Index           =   0
      X1              =   1425
      X2              =   5265
      Y1              =   2895
      Y2              =   2895
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   810
      Left            =   5460
      TabIndex        =   5
      Top             =   195
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1429
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmVolume.frx":228AE
      Caption         =   "OK"
      AutoSize        =   1
      Alignment       =   8
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
   Begin VB.Label lblVol 
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
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   0
      Left            =   4590
      TabIndex        =   4
      Top             =   3825
      Width           =   645
   End
   Begin VB.Label lblHigh 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Treble"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   0
      Left            =   3165
      TabIndex        =   3
      Top             =   3825
      Width           =   645
   End
   Begin VB.Label lblMid 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Index           =   0
      Left            =   2265
      TabIndex        =   2
      Top             =   3825
      Width           =   645
   End
   Begin VB.Label lblBass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   195
      Index           =   0
      Left            =   1455
      TabIndex        =   1
      Top             =   3825
      Width           =   645
   End
End
Attribute VB_Name = "frmVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
lVolume = cpvVolume.value

lBass = cpvBass.value
lMid = cpvMid.value
lHigh = cpvHigh.value


Unload Me
End Sub


Private Sub cpvVolume_Click()

End Sub

Private Sub cpvVolume_ValueChanged()
lVolume = cpvVolume.value

End Sub

Private Sub Form_Load()

'Set the values according to song values
cpvVolume.value = lVolume

cpvBass.value = lBass
cpvMid.value = lMid
cpvHigh.value = lHigh

End Sub
