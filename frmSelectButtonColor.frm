VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmSelectButtonColor 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7560
   ClientLeft      =   2760
   ClientTop       =   3450
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSelectButtonColor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   4920
      Picture         =   "frmSelectButtonColor.frx":044A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   3150
      Width           =   540
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   495
      Left            =   7995
      TabIndex        =   0
      Top             =   1620
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   131074
      Caption         =   "SSPanel3"
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   555
      Index           =   0
      Left            =   45
      TabIndex        =   5
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   979
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSelectButtonColor.frx":0E06
      Caption         =   "                 BACK"
      BevelOuter      =   0
      Alignment       =   1
      PictureAlignment=   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   405
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Options"
      BorderWidth     =   1
      BevelOuter      =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   270
      Index           =   1
      Left            =   45
      TabIndex        =   3
      Top             =   3180
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   476
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   7104768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "  Button Colors"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   270
      Index           =   2
      Left            =   45
      TabIndex        =   4
      Top             =   1125
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   476
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   7104768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "  Button Options"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   555
      Index           =   1
      Left            =   45
      TabIndex        =   6
      Top             =   1410
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   979
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSelectButtonColor.frx":16E0
      Caption         =   "                 Clear Button"
      BevelOuter      =   0
      Alignment       =   1
      PictureAlignment=   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   555
      Index           =   2
      Left            =   45
      TabIndex        =   7
      Top             =   1980
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   979
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSelectButtonColor.frx":1DDA
      Caption         =   "                 Load new Track"
      BevelOuter      =   0
      Alignment       =   1
      PictureAlignment=   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   555
      Index           =   3
      Left            =   45
      TabIndex        =   8
      Top             =   2565
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   979
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSelectButtonColor.frx":24D4
      Caption         =   "                 Edit ID3 Tags"
      BevelOuter      =   0
      Alignment       =   1
      PictureAlignment=   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   4
      Left            =   45
      TabIndex        =   9
      Top             =   3480
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   6
      Left            =   45
      TabIndex        =   10
      Top             =   4500
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   16727114
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   7
      Left            =   45
      TabIndex        =   11
      Top             =   5010
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   5658640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   8
      Left            =   45
      TabIndex        =   12
      Top             =   5520
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   32013
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   9
      Left            =   45
      TabIndex        =   13
      Top             =   6030
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   32896
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   12
      Left            =   1395
      TabIndex        =   15
      Top             =   3480
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   4195458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   13
      Left            =   1395
      TabIndex        =   16
      Top             =   3990
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   16455060
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   14
      Left            =   1395
      TabIndex        =   17
      Top             =   4500
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   16293644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   15
      Left            =   1395
      TabIndex        =   18
      Top             =   5010
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   16
      Left            =   1395
      TabIndex        =   19
      Top             =   5520
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   20776
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   17
      Left            =   1395
      TabIndex        =   20
      Top             =   6030
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   4391409
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   5
      Left            =   45
      TabIndex        =   21
      Top             =   3990
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   10
      Left            =   45
      TabIndex        =   22
      Top             =   6540
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   18
      Left            =   1395
      TabIndex        =   23
      Top             =   6540
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   31980
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   11
      Left            =   45
      TabIndex        =   24
      Top             =   7050
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   0
      BackColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   480
      Index           =   19
      Left            =   1395
      TabIndex        =   25
      Top             =   7050
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      _Version        =   131074
      ForeColor       =   68478
      BackColor       =   136159
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      TEST"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   555
      Index           =   20
      Left            =   3975
      TabIndex        =   26
      Top             =   3150
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   979
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSelectButtonColor.frx":2E37
      Caption         =   "                 Ducking"
      BevelOuter      =   0
      Alignment       =   1
      PictureAlignment=   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   3
      X1              =   45
      X2              =   2700
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   3915
      X2              =   6555
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   45
      X2              =   2700
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   45
      X2              =   2700
      Y1              =   1965
      Y2              =   1965
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   855
      Left            =   12345
      TabIndex        =   2
      Top             =   1095
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1508
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
      Picture         =   "frmSelectButtonColor.frx":379A
      Caption         =   "OK"
      AutoSize        =   1
      Alignment       =   8
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
End
Attribute VB_Name = "frmSelectButtonColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iI As Integer
Dim bLoading As Boolean

Private Sub cmdClear_Click()
bClearButton = True
Unload Me
End Sub

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdColor_Click(Index As Integer)

   'ClearButtonProperties
   
   SetButtonColor Index
   
   Unload Me
   
End Sub

Private Sub cmdExit_Click()

   bExitSetup = True
   Unload Me
   
End Sub

Private Sub cmdLoad_Click()

bLoadButton = True
Unload Me

End Sub

Private Sub cmdOK_Click()

   Unload Me

End Sub

Private Sub cmdTag_Click()

bTagEdit = True
Unload Me

End Sub

Private Sub Form_Activate()

  ' EnableCloseButton Me.hWnd, False
   'SSPanel3.SetFocus
   
  ' ctrfrm Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
   bExitSetup = True
   Unload Me
End If

End Sub

Private Sub Form_Load()

''Public Const vbNDefault As Long = &H404040    'Default (0)
''Public Const vbNDGrey As Long = &HC8B5BA
''Public Const vbNSeaGreen As Long = &H7DA71D
''Public Const vbNDOrange As Long = &H58FF&
''Public Const vbNLOrange As Long = &H50AFFE
''Public Const vbNLPink As Long = &HFF80FF
''Public Const vbNGreen As Long = &H33B234
''Public Const vbNLRed As Long = &H4F61FF
''Public Const vbNDRose As Long = &H6D2BF9
''Public Const vbNGreenYellow As Long = &HDCAAD
''Public Const vbNViolet As Long = &HEF748C
''Public Const vbNLBlue As Long = &HC5975F


  
   bLoading = True
      
   bClearButton = False
   bLoadButton = False
   bExitSetup = False
   bTagEdit = False
   bDucking = False
   
   HelpContextID = hlpButtons
   
   Me.BackColor = vbBlack
   
   ClearButtonProperties
   
  ''' optSelected(gColor).value = True
   SetButtonColor gColor
   
   'cmdClear.Picture = LoadPicture(App.Path & "\tmpClear")
   SSPanel2(1).Picture = LoadPicture(App.Path & "\tmpClear")
   
   'cmdLoad.Picture = LoadPicture(App.Path & "\tmpOpen")
   SSPanel2(2).Picture = LoadPicture(App.Path & "\tmpOpen")
   
'   If bTagEditMP3 = True Then
'      SSPanel2(3).Enabled = True
'      SSPanel2(3).ForeColor = vbWhite
'   Else
'      SSPanel2(3).ForeColor = vbBlack
'      SSPanel2(3).Enabled = False
'      SSPanel2(3).Picture = Picture1.Picture
'   End If
   
    SSPanel2(3).Enabled = True
    SSPanel2(3).ForeColor = vbWhite
    If bTagEditMP3 Then
      SSPanel2(3).Caption = "                 Edit ID3 Tags"
    Else
      SSPanel2(3).Caption = "                 Show File Info"
    End If
      
   
   Me.Top = ButTop
   Me.Left = ButLeft
   Me.Width = 2835
   Me.Height = 7650   '7875  '7560    7305
   
'   For i = 4 To 19
'      Debug.Print "Public Const vbNColor" & i - 3 & " As Long = &H" & Hex(SSPanel2(i).BackColor)
'   Next i
   
   bLoading = False
   
End Sub

Private Sub Image1_Click(Index As Integer)
   
'   ClearButtonProperties
   
  ' optSelected(Index).value = True
   
   SetButtonColor Index
   
   Unload Me
   
End Sub

Sub SetButtonColor(iIndex As Integer)

 '  ClearButtonProperties
   
   gColor = iIndex

'''   optSelected(iIndex).ForeColor = vbYellow
'''   optSelected(iIndex).Caption = "Selected                            ."
'''   optSelected(iIndex).Font.Bold = True
'''   optSelected(iIndex).Font.Italic = False
'''   optSelected(iIndex).Font.Size = 11
'''
'''   optSelected(iIndex).MarqueeStyle = ssBlinkingMarquee
'''   optSelected(iIndex).MarqueeDelay = 300
'''   optSelected(iIndex).value = ssCBChecked
   
   
End Sub

Sub ClearButtonProperties()

''''   For iI = 0 To 5
''''      optSelected(iI).MarqueeStyle = ssNoneMarquee
''''      optSelected(iI).MarqueeDelay = 0
''''      optSelected(iI).ForeColor = vbWhite
''''      optSelected(iI).Caption = "Click to select                        ."
''''      optSelected(iI).Font.Bold = False
''''      optSelected(iI).Font.Italic = True
''''      optSelected(iI).Font.Size = 10
''''      optSelected(iI).value = ssCBUnchecked
''''   Next iI

For iI = 4 To 19
   SSPanel2(iI).Caption = ""
Next iI
   
End Sub

Private Sub optSelected_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Not bLoading Then
      SetButtonColor Index - 1
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault

End Sub

Private Sub SSPanel2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

SSPanel2(Index).BackColor = &H9D5F00

End Sub



Private Sub SSPanel2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Secret code to generate serial numbers
'Right-click while ctrl + shift is pressed
If Button = 2 And Shift = 3 Then
   Dim sSerial As String
   sSerial = GenerateNewSerial
   MsgBox "Serial number : " & sSerial, vbInformation, "Serial Number"
   Debug.Print "Serial number : " & sSerial
   Exit Sub
End If



SSPanel2(Index).BackColor = vbBlack
DoEvents

Select Case Index
   Case 0
      bExitSetup = True
   Case 1
      bClearButton = True
   Case 2
      bLoadButton = True
   Case 3
      bTagEdit = True
   Case 4 To 19
      SetButtonColor (Index - 3)
   Case 20   'Ducking
      bDucking = True
End Select

Unload Me

End Sub
