VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmDucking 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ducking"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDucking.frx":0000
   ScaleHeight     =   3180
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6045
      Top             =   225
   End
   Begin VB.PictureBox sspProgress 
      BackColor       =   &H00000000&
      Height          =   825
      Left            =   195
      ScaleHeight     =   765
      ScaleWidth      =   8040
      TabIndex        =   3
      Top             =   1770
      Width           =   8100
   End
   Begin VB.Timer tmrCustLoop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6585
      Top             =   270
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   405
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   930
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
      Caption         =   "                    Ducking positioning"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin VB.Label lblPos 
      Alignment       =   2  'Center
      Caption         =   "0:00"
      Height          =   270
      Left            =   1200
      TabIndex        =   4
      Top             =   2745
      Width           =   1455
   End
   Begin VB.Label lblSongName 
      BackStyle       =   0  'Transparent
      Caption         =   "Song Name.mp3"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   300
      TabIndex        =   2
      Top             =   1440
      Width           =   7890
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   810
      Left            =   7275
      TabIndex        =   0
      Top             =   75
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
      Picture         =   "frmDucking.frx":4EDD
      Caption         =   "OK"
      AutoSize        =   1
      Alignment       =   8
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
End
Attribute VB_Name = "frmDucking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

tmrCustLoop.Enabled = False

Unload Me

End Sub

Private Sub Form_Load()


   Me.lblSongName.Caption = FilenameToLoad
   
   
   Timer1.Enabled = True

   
   'Me.Show
   
End Sub

Private Sub sspProgress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    tmrCustLoop.Enabled = False
    
    If (Button = vbLeftButton) Then ' set loop start
        Call SetLoopStart(X * bpp)
        Call DrawTimeLine(sspProgress.hdc, loop_(0), &HFFFF00, 12)    ' loop start
    ElseIf (Button = vbRightButton) Then    ' set loop end
        Call SetLoopEnd(X * bpp)
        Call DrawTimeLine(sspProgress.hdc, loop_(1), vbYellow, 24) ' loop end
    End If
    
End Sub

Private Sub Timer1_Timer()
   
   Timer1.Enabled = False

   tmrCustLoop.Enabled = True
   PlayFile FilenameToLoad
   
End Sub

Private Sub tmrCustLoop_Timer()

    With sspProgress
        ' draw buffered peak waveform
        Call SetDIBitsToDevice(.hdc, 0, 0, WIDTH_, HEIGHT_, 0, 0, 0, HEIGHT_, wavebuf(-(WIDTH_ / 2)), bh, 0)
        Call DrawTimeLine(.hdc, BASS_ChannelGetPosition(chanFreq, BASS_POS_BYTE), &HFFFFFF, 0)  ' current pos
        Call DrawTimeLine(.hdc, loop_(0), &HFFFF00, 12) ' loop start
        Call DrawTimeLine(.hdc, loop_(1), vbYellow, 24) ' loop end
    End With
    
    
End Sub

