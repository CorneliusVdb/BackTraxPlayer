VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEq 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2220
   ClientLeft      =   15
   ClientTop       =   45
   ClientWidth     =   5145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   1
      Left            =   195
      TabIndex        =   4
      Top             =   180
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   180
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   3
      Left            =   1965
      TabIndex        =   6
      Top             =   180
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   4
      Left            =   1635
      TabIndex        =   7
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   5
      Left            =   2220
      TabIndex        =   8
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   6
      Left            =   2805
      TabIndex        =   9
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   7
      Left            =   3390
      TabIndex        =   10
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   8
      Left            =   3975
      TabIndex        =   11
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   9
      Left            =   4560
      TabIndex        =   16
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   10
      Left            =   5145
      TabIndex        =   17
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   11
      Left            =   5730
      TabIndex        =   20
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   12
      Left            =   6315
      TabIndex        =   21
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   13
      Left            =   6900
      TabIndex        =   24
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   1740
      Left            =   3240
      TabIndex        =   26
      Top             =   135
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   3069
      _Version        =   131074
      BackColor       =   3092271
      BevelOuter      =   0
      Begin Threed.SSCommand cmdExit 
         Height          =   810
         Left            =   45
         TabIndex        =   28
         Top             =   885
         Width           =   1425
         _ExtentX        =   2514
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
         Caption         =   "CLOSE"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdResetEq 
         Height          =   810
         Left            =   45
         TabIndex        =   27
         Top             =   45
         Width           =   1425
         _ExtentX        =   2514
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
         Caption         =   "Reset EQ"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
   End
   Begin MSComctlLib.Slider sldEQR 
      Height          =   1635
      Index           =   14
      Left            =   7485
      TabIndex        =   29
      Top             =   3360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   2884
      _Version        =   393216
      Orientation     =   1
      Min             =   -10
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "16k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   7485
      TabIndex        =   30
      Top             =   5025
      Width           =   555
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "12.5k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   6960
      TabIndex        =   25
      Top             =   5025
      Width           =   435
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "10k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   6450
      TabIndex        =   23
      Top             =   5025
      Width           =   285
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "6.3k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   5835
      TabIndex        =   22
      Top             =   5025
      Width           =   345
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "2.5k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4665
      TabIndex        =   19
      Top             =   5025
      Width           =   345
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "4k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   5325
      TabIndex        =   18
      Top             =   5025
      Width           =   195
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "310"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2355
      TabIndex        =   15
      Top             =   5025
      Width           =   285
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "450"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   2940
      TabIndex        =   14
      Top             =   5025
      Width           =   285
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "630"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   3525
      TabIndex        =   13
      Top             =   5025
      Width           =   285
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "1k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4155
      TabIndex        =   12
      Top             =   5025
      Width           =   195
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "250"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1635
      TabIndex        =   3
      Top             =   5025
      Width           =   555
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "8k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2145
      TabIndex        =   2
      Top             =   1845
      Width           =   195
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "1k"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1260
      TabIndex        =   1
      Top             =   1845
      Width           =   195
   End
   Begin VB.Label lblEQR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "125"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   0
      Top             =   1845
      Width           =   285
   End
End
Attribute VB_Name = "frmEq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'///////////////////////////////////////////////////////////////////////////////
' frmFXtest.frm - Copyright (c) 2001-2007 (: JOBnik! :) [Arthur Aminov, ISRAEL]
'                                                       [http://www.jobnik.org]
'                                                       [  jobnik@jobnik.org  ]
'
' BASS DX8 effects test
' Originally translated from - fxtest.c - example of Ian Luck
'///////////////////////////////////////////////////////////////////////////////

Dim chanEQ As Long
Dim mSongIndex As Integer
Dim bLoading As Boolean
Dim SetEq As Boolean



Public Property Let SongIndex(value As Integer)
  mSongIndex = value
End Property

'Dim chan As Long         ' channel handle
'Dim fx(3) As Long        ' 3 eq band + reverb

' display error messages
Public Sub Error_(ByVal es As String)
    Call MsgBox(es & vbCrLf & vbCrLf & "error code: " & BASS_ErrorGetCode, vbExclamation, "Error")
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdResetEq_Click()

  For i = 1 To MaxEqs
    sldEQR(i).value = 0
    SetButtonEqArrayValue mSongIndex, i, 0
    Call UpdateSliderFX(i) ' 32
  Next i
        
End Sub

Private Sub Form_Initialize()
'    ' change and set the current path, to prevent from VB not finding BASS.DLL
'    ChDrive "C:\DevOther\BASS" 'App.Path
'    ChDir "C:\DevOther\BASS" 'App.Path
'
'    ' check the correct BASS was loaded
'    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
'        Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
'        End
'    End If
'
'    ' setup output - default device
'    If (BASS_Init(-1, 44100, 0, Me.hWnd, 0) = 0) Then
'        Call Error_("Can't initialize device")
'        End
'    End If

    ' check that DX8 features are available
    Dim bi As BASS_INFO
    Call BASS_GetInfo(bi)
    If (bi.dsver < 8) Then
        Call BASS_Free
        Call Error_("DirectX 8 is not installed")
        End
    End If
End Sub

Private Sub Form_Load()

If iPlayingChan = 0 Then Unload Me: Exit Sub

  bLoading = True
  
  chanEQ = iPlayingChan
  
  ' you can add more EQ bands with changing:
  ' p.fCenter = N [Hz] N>=80 and N<=16000
  
  
   'Setup the Effects Chan info so we can do EQ
   
    Dim i As Integer
    Dim eq As BASS_BFX_PEAKEQ
    
    ' check the correct BASS_FX was loaded
    If (HiWord(BASS_FX_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of BASS_FX.DLL was loaded (2.4 is required)", vbCritical)
    End If
       
       
'    Dim eq As BASS_BFX_PEAKEQ

    ' set peaking equalizer effect with no bands
    fxEQ = BASS_ChannelSetFX(chanEQ, BASS_FX_BFX_PEAKEQ, 0)

    eq.fBandwidth = 2.5
    eq.fQ = 0#
    eq.fGain = 0#
    eq.lChannel = BASS_BFX_CHANALL

    ' create 1st band for bass
    eq.lBand = 1
    eq.fCenter = 125
    Call BASS_FXSetParameters(fxEQ, eq)
    
    ' create 2nd band for mid
    eq.lBand = 2
    eq.fCenter = 1000
    Call BASS_FXSetParameters(fxEQ, eq)
    
    ' create 3rd band for treble
    eq.lBand = 3
    eq.fCenter = 8000
    Call BASS_FXSetParameters(fxEQ, eq)

'    ' update dsp eq
'    Call sldEQR_Change(1)
'    Call sldEQR_Change(2)
'    Call sldEQR_Change(3)
        
 '  SetupEqFx iPlayingChan


    
  GetEqArrayValues 'Set the sliders to correct positions
  
  
  
'   'If any EQ in the array is NOT = 10 then Do this
'   For iEqBands = 1 To MaxEqs
'     ' If aEQ(mSongIndex, iEqBands) <> 10 Then
'        UpdateSongFreqEq iEqBands, aEQ(mSongIndex, iEqBands)
'     ' End If
'   Next iEqBands

  bLoading = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call BASS_Free
End Sub

Private Sub SetupInitialEQ()

End Sub

Public Sub UpdateSliderFX(ByVal b As Integer)
    Dim v As Integer
    Dim eq As BASS_BFX_PEAKEQ
    
    v = sldEQR(b).value

    eq.lBand = b    ' Band values you would like to get

    Call BASS_FXGetParameters(fxEQ, eq)
    eq.fGain = sldEQR(b).value * -1
    Call BASS_FXSetParameters(fxEQ, eq)
    
    SetButtonEqArrayValue mSongIndex, b, v  'sldEQR(Index).value
    
    
    
End Sub

Private Sub GetEqArrayValues()
  
  For iEqBands = 1 To MaxEqs
    sldEQR(iEqBands).value = aEQ(mSongIndex, iEqBands)    'Slider value at default (middle is 0)
    If aEQ(mSongIndex, iEqBands) <> 0 Then SetEq = True
  Next iEqBands

End Sub


Private Sub sldEQR_Change(Index As Integer)
    
    If bLoading Then Exit Sub 'Exit if form load is happening
    
    Dim eq As BASS_BFX_PEAKEQ

    eq.lBand = Index    ' Band values you would like to get

    Call BASS_FXGetParameters(fxEQ, eq)
    eq.fGain = sldEQR(Index).value * -1
    Call BASS_FXSetParameters(fxEQ, eq)
    
    SetButtonEqArrayValue mSongIndex, Index, sldEQR(Index).value 'Set the array with the values of the sliders, per song

    
End Sub

Private Sub sldEQR_Scroll(Index As Integer)
  
  If bLoading Then Exit Sub
   
  Call sldEQR_Change(Index)
  sldEQR(Index).text = sldEQR(Index).value * -1
     
End Sub

'--------------------
' useful function :)
'--------------------

' get file name from file path
Public Function GetFileName(ByVal fp As String) As String
    GetFileName = Mid(fp, InStrRev(fp, "\") + 1)
End Function
