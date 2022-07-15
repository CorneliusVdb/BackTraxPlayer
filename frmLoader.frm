VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmLoader 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4875
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLoader.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel sspDemo 
      Height          =   570
      Left            =   60
      TabIndex        =   11
      Top             =   1275
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1005
      _Version        =   131074
      ForeColor       =   16776960
      MarqueeDelay    =   250
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DEMO MODE"
      BevelOuter      =   0
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   330
      Picture         =   "frmLoader.frx":000C
      ScaleHeight     =   1050
      ScaleWidth      =   3300
      TabIndex        =   8
      Top             =   375
      Width           =   3300
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H004F423C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   5910
      Picture         =   "frmLoader.frx":3551
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   7
      Top             =   3780
      Width           =   720
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   9090
      Top             =   1020
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   9075
      Top             =   525
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9075
      Top             =   60
   End
   Begin Threed.SSPanel sspError 
      Height          =   1050
      Left            =   10605
      TabIndex        =   0
      Top             =   15
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   1852
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   255
      BackColor       =   5194300
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Begin Threed.SSPanel SSPanel2 
         Height          =   405
         Left            =   60
         TabIndex        =   1
         Top             =   45
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   714
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   16777215
         BackColor       =   5194300
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Pastel Accounting is already running."
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H009C9B9A&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1890
         TabIndex        =   2
         Top             =   540
         Width           =   1275
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1500
      Left            =   0
      TabIndex        =   3
      Top             =   3450
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   2646
      _Version        =   131074
      BackColor       =   192
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Begin VB.Label lstPlugins 
         BackColor       =   &H004F423C&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1035
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   4500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sound Cards :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   255
         TabIndex        =   13
         Top             =   285
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lstPlugins 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E7DB49&
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   10
         Top             =   60
         Width           =   4500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Plug-ins :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E7DB49&
         Height          =   210
         Left            =   255
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   17
      Left            =   6165
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   16
      Left            =   5850
      Top             =   2550
      Width           =   345
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 99.99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E7DB49&
      Height          =   210
      Left            =   570
      TabIndex        =   12
      Top             =   3060
      Width           =   5820
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E7DB49&
      Height          =   210
      Left            =   555
      TabIndex        =   6
      Top             =   2190
      Width           =   6210
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E7DB49&
      Height          =   285
      Left            =   555
      TabIndex        =   5
      Top             =   1890
      Width           =   6210
   End
   Begin VB.Label lblLoading 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Line ghd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E7DB49&
      Height          =   255
      Index           =   0
      Left            =   570
      TabIndex        =   4
      Top             =   2760
      Width           =   6210
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   15
      Left            =   5505
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   14
      Left            =   5175
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   13
      Left            =   4845
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   12
      Left            =   4515
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   11
      Left            =   4185
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   10
      Left            =   3855
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   9
      Left            =   3525
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   8
      Left            =   3195
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   7
      Left            =   2865
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   6
      Left            =   2535
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   5
      Left            =   2205
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   4
      Left            =   1875
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   3
      Left            =   1545
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   2
      Left            =   1215
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H004F423C&
      Height          =   150
      Index           =   1
      Left            =   885
      Top             =   2550
      Width           =   345
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00929897&
      FillColor       =   &H00E7DB49&
      Height          =   150
      Index           =   0
      Left            =   555
      Top             =   2550
      Width           =   345
   End
End
Attribute VB_Name = "frmLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iCurrImg As Integer
Dim CntReloaded As Integer
Dim iLoadedImg As Integer

Private Sub cmdClose_Click()
   End
End Sub

Private Sub Form_Load()
  Dim FD
  Dim J As Integer
  Dim FileToOpen As String
  Dim Ctrl As Control
  
  
  lblProduct.Caption = LoadARR(MsgLits.ProdName) ' App.ProductName
  
  Me.BackColor = vbBlack   'vbWhite
  
  SSPanel1.BackColor = &H4F423C
  For i = 0 To 1
   If i = 0 Then
      lstPlugins(i).ForeColor = &HE7DB49
      lstPlugins(i).Top = 60
   Else
      lstPlugins(i).ForeColor = &HFFFF&      '&HE7DB49
      lstPlugins(i).Top = 290
   End If
   lstPlugins(i).BorderStyle = 0
   
  Next i
  
'  Picture3.Picture = LoadPicture(Win3(App.Path) & "\custom\Images\XP Images\" & "BrandingSageName.bmp")
'  Picture2.Picture = LoadPicture(Win3(App.Path) & "\custom\Images\XP Images\" & "BrandingSageIcon.bmp")



'  If App.PrevInstance Then
'    Timer1.Enabled = False
'    MsgBox Arr(18), vbOKOnly, Arr(17)
'    'MsgBox "Pastel Accounting is already running.", 16, "Pastel Accounting"
'    End
'  End If
      
  Dim strDate As String

  lblCopyright.Caption = Replace(LoadARR(MsgLits.Copyrite), "#", Format(Now(), "YYYY"))
  lblVersion.Caption = LoadARR(MsgLits.Version) & " " & App.Major & "." & App.Minor & "." & App.Revision & "  -  Created on : " & FileDateTime(Win3(App.Path) & "\" & App.EXEName & ".EXE")
  
  lblLoading(0).Caption = LoadARR(MsgLits.Start) & " " & App.Major & "." & App.Minor & "." & App.Revision & "..."
  
  For J = 0 To MaxLits
      shpProgress(J).FillStyle = 1 'Transparent
      shpProgress(J).FillColor = &HE7DB49
      shpProgress(J).BorderColor = vbBlack    '&HE7DB49
  Next J
  
  shpProgress(0).FillStyle = 0  'Solid
  shpProgress(0).BorderColor = vbBlack  '&HE7DB49
  
'  bSystemStarting = True
  
  Me.Width = 6870    '6345
  Me.Height = 4875    '5490
  
  MakeFormRound Me, 5 '  20
  
  iLoadedImg = 0
    
  Me.Enabled = False
        
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Sub ShowReloading()
Dim i As Integer
Dim iPrevImg As Integer

   iPrevImg = iCurrImg
   iCurrImg = iCurrImg + 1
   
   If iCurrImg > MaxLits Then
     iCurrImg = 0
     iPrevImg = MaxLits
   End If
   
   If iPrevImg <> 0 Then
     shpProgress(iPrevImg).FillStyle = 1 'Transparent
   End If
   shpProgress(iCurrImg).FillStyle = 0 'solid
   
   
   DoEvents

End Sub

Private Sub List1_Click()

End Sub

Private Sub SSPanel9_Click()

End Sub

Private Sub Timer1_Timer()
'   Timer1.Enabled = False
'
'   Screen.MousePointer = vbHourglass
'   DoEvents
'   Load Main
'   DoEvents
End Sub

Sub ResetImages(iDark As Integer)
Dim J As Integer
   
   For J = 0 To MaxLits
     shpProgress(J).FillStyle = 1 'Transparent
   Next J
   shpProgress(iDark).FillStyle = 0 'Solid
   DoEvents

End Sub

Sub ShowLoading(iText As Integer, iSleep As Integer)
Dim i As Integer
Dim iCurImg As Integer

   'Keep the stuff where it was
   lblLoading(0).Tag = lblLoading(0).Caption
   DoEvents
   
'   'Reset the buttons to all be
'   For I = 0 To 15
'     If shpProgress(I).FillStyle = 0 Then
'       iCurImg = I
'     End If
'     shpProgress(I).FillStyle = 1 'Transparent
'   Next I
   iLoadedImg = iLoadedImg + 1
   If iLoadedImg > MaxLits Or iText = MaxLits + 1 Then
     iLoadedImg = MaxLits
   End If
   
   iCurImg = iLoadedImg   'iCurImg + 1
'
'   If iCurImg > 15 Or iText = 16 Then
'     iCurImg = 15
'   End If
   
   shpProgress(iCurImg).FillStyle = 0 'Solid
   shpProgress(iCurImg).BorderColor = vbBlack  '&HE7DB49
   
   DoEvents
   
   If iText = 0 Then
     lblLoading(0).Caption = LoadARR(0)
     Screen.MousePointer = vbHourglass
     Me.Refresh
''     For I = 0 To 15
''       If shpProgress(I).FillStyle = 0 Then
''         iCurImg = I
''       End If
''       shpProgress(I).FillStyle = 1 'Transparent
''     Next I
     For i = 1 To MaxLits
       shpProgress(i - 1).FillStyle = 0 'Solid
       Sleep 25
       DoEvents
     Next i
'     bSystemStarting = False
   Else
     lblLoading(0).Caption = ""
     Me.Refresh
     If iText = 99 Then iText = MsgLits.Finalise
     lblLoading(0).Caption = LoadARR(iText)
     Screen.MousePointer = vbHourglass
     Me.Refresh
     Sleep 50
   End If
   
   DoEvents

End Sub

Private Sub Timer2_Timer()
   Timer2.Enabled = False
'   Timer1.Enabled = True
   Timer3.Enabled = False
End Sub

Private Sub Timer3_Timer()
   ShowReloading
End Sub
