VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmTestMp3Tags 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6495
   ClientLeft      =   4140
   ClientTop       =   1605
   ClientWidth     =   8910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTestMp3Tags.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "frmTestMp3Tags.frx":1272
   ScaleHeight     =   6495
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   3300
      Index           =   1
      Left            =   915
      TabIndex        =   37
      Top             =   3195
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   5821
      _Version        =   131074
      BackColor       =   0
      BevelOuter      =   0
      Begin VB.TextBox txtTrackNo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1245
         MaxLength       =   4
         TabIndex        =   0
         Top             =   585
         Width           =   495
      End
      Begin VB.ComboBox cboGenre 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1245
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2025
         Width           =   1815
      End
      Begin VB.TextBox txtComment 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1245
         MaxLength       =   28
         TabIndex        =   6
         Top             =   2745
         Width           =   5835
      End
      Begin VB.TextBox txtYear 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1245
         MaxLength       =   4
         TabIndex        =   5
         Top             =   2385
         Width           =   1815
      End
      Begin VB.TextBox txtAlbum 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1245
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1665
         Width           =   5835
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1245
         MaxLength       =   30
         TabIndex        =   1
         Top             =   945
         Width           =   5835
      End
      Begin VB.TextBox txtArtist 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1245
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1305
         Width           =   5835
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   45
         Top             =   135
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   556
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   7104768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "       ID3 v1 Tag Information"
         BorderWidth     =   1
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Track # :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   18
         Left            =   435
         TabIndex        =   44
         Top             =   645
         Width           =   795
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Comment :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   435
         TabIndex        =   43
         Top             =   2805
         Width           =   795
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   435
         TabIndex        =   42
         Top             =   2445
         Width           =   795
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Genre :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   435
         TabIndex        =   41
         Top             =   2085
         Width           =   795
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Album :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   435
         TabIndex        =   40
         Top             =   1725
         Width           =   795
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Title :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   435
         TabIndex        =   39
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Artist :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   435
         TabIndex        =   38
         Top             =   1365
         Width           =   795
      End
   End
   Begin VB.PictureBox picID3v2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5715
      Left            =   960
      ScaleHeight     =   5715
      ScaleWidth      =   6735
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8070
      Width           =   6735
      Begin VB.ComboBox cboGenreV2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   1380
         Width           =   1815
      End
      Begin VB.TextBox txtCommentV2 
         Enabled         =   0   'False
         Height          =   1335
         Left            =   900
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2100
         Width           =   5835
      End
      Begin VB.TextBox txtYearV2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   4
         TabIndex        =   21
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox txtAlbumV2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   17
         Top             =   1020
         Width           =   5835
      End
      Begin VB.TextBox txtTitleV2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   15
         Top             =   660
         Width           =   5835
      End
      Begin VB.TextBox txtArtistV2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   13
         Top             =   300
         Width           =   5835
      End
      Begin VB.TextBox txtTrack 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         MaxLength       =   4
         TabIndex        =   25
         Top             =   3480
         Width           =   1875
      End
      Begin VB.TextBox txtOriginalArtist 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   27
         Top             =   3840
         Width           =   5835
      End
      Begin VB.TextBox txtEncodedBy 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   33
         Top             =   4920
         Width           =   5835
      End
      Begin VB.TextBox txtLinkTo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   35
         Top             =   5280
         Width           =   5835
      End
      Begin VB.TextBox txtComposer 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   29
         Top             =   4200
         Width           =   5835
      End
      Begin VB.TextBox txtCopyright 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   31
         Top             =   4560
         Width           =   5835
      End
      Begin VB.Label lblInfo 
         Caption         =   "Comment:"
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   22
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Year:"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   20
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Genre:"
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   18
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Album:"
         Height          =   255
         Index           =   9
         Left            =   60
         TabIndex        =   16
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Title:"
         Height          =   255
         Index           =   10
         Left            =   60
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Caption         =   "Artist:"
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   12
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Track:"
         Height          =   255
         Index           =   12
         Left            =   60
         TabIndex        =   24
         Top             =   3540
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Org. Art.:"
         Height          =   255
         Index           =   13
         Left            =   60
         TabIndex        =   26
         Top             =   3900
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Encoded:"
         Height          =   255
         Index           =   14
         Left            =   60
         TabIndex        =   32
         Top             =   4980
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Link To:"
         Height          =   255
         Index           =   15
         Left            =   60
         TabIndex        =   34
         Top             =   5340
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Composer:"
         Height          =   195
         Index           =   16
         Left            =   60
         TabIndex        =   28
         Top             =   4260
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Copyright: (Yeah, Right)"
         Height          =   195
         Index           =   17
         Left            =   60
         TabIndex        =   30
         Top             =   4620
         Width           =   795
      End
      Begin VB.Label lblID3V2 
         BackColor       =   &H80000010&
         Caption         =   " ID3 v2 Tag Information"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   7515
      End
   End
   Begin VB.PictureBox picID3v1 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   10635
      ScaleHeight     =   2595
      ScaleWidth      =   6735
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3495
      Width           =   6735
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10665
      TabIndex        =   36
      Top             =   795
      Width           =   1155
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "..."
      Height          =   315
      Left            =   11235
      TabIndex        =   8
      Top             =   255
      Width           =   375
   End
   Begin VB.TextBox txtMp3File 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1260
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Text            =   "frmTestMp3Tags.frx":47B7
      Top             =   2490
      Width           =   7140
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   405
      Index           =   2
      Left            =   0
      TabIndex        =   46
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
      Caption         =   "                    ID3 v1 Tag Information"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Index           =   3
      Left            =   915
      TabIndex        =   47
      Top             =   2160
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   556
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   7104768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "     Music File Name"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   6105
      Left            =   -45
      TabIndex        =   48
      Top             =   1755
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   10769
      _Version        =   131074
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   885
      Left            =   6525
      TabIndex        =   50
      Top             =   165
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1561
      _Version        =   131074
      BackColor       =   3092271
      BevelOuter      =   0
      Begin Threed.SSCommand cmdLoadPalette 
         Height          =   810
         Left            =   60
         TabIndex        =   52
         Top             =   45
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1429
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   15194953
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Save Tag Info"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   810
         Left            =   960
         TabIndex        =   51
         Top             =   45
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1429
         _Version        =   131074
         ForeColor       =   15194953
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
      TabIndex        =   49
      Top             =   810
      Width           =   2205
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmTestMp3Tags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cID3v1 As New cMP3ID3v1
Private m_cID3v2 As New cMP3ID3v2

Const ScreenWidth As Integer = 7995
Const ScreenHeight As Integer = 4515
Dim Dirty As Boolean
Dim bLoading As Boolean


Private Sub ParseCommand(ByVal sCmd As String)
Dim sC As String
Dim bInQuote As Boolean
Dim bOk As Boolean
Dim sParsed As String
Dim i As Long

   For i = 1 To Len(sCmd)
      sC = Mid(sCmd, i, 1)
      If (sC = """") Then
         If (bInQuote) Then
            bInQuote = False
            bOk = True
            Exit For
         Else
            bInQuote = True
         End If
      ElseIf (sC = " ") Then
         If (bInQuote) Then
            sParsed = sParsed & sC
         Else
            bOk = True
            Exit For
         End If
      Else
         sParsed = sParsed & sC
      End If
   Next i
   
   If (bOk) Or Not (bInQuote) Then
      LoadFile sParsed
   End If
   
End Sub

Private Sub EnableControls(ByVal State As Boolean)
Dim ctl As Control
Dim txt As TextBox
Dim cbo As ComboBox
Dim lst As ComboBox
Dim oCol As OLE_COLOR
   oCol = IIf(State, vbWindowBackground, vbButtonFace)
   For Each ctl In Controls
      If (ctl Is cmdPick) Then
      ElseIf (ctl Is txtMp3File) Then
      ElseIf (TypeOf ctl Is Menu) Then
         If (ctl Is mnuFile(1)) Then
            ctl.Enabled = State
         End If
      ElseIf (TypeOf ctl Is Frame) Then
      ElseIf (TypeOf ctl Is Label) Then
      ElseIf (TypeOf ctl Is TextBox) Then
         Set txt = ctl
         txt.BackColor = oCol
         txt.Enabled = State
      ElseIf (TypeOf ctl Is ComboBox) Then
         Set cbo = ctl
         cbo.BackColor = oCol
         cbo.Enabled = State
      ElseIf (TypeOf ctl Is ListBox) Then
         Set lst = ctl
         lst.BackColor = oCol
         lst.Enabled = False
      Else
         ctl.Enabled = State
      End If
   Next
End Sub

Private Sub showMp3Tags()
   
   ' Get ID3v1 Tag Information:
   With m_cID3v1
      .MP3File = txtMp3File.text
   
      txtArtist.text = .Artist
      txtTitle.text = .Title
      txtAlbum.text = .Album
      If (.GenreName(m_cID3v1.Genre) = "") Then
         ' not sure why VB doesn't allow .Text to be set
         cboGenre.ListIndex = 0
      Else
         cboGenre.text = .GenreName(m_cID3v1.Genre)
      End If
      txtYear.text = .Year
      txtComment.text = .Comment
      txtTrackNo.text = .Track
      
   End With
   
'   With m_cID3v2
'      .MP3File = txtMp3File.text
'
'      txtArtistV2.text = .Artist
'      txtTitleV2.text = .title
'      txtAlbumV2.text = .Album
'      cboGenreV2.text = .GenreName(m_cID3v1.Genre)
'      txtYearV2.text = .Year
'      txtCommentV2.text = .Comment
'      txtTrack.text = .Track
'      txtOriginalArtist.text = .OriginalArtist
'      txtComposer.text = .Composer
'      txtCopyright.text = .Copyright
'      txtEncodedBy.text = .EncodedBy
'      txtLinkTo.text = .LinkTo
'   End With
   
End Sub

Private Sub cboGenre_Change()
If bLoading Then Exit Sub
Dirty = True
End Sub

Private Sub cboGenre_Click()
If bLoading Then Exit Sub
Dirty = True
End Sub

Private Sub cboGenre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdExit_Click()
bTagsUpdated = False
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

Private Sub cmdLoadPalette_Click()

If Dirty Then
   SetTags
   bTagsUpdated = True
End If
Unload Me

End Sub

Private Sub cmdLoadPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdLoadPalette.BackColor = &HE7DB49
cmdLoadPalette.ForeColor = vbBlack
End Sub

Private Sub cmdLoadPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdLoadPalette.ForeColor = &HE7DB49
cmdLoadPalette.BackColor = vbBlack
End Sub

Private Sub cmdPick_Click()
   PickFile
End Sub

Public Sub LoadFile(ByVal sFile As String)
Dim bOk As Boolean

On Error GoTo errorHandler
   txtMp3File.text = sFile
   showMp3Tags
   bOk = True
   
errorHandler:
   EnableControls bOk
   cmdExit.Enabled = True
   If Not (Err.Number = 0) Then
      MsgBox "Could not load file '" & sFile & "'" & vbCrLf & vbCrLf & Err.Description, vbInformation
   End If
End Sub

Private Sub PickFile()
   Dim sFile As String
   If (VBGetOpenFileName( _
         FileName:=sFile, _
         Filter:="MP3 Files (*.MP3)|*.MP3|All Files (*.*)|*.*", _
         DefaultExt:="MP3", _
         Owner:=Me.hWnd)) Then
      LoadFile sFile
   End If
End Sub

Private Sub cmdSet_Click()
   SetTags
End Sub

Private Sub SetTags()

'On Error Resume Next
On Error GoTo err1
   
   Dim bFound As Boolean
   Dim i As Long
   
   ' set genre v1:
   If Len(cboGenre.text) > 0 Then
      For i = 0 To cboGenre.ListCount - 1
         If InStr(cboGenre.List(i), cboGenre.text) > 0 Then
            cboGenre.ListIndex = i
            bFound = True
            Exit For
         End If
      Next i
   End If
   If Not bFound Then
      cboGenre.ListIndex = 0
   End If
   
   ' set genre v2:
   If Len(cboGenreV2.text) > 0 Then
      For i = 0 To cboGenreV2.ListCount - 1
         If StrComp(cboGenreV2.List(i), cboGenreV2.text, vbTextCompare) = 0 Then
            cboGenreV2.ListIndex = i
            Exit For
         End If
      Next i
   End If
   
   
   ' Update ID3 v1 tag:
   With m_cID3v1
      .Artist = txtArtist.text
      .Title = txtTitle.text
      .Album = txtAlbum.text
      .Year = txtYear.text
      .Comment = txtComment.text
      .Track = txtTrackNo.text
      .Genre = cboGenre.ItemData(cboGenre.ListIndex)
      .Update
   End With
   
'''   ' Update ID3 v2 tag:
'''   If UCase(Right(Trim(txtMp3File.text), 3)) = "MP3" Then
'''      With m_cID3v2
'''         .Artist = txtArtist.text   'txtArtistV2.Text
'''         .Title = txtTitle.text     'txtTitleV2.Text
'''         .Album = txtAlbum.text     'txtAlbumV2.Text
'''         .Year = txtYear.text       'txtYearV2.Text
'''         .Comment = txtComment.text 'txtCommentV2.Text
'''   '      If (cboGenreV2.ListIndex > -1) Then
'''   '         .Genre = cboGenreV2.ItemData(cboGenreV2.ListIndex)
'''   '      End If
'''         .Genre = cboGenre.ItemData(cboGenre.ListIndex)
'''         .OtherGenreName = ""  'cboGenreV2.Text
'''         .Track = txtTrackNo.text  'txtTrack.Text
'''         .OriginalArtist = txtArtist.text   'txtOriginalArtist.Text
'''         .Composer = ""   'txtComposer.Text
'''         .Copyright = ""  'txtCopyright.Text
'''         .EncodedBy = ""  'txtEncodedBy.Text
'''         .LinkTo = ""     'txtLinkTo.Text
'''         .Update
'''      End With
'''   End If
   
   Exit Sub
   
err1:
   MsgBox Err.Description
   
   
   
End Sub

Private Sub Form_Load()
   
  Dim FileExt As String

  EnableCloseButton Me.hWnd, False
  
 ' Me.Picture = LoadPicture(App.Path & "\tmpBanner")
 
   lblVersion.Caption = "Version   :    " & App.Major & "." & App.Minor & "." & App.Revision
   lblVersion.Top = 750
   lblVersion.Left = 255
   lblVersion.Width = 2940
   lblVersion.FontSize = 7
 
 ''' HelpContextID = hlpButtons
 
 SSPanel2.BackColor = vbBlack
 SSPanel2.Width = Me.Width + 300
'''
''' For i = Len(FilenameToLoad) To 1 Step -1
'''    If Mid(FilenameToLoad, i, 1) = "." Then
'''      FileExt = Mid(FilenameToLoad, i + 1)
'''      Exit For
'''    End If
''' Next i
  
   txtMp3File.text = FilenameToLoad
   
   If bTagEditMP3 Then
    Set m_cID3v1 = New cMP3ID3v1
    Set m_cID3v2 = New cMP3ID3v2
    
    bTagsUpdated = False
    Dirty = False
    bLoading = True
    
    cboGenre.AddItem ""
    cboGenre.ItemData(cboGenre.NewIndex) = 255
    cboGenreV2.AddItem ""
    cboGenreV2.ItemData(cboGenreV2.NewIndex) = 255
    
    For i = 0 To 255
      If Len(m_cID3v1.GenreName(i)) > 0 Then
         cboGenre.AddItem m_cID3v1.GenreName(i)
         cboGenre.ItemData(cboGenre.NewIndex) = i
         cboGenreV2.AddItem m_cID3v1.GenreName(i)
         cboGenreV2.ItemData(cboGenreV2.NewIndex) = i
      End If
    Next i
    LoadFile txtMp3File.text  'Load the tag info
    EnableControls True
  Else
    'Disable the fields
    EnableControls True
    SSPanel1(1).Enabled = False
    cmdLoadPalette.Enabled = False
    cmdExit.Enabled = True
    txtMp3File.FontBold = True
    Me.Height = 3390
  End If

  
 
  'Me.Show
  Me.Refresh
  
  Dim sCmd As String
  sCmd = Command
  If (Len(sCmd) > 0) Then
    ParseCommand sCmd
  End If
  
  bLoading = False
   
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If (Data.Files.Count > 0) Then
      If (Err.Number = 0) Then
         Effect = vbDropEffectCopy
         LoadFile Data.Files(1)
      End If
   End If
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   '
   On Error Resume Next
   If (Data.Files.Count > 0) Then
      If (Err.Number = 0) Then
         Effect = vbDropEffectCopy
         Exit Sub
      End If
   End If
   Effect = vbDropEffectNone
   '
End Sub

Private Sub mnuFile_Click(Index As Integer)
   Select Case Index
   Case 0
      PickFile
   Case 1
      SetTags
   Case 3
      Unload Me
   End Select
End Sub

Private Sub txtAlbum_Change()
If bLoading Then Exit Sub
Dirty = True
End Sub

Private Sub txtAlbum_GotFocus()
On Error Resume Next
ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(ActiveControl.text)
End Sub

Private Sub txtAlbum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtArtist_Change()
If bLoading Then Exit Sub
Dirty = True
End Sub

Private Sub txtArtist_GotFocus()
On Error Resume Next
ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(ActiveControl.text)
End Sub

Private Sub txtArtist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtComment_Change()
If bLoading Then Exit Sub
Dirty = True
End Sub

Private Sub txtComment_GotFocus()
On Error Resume Next
ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(ActiveControl.text)
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtTitle_Change()
If bLoading Then Exit Sub
Dirty = True
End Sub

Private Sub txtTitle_GotFocus()
On Error Resume Next
ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(ActiveControl.text)
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtTrackNo_GotFocus()
On Error Resume Next
ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(ActiveControl.text)
End Sub

Private Sub txtTrackNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtYear_Change()
If bLoading Then Exit Sub
Dirty = True
End Sub

Private Sub txtYear_GotFocus()
On Error Resume Next
ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(ActiveControl.text)
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
