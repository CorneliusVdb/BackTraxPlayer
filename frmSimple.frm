VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSimple 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple MP3 player using MP3 CLASS"
   ClientHeight    =   5535
   ClientLeft      =   3525
   ClientTop       =   3375
   ClientWidth     =   10335
   Icon            =   "frmSimple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10335
   Begin VB.Frame Frame 
      Caption         =   "ID3v2 tag information:"
      Height          =   1815
      Index           =   4
      Left            =   3720
      TabIndex        =   44
      Top             =   3600
      Width           =   6495
      Begin VB.ListBox TagsV2 
         Height          =   1230
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   6255
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "ID3v1 tag information:"
      Height          =   2295
      Index           =   3
      Left            =   3720
      TabIndex        =   27
      Top             =   1200
      Width           =   6495
      Begin VB.CommandButton butV1save 
         Caption         =   "Save &ID3v1"
         Height          =   375
         Left            =   4800
         TabIndex        =   43
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox V1genre 
         Height          =   315
         Left            =   4680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.HScrollBar V1trackS 
         Height          =   255
         LargeChange     =   25
         Left            =   5040
         Max             =   50
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox V1comments 
         Height          =   645
         Left            =   960
         MaxLength       =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox V1year 
         Height          =   285
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   42
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox V1album 
         Height          =   285
         Left            =   960
         MaxLength       =   30
         TabIndex        =   37
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox V1artist 
         Height          =   285
         Left            =   960
         MaxLength       =   30
         TabIndex        =   36
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox V1title 
         Height          =   285
         Left            =   960
         MaxLength       =   30
         TabIndex        =   35
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label V1track 
         Caption         =   "0"
         Height          =   255
         Left            =   4680
         TabIndex        =   39
         Top             =   360
         Width           =   375
      End
      Begin VB.Label txt 
         Caption         =   "Genre:"
         Height          =   255
         Index           =   13
         Left            =   3960
         TabIndex        =   33
         Top             =   720
         Width           =   735
      End
      Begin VB.Label txt 
         Caption         =   "Track:"
         Height          =   255
         Index           =   12
         Left            =   3960
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label txt 
         Caption         =   "Comments:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label txt 
         Caption         =   "Year:"
         Height          =   255
         Index           =   10
         Left            =   3960
         TabIndex        =   34
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label txt 
         Caption         =   "Album:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label txt 
         Caption         =   "Artist:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.Label txt 
         Caption         =   "Title:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.OptionButton DisplayMode2 
      Caption         =   "minutes && seconds"
      Height          =   255
      Left            =   1680
      TabIndex        =   26
      Top             =   5160
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton DisplayMode 
      Caption         =   "milliseconds"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Timer StatusUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   600
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame 
      Caption         =   "Control:"
      Height          =   975
      Index           =   2
      Left            =   3720
      TabIndex        =   16
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton butStop 
         Caption         =   "&Stop"
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton butResume 
         Caption         =   "&Resume"
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton butPause 
         Caption         =   "P&ause"
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton butPlay 
         Caption         =   "&Play"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Status:"
      Height          =   3375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
      Begin VB.Label txt 
         Caption         =   "Last error:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label LastError 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label Remaining 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Position 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Length 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Errors 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Status 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Filename 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label txt 
         Caption         =   "Remaining:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label txt 
         Caption         =   "Position:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label txt 
         Caption         =   "Length:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label txt 
         Caption         =   "Errors:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label txt 
         Caption         =   "Status:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label txt 
         Caption         =   "Filename:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "File:"
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton butClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton butOpen 
         Caption         =   "&Open"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton butFile 
         Caption         =   "..."
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtFile 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' A simple example program
' Using MP3_CLASS by Vesa Piittinen aka Merri
' http://merri.net
'
' What could be done better and what is missing:
' - have information on if a file is opened
' - the error handling on the client side
' - seeking
' - well, the looks and the overall design :P
'
' Enjoy!

Dim MyMP3 As New MP3_CLASS
Private Function GetMinSec(ByVal Value As Long) As String
    Dim Seconds As Single
    Seconds = CSng(Value) / 1000
    GetMinSec = CStr(Seconds \ 60) & ":" & Format$(Seconds Mod 60, "00")
End Function
Private Function StatusToText(ByVal Index As MP3_CLASS_STATUS) As String
    'returns status numbers as text
    Select Case Index
        Case -1
            StatusToText = "error"
        Case 0
            StatusToText = "unknown"
        Case 1
            StatusToText = "not ready"
        Case 2
            StatusToText = "playing"
        Case 3
            StatusToText = "paused"
        Case 4
            StatusToText = "stopped"
        Case 5
            StatusToText = "seeking"
        Case 6
            StatusToText = "recording"
    End Select
End Function
Private Sub butClose_Click()
    'disable statusupdate
    StatusUpdate.Enabled = False
    'stop playback
    butStop_Click
    'close and check for error
    If Not MyMP3.CloseMP3 Then LastError = MyMP3.GetLastError
End Sub
Private Sub butFile_Click()
    On Error Resume Next
    'set dialog properties
    Dialog.DialogTitle = "Open MP3"
    Dialog.Filter = "MP3 file (*.mp3)|*.mp3"
    'show dialog
    Dialog.ShowOpen
    'CancelError = True, if user selects cancel we get an error
    If Err Then Exit Sub
    'set filename
    txtFile.Text = Dialog.Filename
End Sub
Private Sub butOpen_Click()
    Dim Title As String, Artist As String, Album As String, Year As Integer, Comment As String, Track As Byte, Genre As Byte
    Dim ID As String, Data As String
    Dim A As Byte
    'check if a file exists
    If txtFile.Text = "" Then Exit Sub
    If StatusUpdate.Enabled Then butClose_Click
    'open and play, if can't report error
    If MyMP3.OpenMP3(txtFile.Text) Then
        StatusUpdate.Enabled = True
        Filename = MyMP3.GetFilename
        butPlay_Click
        'read ID3v1 tag information
        MyMP3.GetTagV1 Title, Artist, Album, Year, Comment, Track, Genre
        'view tag information
        V1title.Text = Title
        V1artist.Text = Artist
        V1album.Text = Album
        V1year.Text = CStr(Year)
        V1comments.Text = Comment
        V1trackS.Value = CInt(Track)
        For A = 0 To V1genre.ListCount - 1
            If V1genre.ItemData(A) = Genre Then V1genre.ListIndex = A: Exit For
        Next A
        'read and view ID3v2 tag information
        TagsV2.Clear
        For A = 1 To MyMP3.TagsV2
            'read tag
            MyMP3.GetTagV2 A, ID, Data
            'add to list
            TagsV2.AddItem ID & " : " & MyMP3.ConvertText(Data)
        Next A
    Else
        LastError = MyMP3.GetLastError
    End If
End Sub
Private Sub butPause_Click()
    'pause and check for error
    If Not MyMP3.PauseMP3 Then LastError = MyMP3.GetLastError
End Sub
Private Sub butPlay_Click()
    'play and check for error
    If Not MyMP3.PlayMP3 Then LastError = MyMP3.GetLastError
End Sub
Private Sub butResume_Click()
    'resume and check for error
    If Not MyMP3.ResumeMP3 Then LastError = MyMP3.GetLastError
End Sub
Private Sub butStop_Click()
    'stop and check for error
    If Not MyMP3.StopMP3 Then LastError = MyMP3.GetLastError
End Sub
Private Sub butV1save_Click()
    If Not MyMP3.SetTagV1(V1title.Text, V1artist.Text, V1album.Text, Val(V1year.Text), V1comments.Text, CByte(V1trackS.Value), CByte(V1genre.ItemData(V1genre.ListIndex))) Then LastError = MyMP3.GetLastError
End Sub
Private Sub Form_Load()
    Dim A As Byte, Temp As String
    'get all genres to combobox
    V1genre.AddItem " none"
    V1genre.ItemData(V1genre.NewIndex) = 255
    V1genre.ListIndex = V1genre.NewIndex
    For A = 0 To 254
        Temp = MyMP3.GetGenre(A)
        If Temp <> "" Then V1genre.AddItem Temp: V1genre.ItemData(V1genre.NewIndex) = A Else Exit For
    Next A
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'close file before exit
    butClose_Click
End Sub
Private Sub StatusUpdate_Timer()
    On Error Resume Next
    'update status messages
    Status = StatusToText(MyMP3.GetStatus)
    Errors = MyMP3.Errors
    If LastError <> MyMP3.GetLastError Then LastError = MyMP3.GetLastError
    If Val(Status) > -1 Then
        'check if to show in milliseconds or minutes and seconds
        If DisplayMode.Value Then
            'milliseconds
            Length = MyMP3.Length & " ms"
            Position = MyMP3.Position & " ms"
            Remaining = MyMP3.Remaining & " ms"
        Else
            'minutes and seconds
            Length = GetMinSec(MyMP3.Length)
            Position = GetMinSec(MyMP3.Position)
            Remaining = GetMinSec(MyMP3.Remaining)
        End If
        If Status = StatusToText(mp3classstopped) And Length = Position Then MyMP3.SeekTo 0: MyMP3.PlayMP3
    End If
End Sub
Private Sub V1trackS_Change()
    V1track = V1trackS.Value
End Sub
Private Sub V1trackS_Scroll()
    V1trackS_Change
End Sub
