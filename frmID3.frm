VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmID3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ID3 Tags"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdateID3v1tag 
      Caption         =   "Update"
      Height          =   495
      Left            =   4995
      TabIndex        =   37
      ToolTipText     =   "Update the ID3v1 tag"
      Top             =   7380
      Width           =   1125
   End
   Begin VB.CheckBox chkv1 
      BackColor       =   &H00000000&
      Caption         =   "Has ID3v1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   90
      TabIndex        =   30
      ToolTipText     =   "Checked = contains ID3v1 info"
      Top             =   4440
      Width           =   1605
   End
   Begin VB.CheckBox chkv2 
      BackColor       =   &H00000000&
      Caption         =   "Has ID3v2"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   90
      TabIndex        =   29
      ToolTipText     =   "Checked = contains ID3v2 info"
      Top             =   870
      Width           =   2055
   End
   Begin VB.Frame fraID3v1 
      BackColor       =   &H00000000&
      Caption         =   "ID3v1 tag info"
      ForeColor       =   &H00FFFF00&
      Height          =   2625
      Left            =   90
      TabIndex        =   16
      Top             =   4680
      Width           =   6045
      Begin VB.ComboBox cboGenre1 
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1380
         Width           =   2145
      End
      Begin VB.TextBox txtTrack1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   22
         Top             =   210
         Width           =   4185
      End
      Begin VB.TextBox txtArtist1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   21
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtAlbum1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   20
         Top             =   990
         Width           =   4185
      End
      Begin VB.TextBox txtGenre1 
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   19
         Top             =   1380
         Width           =   555
      End
      Begin VB.TextBox txtYear1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1770
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1770
         Width           =   4185
      End
      Begin VB.TextBox txtComments1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   17
         Top             =   2160
         Width           =   4185
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "genre name:"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   2520
         TabIndex        =   32
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "song:"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   240
         TabIndex        =   28
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "artist:"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   240
         TabIndex        =   27
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "album:"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   240
         TabIndex        =   26
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "genre #:"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   900
         TabIndex        =   25
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "year:"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   240
         TabIndex        =   24
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "comments:"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   2220
         Width           =   1455
      End
   End
   Begin VB.Frame fraID3v2 
      BackColor       =   &H00000000&
      Caption         =   "ID3v2 tag info"
      ForeColor       =   &H0000FF00&
      Height          =   3195
      Left            =   90
      TabIndex        =   1
      Top             =   1140
      Width           =   6045
      Begin VB.ComboBox cboGenre2 
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1470
         Width           =   2145
      End
      Begin VB.TextBox txtTrackNbr 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1770
         TabIndex        =   14
         Top             =   2640
         Width           =   585
      End
      Begin VB.TextBox txtTrack 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1770
         TabIndex        =   7
         Top             =   330
         Width           =   4185
      End
      Begin VB.TextBox txtArtist 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1770
         TabIndex        =   6
         Top             =   715
         Width           =   4185
      End
      Begin VB.TextBox txtAlbum 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1770
         TabIndex        =   5
         Top             =   1100
         Width           =   4185
      End
      Begin VB.TextBox txtGenre 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1770
         TabIndex        =   4
         Top             =   1485
         Width           =   555
      End
      Begin VB.TextBox txtYear 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1770
         TabIndex        =   3
         Top             =   1870
         Width           =   4185
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1770
         TabIndex        =   2
         Top             =   2255
         Width           =   4185
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "genre name:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   2430
         TabIndex        =   33
         Top             =   1530
         Width           =   1245
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "track nbr:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   840
         TabIndex        =   15
         Top             =   2670
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "song:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   270
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "artist:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   270
         TabIndex        =   12
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "album:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   270
         TabIndex        =   11
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "genre #:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   930
         TabIndex        =   10
         Top             =   1530
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "year:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   270
         TabIndex        =   9
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "comments:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   270
         TabIndex        =   8
         Top             =   2310
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5580
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Get ID3 info for MP3 file..."
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      ToolTipText     =   "Select an MP3 you want to read."
      Top             =   45
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmID3.frx":0000
      ForeColor       =   &H0000FFFF&
      Height          =   750
      Left            =   870
      TabIndex        =   36
      Top             =   60
      Width           =   5130
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      TabIndex        =   35
      Top             =   45
      Width           =   855
   End
End
Attribute VB_Name = "frmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Last Modified: April 26, 2001
' Contact: Kevin Pohl
' email: kevinpohl@threefifteen.net
'
' Code Purpose:
' 1. To read ID3v1 tag info
' 2. To read ID3v2 tag info
'
' Basically, I'm trying to extract this info from ID3v1 and
' ID3v2 tags:
' 1. Track Title
' 2. Artist
' 3. Album
' 4. Genre number
' 5. Genre name (use the genre number to get the genre name)
' 6. Track number (so far, only for ID3v2)
' 7. Comments
' 8. Year

' Missing features:
' 1. Doesn't update ID3v1 or ID3v2 tags (I have the code to
'    update ID3v1 tags, but haven't plugged it in yet.
'
' Features to be added:
' 1. Writing/Updating ID3v1 tags
' 2. Writing ID3v2 tags
' 3. This code currently works after selecting a file with a
'    common dialog.  I'm going to expand on this code to
'    batch read MP3 files after selecting MP3s in a list box
'    that is populated from a recursive directory search.
'    Obviously, this code can serve as the foundation for that
'    but will have to be modified.
'
' Wish list:
' 1. To be able to write/update more than the basic tag info
'    for ID3v2 tags, including image, lyrics, MusicMatch fields
'    (Tempo, Situation, Preference, Mood), and URL link fields.
' 2. To differentiate the difference ID3v1 tags that contain a
'    track number and those that do not.
'
'
' Known issues:
' 1. This code isn't reading the 'Comment' field of ID3v2
'    tags properly.
' 2. This code isn't reading ID3 tags of MP3 files that
'    were created by RealJukebox.  RJB does something
'    different that I haven't totally figured out yet.
'    RJB seems to fill the tag with a bunch of info
'    in the 'GEOB" frame of the tag, which makes it
'    difficult to extract the info.  Most of the software
'    I've used that reads tags doesn't read RJB MP3 file
'    tags.
' 3. Doesn't read or write ID3v2 tags from Windows Media
'    Audio files.
' 4. Some MP3-related software (like MusicMatch) tags
'    the ID3v2 genre as "(17) General Rock".  I've written
'    some code that parses the ID3v2 genre to determine if
'    parentheses exists in the genre "()" and to extract
'    only the genre number.  I haven't plugged this feature
'    into this code yet.
''
' Other comments:
' 1. I'm using VB 6.0 Enterprise Edition w/SP4
' 2. This code should work with any version of VB with
'    any SP.
' 3. A compile exe of this code should run on all
'    operating systems (Win95/98/Me/NT/2000)
'
' Code sources:
' 1. Code for reading and writing ID3v1 tags is basically
'    all the same.  You can find variations on:
'    http://www.planet-source-code
'    http://www.freevbcode.com
' 2. I found code written by a guy with the alias of
'    "The Frog Prince".  He has written VB code that utilizes
'    a DLL named VBID3LIB.dll.  It can read and write ID3v2
'    tags.  But it's main weakness is that it will read the
'    ID3v1 tag instead of the ID3v2 tag if the MP3 contains
'    both ID3v2 AND an ID3v2 tag.  You can find the code at
'    his web site:
'    http://members.tripod.com/thefrogprince/id3lib.htm
' 3. There's also some code contained in a VB project
'    written by "Joe Hart".  His code can be found on various
'    VB source code sites, including http://www.freevbcode.com
'    This is the link to his project at freevbcode:
'    http://www.freevbcode.com/ShowCode.Asp?ID=1127
'    This code reads ID3v1 and ID3v2 tags, and supposedly
'    writes ID3v1 tags.  His project is an MP3 player and file
'    manager.  But I found a lot of problems with it and very
'    difficult to follow.  I've extracted what I believe is the
'    "good stuff" in his code to read tags and have include it
'    in this project.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public i As Integer
Public strEmptyString As String
Public B As Byte
Public s As String

Private Type ID3v1Tag
  id As String * 3
  title As String * 30
  Artist As String * 30
  Album As String * 30
  Year As String * 4
  Comment As String * 30
  Genre As Byte
End Type

Private Version As Byte

Private Sub cmdBrowse_Click()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the CommonDialog to choose an MP3 file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim strFilePath As String
    strFilePath = FileChoice(0, "Choose MP3 file")
    If strFilePath <> "" Then
        Me.Caption = strFilePath
    End If
'    FilenameToLoad
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clear all fields on the window
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ID3v1 text fields
        chkv1.value = 0         'This becomes checked if the file has an ID3v1 tag
        txtTrack1.Text = ""     'The file's track title (limited to 30 characters)
        txtArtist1.Text = ""    'The file's artist (limited to 30 characters)
        txtAlbum1.Text = ""     'The file's album (limited to 30 characters)
        txtGenre1.Text = ""     'The file's genre
        txtYear1.Text = ""      'The file's year (limited to 4 characters)
        txtComments1.Text = ""  'The file's comments (limited to 30 characters)
        
    'ID3v2 text fields
        chkv2.value = 0         'This becomes checked if the file has an ID3v2 tag
        txtTrack.Text = ""      'The file's track title (not limited to 30 characters)
        txtArtist.Text = ""     'The file's artist (not limited to 30 characters)
        txtAlbum.Text = ""      'The file's album (not limited to 30 characters)
        txtGenre.Text = ""      'The file's genre (not limited to 30 characters)
        txtYear.Text = ""       'The file's year (not limited to 4 characters)
        txtComments.Text = ""   'The file's comments (not limited to 30 characters)
        txtTrackNbr.Text = ""   'The file's track number (not limited to 30 characters)
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Read the file to extract any existing ID3v1 and ID3v2 tags
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReadFile

End Sub


Private Function ReadFile()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function:
' 1) Opens the file that was selected in the common dialog.
' 2) Checks for a valid ID3 header.
' 3) Extracts any ID3v1 tag info (and displays it on window)
' 4) Extracts anu ID3v2 tag info (and displays it on window)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo errorHandler

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' use the filename to get ID3 info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim strFilename As String
    Dim lngFilesize As Long
    
    strFilename = FilenameToLoad  'Me.Caption

    Dim fn As Integer
    Dim lngHeaderPosition As Long
    Dim Tag1 As ID3v1Tag
    Dim Tag2 As String
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    fn = FreeFile
    
    Open strFilename For Binary As #fn                      'Open the file so we can read it
    lngFilesize = LOF(fn)                                   'Size of the file, in bytes

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check for a Header
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
        Get #fn, 1, B
            
        If B <> 255 Then '(255 is where an ID3v2 header should start)
            If B <> 73 Then
                'Exit Function
            End If
        End If
         
        lngHeaderPosition = 1
        Get #fn, 2, B
        If (B < 250 Or B > 251) Then
            'We have an ID3v2 tag
            If B = 68 Then
                Get #fn, 3, B
                If B = 51 Then
                    Dim R As Double
                    Get #fn, 4, Version
                    Get #fn, 7, B
                    R = B * 20917152
                    Get #fn, 8, B
                    R = R + (B * 16384)
                    Get #fn, 9, B
                    R = R + (B * 128)
                    Get #fn, 10, B
                    R = R + B
                    If R > lngFilesize Or R > 2147483647 Then
                        Exit Function
                    End If
                    Tag2 = Space$(R)
                    Get #fn, 11, Tag2
                    lngHeaderPosition = R + 11
                End If
            End If
        Else
            'ID3v2 tag is missing
        End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check for an ID3v1 tag
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'ID3v1 tag
        Get #fn, lngFilesize - 127, Tag1.id
        
        If Tag1.id = "TAG" Then 'If "TAG" is present, then we have a valid ID3v1 tag and will extract all available ID3v1 info from the file
            Get #fn, , Tag1.title   'Always limited to 30 characters
            Get #fn, , Tag1.Artist  'Always limited to 30 characters
            Get #fn, , Tag1.Album   'Always limited to 30 characters
            Get #fn, , Tag1.Year    'Always limited to 4 characters
            Get #fn, , Tag1.Comment 'Always limited to 30 characters
            Get #fn, , Tag1.Genre   'Always limited to 1 byte (?)
            
            frmID3.chkv1.value = 1 'Indicates that the file contains ID3v1 tag info
    
            'Populate the form with the ID3v1 info
            With frmID3
                txtTrack1.Text = Trim$(Tag1.title)
                txtArtist1.Text = Trim$(Tag1.Artist)
                txtAlbum1.Text = Trim$(Tag1.Album)
                txtYear1.Text = Trim$(Tag1.Year)
                txtComments1.Text = Trim$(Tag1.Comment)
                txtGenre1.Text = Tag1.Genre
            End With
            
            cboGenre1.ListIndex = Tag1.Genre + 1
        End If
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Proceed to extract the ID3v2 tag info if any exists
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
        If Tag2 <> strEmptyString Then
            frmID3.chkv2.value = 1
            GetID3v2Tag1 (Tag2) 'Pass the Id3v2 TagId to the GetID3v2Tag1 function
        End If
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Close the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Close

    Exit Function
        
errorHandler:
    'MsgBox "Error reading file"
    Err.Clear
    Close
    Resume Next
End Function

Private Function GetID3v2Tag1(Tag2 As String) As Boolean

On Error GoTo errorHandler

   Dim TitleField As String
   Dim ArtistField As String
   Dim AlbumField As String
   Dim YearField As String
   Dim GenreField As String
   Dim FieldSize As Long
   Dim SizeOffset As Long
   Dim FieldOffset As Long
   Dim TrackNbr As String
   Dim SituationField As String
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Determine if the ID3v2 tag is ID3v2.2 or ID3v2.3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Notes: I haven't tested reading an MP3 file that has a ID3v2.2 tag
   
    Select Case Version
    
        Case 2 'ID3v2.2
        'Set the fieldnames for version 2.0
            TitleField = "TT2"
            ArtistField = "TOA"
            AlbumField = "TAL"
            YearField = "TYE"
            GenreField = "TCO"
            FieldOffset = 7
            SizeOffset = 5
            TrackNbr = "TRCK"
       
        Case 3 'ID3v2.3
        'Set the fieldnames for version 3.0
            TitleField = "TIT2"
            ArtistField = "TPE1"
            AlbumField = "TALB"
            YearField = "TYER"
            GenreField = "TCON"
            TrackNbr = "TRCK"
       
            FieldOffset = 11
            SizeOffset = 7
        Case Else
        'We don't have a valid ID3v2 tag, so bail
            Exit Function
            
    End Select
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract track title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       i = InStr(Tag2, TitleField)
       If i > 0 Then
          'read the title
          FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
          If Version = 3 Then
             'check for compressed or encrypted field
             B = Asc(Mid$(Tag2, i + 9))
             If (B And 128) = True Or (B And 64) = True Then GoTo ReadAlbum
          End If
          txtTrack.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
       End If
       
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract album title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadAlbum:
    i = InStr(Tag2, AlbumField)
    If i > 0 Then
       FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
       If Version = 3 Then
          'check for compressed or encrypted field
          B = Asc(Mid$(Tag2, i + 9))
          If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadArtist
       End If
       txtAlbum.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
       
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract artist name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadArtist:
   i = InStr(Tag2, ArtistField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadYear
      End If
      txtArtist.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract year title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadYear:
   i = InStr(Tag2, YearField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadGenre
      End If
      txtYear.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract genre
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadGenre:
   i = InStr(Tag2, GenreField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      
      If Version = 3 Then
         'check for compressed or encrypted field
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadTrackNbr
      End If
      
      s = Mid$(Tag2, i + FieldOffset, FieldSize)
      If Left$(s, 1) = "(" Then
        txtGenre.Text = Val(Mid$(s, 2, 2))
        cboGenre2.ListIndex = Val(txtGenre.Text) + 1
        
      Else
         'i = InStr(gsGenres, s & Space$(22 - Len(s)))
         txtGenre.Text = i
         
         cboGenre2.ListIndex = i
         If i > 0 Then
            txtGenre.Text = Int(i / 22)
         End If
      End If
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract track number
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadTrackNbr:
   i = InStr(Tag2, TrackNbr)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo Done
      End If
      txtTrackNbr.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We're done looking for ID3v2 info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Done:
   
   Exit Function

errorHandler:
   Err.Clear
   Resume Next
End Function

Private Function FileChoice(iFileType As Integer, title As String) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Prepares the common dialog
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo errorHandler

    With CommonDialog1
        .DialogTitle = "Choose an MP3 audio file"           'Sets the caption of the ShowOpen common dialog
        .CancelError = True                                 'Sets a value indicating whether an error is generated when the user chooses the Cancel button.
        .Filter = "MP3 (*.mp3)|*.mp3|"                      'Sets the filters that are displayed in the Typelist box of a dialog box.
        .DefaultExt = ".mp3"                                'Sets the default filename extension for the dialog box.
        .FilterIndex = 2                                    'Sets the default filter for the dialog box.
        .ShowOpen                                           'Displays the CommonDialog control's Open dialog box.
    End With
    
    FileChoice = CommonDialog1.FileName                     'Returns the path and filename of a selected file
    Exit Function
    
errorHandler:
    'MsgBox "Error selecting MP3 audio file."
    Err.Clear
    Resume Next
    
End Function

Private Sub Form_Load()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Call the function to populate the ID3v1 and ID3v2 genre
' combo boxes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    PopulateGenres
    
    LoadFileInfo

End Sub

Sub LoadFileInfo()
'    Dim strFilePath As String
'    strFilePath = FileChoice(0, "Choose MP3 file")
'    'If strFilePath <> "" Then
'        Me.Caption = strFilePath
'    'End If
   lblFileName.Caption = FilenameToLoad
   'Me.Caption = FilenameToLoad
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clear all fields on the window
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ID3v1 text fields
        chkv1.value = 0         'This becomes checked if the file has an ID3v1 tag
        txtTrack1.Text = ""     'The file's track title (limited to 30 characters)
        txtArtist1.Text = ""    'The file's artist (limited to 30 characters)
        txtAlbum1.Text = ""     'The file's album (limited to 30 characters)
        txtGenre1.Text = ""     'The file's genre
        txtYear1.Text = ""      'The file's year (limited to 4 characters)
        txtComments1.Text = ""  'The file's comments (limited to 30 characters)
        
    'ID3v2 text fields
        chkv2.value = 0         'This becomes checked if the file has an ID3v2 tag
        txtTrack.Text = ""      'The file's track title (not limited to 30 characters)
        txtArtist.Text = ""     'The file's artist (not limited to 30 characters)
        txtAlbum.Text = ""      'The file's album (not limited to 30 characters)
        txtGenre.Text = ""      'The file's genre (not limited to 30 characters)
        txtYear.Text = ""       'The file's year (not limited to 4 characters)
        txtComments.Text = ""   'The file's comments (not limited to 30 characters)
        txtTrackNbr.Text = ""   'The file's track number (not limited to 30 characters)
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Read the file to extract any existing ID3v1 and ID3v2 tags
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReadFile
End Sub

Sub PopulateGenres()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate the ID3v1 and ID3v2 genre combo boxes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Notes: There may be an easier way to populate the Id3v1 and
' ID3v2 genre combo boxes with a multidimensional array.
' I'm doing it this way so that:
' 1. The genres are listed in the combo box alphabetically.
' 2. Even though the genres are listed alphabetically, they
'    have a listindex which corresponds to the standardized
'    genre list, i.e., Blues = 0, Classic Rock = 1, etc.
' 3. I add 1 to the listindex so that we have a selectable
'    list item in the combo boxes that is blank (like Winamp).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Populate the ID3v1 genre combo box

    With cboGenre1
    
        .AddItem Chr(0)
        .ItemData(cboGenre1.NewIndex) = 0
        
        .AddItem "Blues"
        .ItemData(cboGenre1.NewIndex) = 1
        
        .AddItem "Classic Rock"
        .ItemData(cboGenre1.NewIndex) = 2
        
        .AddItem "Country"
        .ItemData(cboGenre1.NewIndex) = 3
        
        .AddItem "Dance"
        .ItemData(cboGenre1.NewIndex) = 4
        
        .AddItem "Disco"
        .ItemData(cboGenre1.NewIndex) = 5
        
        .AddItem "Funk"
        .ItemData(cboGenre1.NewIndex) = 6
        
        .AddItem "Grunge"
        .ItemData(cboGenre1.NewIndex) = 7
        
        .AddItem "Hip-Hop"
        .ItemData(cboGenre1.NewIndex) = 8
        
        .AddItem "Jazz"
        .ItemData(cboGenre1.NewIndex) = 9
        
        .AddItem "Metal"
        .ItemData(cboGenre1.NewIndex) = 10
        
        .AddItem "New Age"
        .ItemData(cboGenre1.NewIndex) = 11
        
        .AddItem "Oldies"
        .ItemData(cboGenre1.NewIndex) = 12
        
        .AddItem "Other"
        .ItemData(cboGenre1.NewIndex) = 13
        
        .AddItem "Pop"
        .ItemData(cboGenre1.NewIndex) = 14
        
        .AddItem "R&B"
        .ItemData(cboGenre1.NewIndex) = 15
        
        .AddItem "Rap"
        .ItemData(cboGenre1.NewIndex) = 16
        
        .AddItem "Reggae"
        .ItemData(cboGenre1.NewIndex) = 17
        
        .AddItem "Rock"
        .ItemData(cboGenre1.NewIndex) = 18
        
        .AddItem "Techno"
        .ItemData(cboGenre1.NewIndex) = 19
        
        .AddItem "Industrial"
        .ItemData(cboGenre1.NewIndex) = 20
        
        .AddItem "Alternative"
        .ItemData(cboGenre1.NewIndex) = 21
        
        .AddItem "Ska"
        .ItemData(cboGenre1.NewIndex) = 22
        
        .AddItem "Death Metal"
        .ItemData(cboGenre1.NewIndex) = 23
        
        .AddItem "Pranks"
        .ItemData(cboGenre1.NewIndex) = 24
        
        .AddItem "Soundtrack"
        .ItemData(cboGenre1.NewIndex) = 25
        
        .AddItem "Euro-Techno"
        .ItemData(cboGenre1.NewIndex) = 26
        
        .AddItem "Ambient"
        .ItemData(cboGenre1.NewIndex) = 27
        
        .AddItem "Trip-Hop"
        .ItemData(cboGenre1.NewIndex) = 28
        
        .AddItem "Vocal"
        .ItemData(cboGenre1.NewIndex) = 29
        
        .AddItem "Jazz+Funk"
        .ItemData(cboGenre1.NewIndex) = 30
        
        .AddItem "Fusion"
        .ItemData(cboGenre1.NewIndex) = 31
        
        .AddItem "Trance"
        .ItemData(cboGenre1.NewIndex) = 32
        
        .AddItem "Classical"
        .ItemData(cboGenre1.NewIndex) = 33
        
        .AddItem "Instrumental"
        .ItemData(cboGenre1.NewIndex) = 34
        
        .AddItem "Acid"
        .ItemData(cboGenre1.NewIndex) = 35
        
        .AddItem "House"
        .ItemData(cboGenre1.NewIndex) = 36
        
        .AddItem "Game"
        .ItemData(cboGenre1.NewIndex) = 37
        
        .AddItem "Sound Clip"
        .ItemData(cboGenre1.NewIndex) = 38
        
        .AddItem "Gospel"
        .ItemData(cboGenre1.NewIndex) = 39
        
        .AddItem "Noise"
        .ItemData(cboGenre1.NewIndex) = 40
        
        .AddItem "AlternRock"
        .ItemData(cboGenre1.NewIndex) = 41
        
        .AddItem "Bass"
        .ItemData(cboGenre1.NewIndex) = 42
        
        .AddItem "Soul"
        .ItemData(cboGenre1.NewIndex) = 43
        
        .AddItem "Punk"
        .ItemData(cboGenre1.NewIndex) = 44
        
        .AddItem "Space"
        .ItemData(cboGenre1.NewIndex) = 45
        
        .AddItem "Meditative"
        .ItemData(cboGenre1.NewIndex) = 46
        
        .AddItem "Instrumental Pop"
        .ItemData(cboGenre1.NewIndex) = 47
        
        .AddItem "Instrumental Rock"
        .ItemData(cboGenre1.NewIndex) = 48
        
        .AddItem "Ethnic"
        .ItemData(cboGenre1.NewIndex) = 49
        
        .AddItem "Gothic"
        .ItemData(cboGenre1.NewIndex) = 50
        
        .AddItem "Darkwave"
        .ItemData(cboGenre1.NewIndex) = 51
        
        .AddItem "Techno-Industrial"
        .ItemData(cboGenre1.NewIndex) = 52
        
        .AddItem "Electronic"
        .ItemData(cboGenre1.NewIndex) = 53
        
        .AddItem "Pop-Folk"
        .ItemData(cboGenre1.NewIndex) = 54
        
        .AddItem "Eurodance"
        .ItemData(cboGenre1.NewIndex) = 55
        
        .AddItem "Dream"
        .ItemData(cboGenre1.NewIndex) = 56
        
        .AddItem "Southern Rock"
        .ItemData(cboGenre1.NewIndex) = 57
        
        .AddItem "Comedy"
        .ItemData(cboGenre1.NewIndex) = 58
        
        .AddItem "Cult"
        .ItemData(cboGenre1.NewIndex) = 59
        
        .AddItem "Gangsta"
        .ItemData(cboGenre1.NewIndex) = 60
        
        .AddItem "Top 40"
        .ItemData(cboGenre1.NewIndex) = 61
        
        .AddItem "Christian Rap"
        .ItemData(cboGenre1.NewIndex) = 62
        
        .AddItem "Pop/Funk"
        .ItemData(cboGenre1.NewIndex) = 63
        
        .AddItem "Jungle"
        .ItemData(cboGenre1.NewIndex) = 64
        
        .AddItem "Native American"
        .ItemData(cboGenre1.NewIndex) = 65
        
        .AddItem "Cabaret"
        .ItemData(cboGenre1.NewIndex) = 66
        
        .AddItem "New Wave"
        .ItemData(cboGenre1.NewIndex) = 67
        
        .AddItem "Psychadelic"
        .ItemData(cboGenre1.NewIndex) = 68
        
        .AddItem "Rave"
        .ItemData(cboGenre1.NewIndex) = 69
        
        .AddItem "Showtunes"
        .ItemData(cboGenre1.NewIndex) = 70
        
        .AddItem "Trailer"
        .ItemData(cboGenre1.NewIndex) = 71
        
        .AddItem "Lo-Fi"
        .ItemData(cboGenre1.NewIndex) = 72
        
        .AddItem "Tribal"
        .ItemData(cboGenre1.NewIndex) = 73
        
        .AddItem "Acid Punk"
        .ItemData(cboGenre1.NewIndex) = 74
        
        .AddItem "Acid Jazz"
        .ItemData(cboGenre1.NewIndex) = 75
        
        .AddItem "Polka"
        .ItemData(cboGenre1.NewIndex) = 76
        
        .AddItem "Retro"
        .ItemData(cboGenre1.NewIndex) = 77
        
        .AddItem "Musical"
        .ItemData(cboGenre1.NewIndex) = 78
        
        .AddItem "Rock & Roll"
        .ItemData(cboGenre1.NewIndex) = 79
        
        .AddItem "Hard Rock"
        .ItemData(cboGenre1.NewIndex) = 80
        
        .AddItem "Folk"
        .ItemData(cboGenre1.NewIndex) = 81
        
        .AddItem "Folk-Rock"
        .ItemData(cboGenre1.NewIndex) = 82
        
        .AddItem "National Folk"
        .ItemData(cboGenre1.NewIndex) = 83
        
        .AddItem "Swing"
        .ItemData(cboGenre1.NewIndex) = 84
        
        .AddItem "Fast Fusion"
        .ItemData(cboGenre1.NewIndex) = 85
        
        .AddItem "Bebop"
        .ItemData(cboGenre1.NewIndex) = 86
        
        .AddItem "Latin"
        .ItemData(cboGenre1.NewIndex) = 87
        
        .AddItem "Revival"
        .ItemData(cboGenre1.NewIndex) = 88
        
        .AddItem "Celtic"
        .ItemData(cboGenre1.NewIndex) = 89
        
        .AddItem "Bluegrass"
        .ItemData(cboGenre1.NewIndex) = 90
        
        .AddItem "Avantgarde"
        .ItemData(cboGenre1.NewIndex) = 91
        
        .AddItem "Gothic Rock"
        .ItemData(cboGenre1.NewIndex) = 92
        
        .AddItem "Progressive Rock"
        .ItemData(cboGenre1.NewIndex) = 93
        
        .AddItem "Psychedlic Rock"
        .ItemData(cboGenre1.NewIndex) = 94
        
        .AddItem "Symphonic Rock"
        .ItemData(cboGenre1.NewIndex) = 95
        
        .AddItem "Slow Rock"
        .ItemData(cboGenre1.NewIndex) = 96
    
        .AddItem "Big Band"
        .ItemData(cboGenre1.NewIndex) = 97
        
        .AddItem "Chorus"
        .ItemData(cboGenre1.NewIndex) = 98
        
        .AddItem "Easy Listening"
        .ItemData(cboGenre1.NewIndex) = 99
        
        .AddItem "Acoustic"
        .ItemData(cboGenre1.NewIndex) = 100
        
        .AddItem "Humour"
        .ItemData(cboGenre1.NewIndex) = 101
        
        .AddItem "Speech"
        .ItemData(cboGenre1.NewIndex) = 102
        
        .AddItem "Chanson"
        .ItemData(cboGenre1.NewIndex) = 103
        
        .AddItem "Opera"
        .ItemData(cboGenre1.NewIndex) = 104
        
        .AddItem "Chamber Music"
        .ItemData(cboGenre1.NewIndex) = 105
        
        .AddItem "Sonota"
        .ItemData(cboGenre1.NewIndex) = 106
        
        .AddItem "Symphony"
        .ItemData(cboGenre1.NewIndex) = 107
        
        .AddItem "Booty Bass"
        .ItemData(cboGenre1.NewIndex) = 108
        
        .AddItem "Primus"
        .ItemData(cboGenre1.NewIndex) = 109
        
        .AddItem "Porn Groove"
        .ItemData(cboGenre1.NewIndex) = 110
        
        .AddItem "Satire"
        .ItemData(cboGenre1.NewIndex) = 111
        
        .AddItem "Slow Jam"
        .ItemData(cboGenre1.NewIndex) = 112
        
        .AddItem "Club"
        .ItemData(cboGenre1.NewIndex) = 113
        
        .AddItem "Tango"
        .ItemData(cboGenre1.NewIndex) = 114
        
        .AddItem "Samba"
        .ItemData(cboGenre1.NewIndex) = 115
        
        .AddItem "Folklore"
        .ItemData(cboGenre1.NewIndex) = 116
        
        .AddItem "Ballad"
        .ItemData(cboGenre1.NewIndex) = 117
        
        .AddItem "Power Ballad"
        .ItemData(cboGenre1.NewIndex) = 118
        
        .AddItem "Rhythmic Soul"
        .ItemData(cboGenre1.NewIndex) = 119
        
        .AddItem "Freestyle"
        .ItemData(cboGenre1.NewIndex) = 120
    
        .AddItem "Duet"
        .ItemData(cboGenre1.NewIndex) = 121
        
        .AddItem "Punk Rock"
        .ItemData(cboGenre1.NewIndex) = 122
        
        .AddItem "Drum Solo"
        .ItemData(cboGenre1.NewIndex) = 123
        
        .AddItem "A Capella"
        .ItemData(cboGenre1.NewIndex) = 124
        
        .AddItem "Eurohouse"
        .ItemData(cboGenre1.NewIndex) = 125
        
        .AddItem "Dance Hall"
        .ItemData(cboGenre1.NewIndex) = 126
            
    End With
    
'Populate the ID3v2 genre combo box
    
    With cboGenre2
    
        .AddItem Chr(0)
        .ItemData(cboGenre2.NewIndex) = 0
        
        .AddItem "Blues"
        .ItemData(cboGenre2.NewIndex) = 1
        
        .AddItem "Classic Rock"
        .ItemData(cboGenre2.NewIndex) = 2
        
        .AddItem "Country"
        .ItemData(cboGenre2.NewIndex) = 3
        
        .AddItem "Dance"
        .ItemData(cboGenre2.NewIndex) = 4
        
        .AddItem "Disco"
        .ItemData(cboGenre2.NewIndex) = 5
        
        .AddItem "Funk"
        .ItemData(cboGenre2.NewIndex) = 6
        
        .AddItem "Grunge"
        .ItemData(cboGenre2.NewIndex) = 7
        
        .AddItem "Hip-Hop"
        .ItemData(cboGenre2.NewIndex) = 8
        
        .AddItem "Jazz"
        .ItemData(cboGenre2.NewIndex) = 9
        
        .AddItem "Metal"
        .ItemData(cboGenre2.NewIndex) = 10
        
        .AddItem "New Age"
        .ItemData(cboGenre2.NewIndex) = 11
        
        .AddItem "Oldies"
        .ItemData(cboGenre2.NewIndex) = 12
        
        .AddItem "Other"
        .ItemData(cboGenre2.NewIndex) = 13
        
        .AddItem "Pop"
        .ItemData(cboGenre2.NewIndex) = 14
        
        .AddItem "R&B"
        .ItemData(cboGenre2.NewIndex) = 15
        
        .AddItem "Rap"
        .ItemData(cboGenre2.NewIndex) = 16
        
        .AddItem "Reggae"
        .ItemData(cboGenre2.NewIndex) = 17
        
        .AddItem "Rock"
        .ItemData(cboGenre2.NewIndex) = 18
        
        .AddItem "Techno"
        .ItemData(cboGenre2.NewIndex) = 19
        
        .AddItem "Industrial"
        .ItemData(cboGenre2.NewIndex) = 20
        
        .AddItem "Alternative"
        .ItemData(cboGenre2.NewIndex) = 21
        
        .AddItem "Ska"
        .ItemData(cboGenre2.NewIndex) = 22
        
        .AddItem "Death Metal"
        .ItemData(cboGenre2.NewIndex) = 23
        
        .AddItem "Pranks"
        .ItemData(cboGenre2.NewIndex) = 24
        
        .AddItem "Soundtrack"
        .ItemData(cboGenre2.NewIndex) = 25
        
        .AddItem "Euro-Techno"
        .ItemData(cboGenre2.NewIndex) = 26
        
        .AddItem "Ambient"
        .ItemData(cboGenre2.NewIndex) = 27
        
        .AddItem "Trip-Hop"
        .ItemData(cboGenre2.NewIndex) = 28
        
        .AddItem "Vocal"
        .ItemData(cboGenre2.NewIndex) = 29
        
        .AddItem "Jazz+Funk"
        .ItemData(cboGenre2.NewIndex) = 30
        
        .AddItem "Fusion"
        .ItemData(cboGenre2.NewIndex) = 31
        
        .AddItem "Trance"
        .ItemData(cboGenre2.NewIndex) = 32
        
        .AddItem "Classical"
        .ItemData(cboGenre2.NewIndex) = 33
        
        .AddItem "Instrumental"
        .ItemData(cboGenre2.NewIndex) = 34
        
        .AddItem "Acid"
        .ItemData(cboGenre2.NewIndex) = 35
        
        .AddItem "House"
        .ItemData(cboGenre2.NewIndex) = 36
        
        .AddItem "Game"
        .ItemData(cboGenre2.NewIndex) = 37
        
        .AddItem "Sound Clip"
        .ItemData(cboGenre2.NewIndex) = 38
        
        .AddItem "Gospel"
        .ItemData(cboGenre2.NewIndex) = 39
        
        .AddItem "Noise"
        .ItemData(cboGenre2.NewIndex) = 40
        
        .AddItem "AlternRock"
        .ItemData(cboGenre2.NewIndex) = 41
        
        .AddItem "Bass"
        .ItemData(cboGenre2.NewIndex) = 42
        
        .AddItem "Soul"
        .ItemData(cboGenre2.NewIndex) = 43
        
        .AddItem "Punk"
        .ItemData(cboGenre2.NewIndex) = 44
        
        .AddItem "Space"
        .ItemData(cboGenre2.NewIndex) = 45
        
        .AddItem "Meditative"
        .ItemData(cboGenre2.NewIndex) = 46
        
        .AddItem "Instrumental Pop"
        .ItemData(cboGenre2.NewIndex) = 47
        
        .AddItem "Instrumental Rock"
        .ItemData(cboGenre2.NewIndex) = 48
        
        .AddItem "Ethnic"
        .ItemData(cboGenre2.NewIndex) = 49
        
        .AddItem "Gothic"
        .ItemData(cboGenre2.NewIndex) = 50
        
        .AddItem "Darkwave"
        .ItemData(cboGenre2.NewIndex) = 51
        
        .AddItem "Techno-Industrial"
        .ItemData(cboGenre2.NewIndex) = 52
        
        .AddItem "Electronic"
        .ItemData(cboGenre2.NewIndex) = 53
        
        .AddItem "Pop-Folk"
        .ItemData(cboGenre2.NewIndex) = 54
        
        .AddItem "Eurodance"
        .ItemData(cboGenre2.NewIndex) = 55
        
        .AddItem "Dream"
        .ItemData(cboGenre2.NewIndex) = 56
        
        .AddItem "Southern Rock"
        .ItemData(cboGenre2.NewIndex) = 57
        
        .AddItem "Comedy"
        .ItemData(cboGenre2.NewIndex) = 58
        
        .AddItem "Cult"
        .ItemData(cboGenre2.NewIndex) = 59
        
        .AddItem "Gangsta"
        .ItemData(cboGenre2.NewIndex) = 60
        
        .AddItem "Top 40"
        .ItemData(cboGenre2.NewIndex) = 61
        
        .AddItem "Christian Rap"
        .ItemData(cboGenre2.NewIndex) = 62
        
        .AddItem "Pop/Funk"
        .ItemData(cboGenre2.NewIndex) = 63
        
        .AddItem "Jungle"
        .ItemData(cboGenre2.NewIndex) = 64
        
        .AddItem "Native American"
        .ItemData(cboGenre2.NewIndex) = 65
        
        .AddItem "Cabaret"
        .ItemData(cboGenre2.NewIndex) = 66
        
        .AddItem "New Wave"
        .ItemData(cboGenre2.NewIndex) = 67
        
        .AddItem "Psychadelic"
        .ItemData(cboGenre2.NewIndex) = 68
        
        .AddItem "Rave"
        .ItemData(cboGenre2.NewIndex) = 69
        
        .AddItem "Showtunes"
        .ItemData(cboGenre2.NewIndex) = 70
        
        .AddItem "Trailer"
        .ItemData(cboGenre2.NewIndex) = 71
        
        .AddItem "Lo-Fi"
        .ItemData(cboGenre2.NewIndex) = 72
        
        .AddItem "Tribal"
        .ItemData(cboGenre2.NewIndex) = 73
        
        .AddItem "Acid Punk"
        .ItemData(cboGenre2.NewIndex) = 74
        
        .AddItem "Acid Jazz"
        .ItemData(cboGenre2.NewIndex) = 75
        
        .AddItem "Polka"
        .ItemData(cboGenre2.NewIndex) = 76
        
        .AddItem "Retro"
        .ItemData(cboGenre2.NewIndex) = 77
        
        .AddItem "Musical"
        .ItemData(cboGenre2.NewIndex) = 78
        
        .AddItem "Rock & Roll"
        .ItemData(cboGenre2.NewIndex) = 79
        
        .AddItem "Hard Rock"
        .ItemData(cboGenre2.NewIndex) = 80
        
        .AddItem "Folk"
        .ItemData(cboGenre2.NewIndex) = 81
        
        .AddItem "Folk-Rock"
        .ItemData(cboGenre2.NewIndex) = 82
        
        .AddItem "National Folk"
        .ItemData(cboGenre2.NewIndex) = 83
        
        .AddItem "Swing"
        .ItemData(cboGenre2.NewIndex) = 84
        
        .AddItem "Fast Fusion"
        .ItemData(cboGenre2.NewIndex) = 85
        
        .AddItem "Bebop"
        .ItemData(cboGenre2.NewIndex) = 86
        
        .AddItem "Latin"
        .ItemData(cboGenre2.NewIndex) = 87
        
        .AddItem "Revival"
        .ItemData(cboGenre2.NewIndex) = 88
        
        .AddItem "Celtic"
        .ItemData(cboGenre2.NewIndex) = 89
        
        .AddItem "Bluegrass"
        .ItemData(cboGenre2.NewIndex) = 90
        
        .AddItem "Avantgarde"
        .ItemData(cboGenre2.NewIndex) = 91
        
        .AddItem "Gothic Rock"
        .ItemData(cboGenre2.NewIndex) = 92
        
        .AddItem "Progressive Rock"
        .ItemData(cboGenre2.NewIndex) = 93
        
        .AddItem "Psychedlic Rock"
        .ItemData(cboGenre2.NewIndex) = 94
        
        .AddItem "Symphonic Rock"
        .ItemData(cboGenre2.NewIndex) = 95
        
        .AddItem "Slow Rock"
        .ItemData(cboGenre2.NewIndex) = 96
    
        .AddItem "Big Band"
        .ItemData(cboGenre2.NewIndex) = 97
        
        .AddItem "Chorus"
        .ItemData(cboGenre2.NewIndex) = 98
        
        .AddItem "Easy Listening"
        .ItemData(cboGenre2.NewIndex) = 99
        
        .AddItem "Acoustic"
        .ItemData(cboGenre2.NewIndex) = 100
        
        .AddItem "Humour"
        .ItemData(cboGenre2.NewIndex) = 101
        
        .AddItem "Speech"
        .ItemData(cboGenre2.NewIndex) = 102
        
        .AddItem "Chanson"
        .ItemData(cboGenre2.NewIndex) = 103
        
        .AddItem "Opera"
        .ItemData(cboGenre2.NewIndex) = 104
        
        .AddItem "Chamber Music"
        .ItemData(cboGenre2.NewIndex) = 105
        
        .AddItem "Sonota"
        .ItemData(cboGenre2.NewIndex) = 106
        
        .AddItem "Symphony"
        .ItemData(cboGenre2.NewIndex) = 107
        
        .AddItem "Booty Bass"
        .ItemData(cboGenre2.NewIndex) = 108
        
        .AddItem "Primus"
        .ItemData(cboGenre2.NewIndex) = 109
        
        .AddItem "Porn Groove"
        .ItemData(cboGenre2.NewIndex) = 110
        
        .AddItem "Satire"
        .ItemData(cboGenre2.NewIndex) = 111
        
        .AddItem "Slow Jam"
        .ItemData(cboGenre2.NewIndex) = 112
        
        .AddItem "Club"
        .ItemData(cboGenre2.NewIndex) = 113
        
        .AddItem "Tango"
        .ItemData(cboGenre2.NewIndex) = 114
        
        .AddItem "Samba"
        .ItemData(cboGenre2.NewIndex) = 115
        
        .AddItem "Folklore"
        .ItemData(cboGenre2.NewIndex) = 116
        
        .AddItem "Ballad"
        .ItemData(cboGenre2.NewIndex) = 117
        
        .AddItem "Power Ballad"
        .ItemData(cboGenre2.NewIndex) = 118
        
        .AddItem "Rhythmic Soul"
        .ItemData(cboGenre2.NewIndex) = 119
        
        .AddItem "Freestyle"
        .ItemData(cboGenre2.NewIndex) = 120
    
        .AddItem "Duet"
        .ItemData(cboGenre2.NewIndex) = 121
        
        .AddItem "Punk Rock"
        .ItemData(cboGenre2.NewIndex) = 122
        
        .AddItem "Drum Solo"
        .ItemData(cboGenre2.NewIndex) = 123
        
        .AddItem "A Capella"
        .ItemData(cboGenre2.NewIndex) = 124
        
        .AddItem "Eurohouse"
        .ItemData(cboGenre2.NewIndex) = 125
        
        .AddItem "Dance Hall"
        .ItemData(cboGenre2.NewIndex) = 126
            
    End With


End Sub

