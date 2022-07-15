Attribute VB_Name = "modMp3TagExtractor"
''' Copy all of the files in the dir, exept the frmPlayer. Use old one, and just replace code below...

'''   Sub Setstate(Index As Integer)

''' Case 2  'Assign                                                                       '
'''    '======================================================================================
'''
'''      frmExplorer.Show vbModal
'''      Screen.MousePointer = vbHourglass
'''      Me.Enabled = False
'''      DoEvents
'''      If FilenameToLoad = "" Then
'''        Screen.MousePointer = vbDefault
'''        Me.Enabled = True
'''        Exit Sub
'''      End If
'''
'''      'Reset button to Loaded state
'''      ResetButton Index
'''      'Add caption to button
'''
'''      'tags.MP3File = FilenameToLoad
'''      'GetId3Tags FilenameToLoad
'''      ExtractTagInfo FilenameToLoad
      
'''    '======================================================================================
'''    Case 4  'Loadfrom drag and drop                                                       '
'''    '======================================================================================
'''      Screen.MousePointer = vbHourglass
'''      Me.Enabled = False
'''      DoEvents
'''      If FilenameToLoad = "" Then
'''        Screen.MousePointer = vbDefault
'''        Me.Enabled = True
'''        Exit Sub
'''      End If
'''
'''      'Reset button to Loaded state
'''      ResetButton Index
'''      'Add caption to button
'''      ExtractTagInfo FilenameToLoad
'''     ' GetId3Tags FilenameToLoad
'''
'''      SetupButton Index, Id3TagArr(1), Id3TagArr(2), FilenameToLoad
      
      
      



Public i As Integer
Public strEmptyString As String
Public b As Byte
Public s As String

Public Type ID3v1Tag
  id As String * 3
  Title As String * 30
  Artist As String * 30
  Album As String * 30
  Year As String * 4
  Comment As String * 30
  Genre As Byte
End Type
Public Tag1 As ID3v1Tag

Public Type ID3v2Tag
  id As String
  Title As String
  Artist As String
  Album As String
  Year As String
  Comment As String
  Genre As Integer
  TrackNr As String
End Type
Public TagV2 As ID3v2Tag

Private Version As Byte

Public Function ExtractTagInfo(strFilename As String)

Dim bTitleArtistLoaded As Boolean
Dim sArtist As String
Dim sTitle As String
Dim sListName As String


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

    'Dim strFilename As String
    Dim lngFilesize As Long
    
    'strFilename = Me.Caption

    Dim fn As Integer
    Dim lngHeaderPosition As Long
    Dim Tag2 As String
    'Clear tags
    Dim tEmtpy As ID3v1Tag
    Tag1 = tEmtpy
    Dim tEmtpy2 As ID3v2Tag
    TagV2 = tEmtpy2
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    fn = FreeFile
    
    Open strFilename For Binary As #fn                      'Open the file so we can read it
    lngFilesize = LOF(fn)                                   'Size of the file, in bytes

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check for a Header
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
    Get #fn, 1, b
        
    If b <> 255 Then '(255 is where an ID3v2 header should start)
        If b <> 73 Then
            'Exit Function
        End If
    End If
     
    lngHeaderPosition = 1
    Get #fn, 2, b
    If (b < 250 Or b > 251) Then
        'We have an ID3v2 tag
        If b = 68 Then
            Get #fn, 3, b
            If b = 51 Then
                Dim R As Double
                Get #fn, 4, Version
                Get #fn, 7, b
                R = b * 20917152
                Get #fn, 8, b
                R = R + (b * 16384)
                Get #fn, 9, b
                R = R + (b * 128)
                Get #fn, 10, b
                R = R + b
                If R > lngFilesize Or R > 2147483647 Then
                    GoTo CheckIDV1
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

CheckIDV1:

  'ID3v1 tag
    Get #fn, lngFilesize - 127, Tag1.id
    bId3V1Found = False
    If Trim(Tag1.id) = "TAG" Then 'If "TAG" is present, then we have a valid ID3v1 tag and will extract all available ID3v1 info from the file
        Get #fn, , Tag1.Title   'Always limited to 30 characters
        Get #fn, , Tag1.Artist  'Always limited to 30 characters
        Get #fn, , Tag1.Album   'Always limited to 30 characters
        Get #fn, , Tag1.Year    'Always limited to 4 characters
        Get #fn, , Tag1.Comment 'Always limited to 30 characters
        Get #fn, , Tag1.Genre   'Always limited to 1 byte (?)
            
'''''''''            frmID3.chkv1.value = 1 'Indicates that the file contains ID3v1 tag info
'''''''''
'''''''''            'Populate the form with the ID3v1 info
'''''''''            With frmID3
'''''''''                txtTrack1.Text = Trim$(Tag1.title)
'''''''''                txtArtist1.Text = Trim$(Tag1.Artist)
'''''''''                txtAlbum1.Text = Trim$(Tag1.Album)
'''''''''                txtYear1.Text = Trim$(Tag1.Year)
'''''''''                txtComments1.Text = Trim$(Tag1.Comment)
'''''''''                txtGenre1.Text = Tag1.Genre
'''''''''            End With
'''''''''
'''''''''            cboGenre1.ListIndex = Tag1.Genre + 1


        bId3V1Found = True
        Tag1.Artist = Trim(Replace(Tag1.Artist, Chr(0), ""))
        Tag1.Title = Trim(Replace(Tag1.Title, Chr(0), ""))
        
        If Trim(Tag1.Artist) = "" Then
          If Trim(Tag1.Title) = "" Then
            bId3V1Found = False
          Else
            sListName = Trim(Tag1.Title)
            sTitle = Trim(Tag1.Title)
          End If
        Else
          sListName = Trim(Tag1.Artist) & " - " & Trim(Tag1.Title)
          sArtist = Trim(Tag1.Artist)
          sTitle = Trim(Tag1.Title)
          bTitleArtistLoaded = True
        End If

      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Proceed to extract the ID3v2 tag info if any exists
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''      If sArtist = "" Or sTitle = "" Or Not bId3V1Found Then
'''        bId3V2Found = False
'''        If Trim(Tag2) <> strEmptyString Then
'''  ''''''''            frmID3.chkv2.value = 1
'''          GetID3v2Tag1 (Tag2) 'Pass the Id3v2 TagId to the GetID3v2Tag1 function
'''
'''         'My own code
'''          bId3V2Found = True
'''          With TagV2
'''            .Artist = Trim(Replace(.Artist, Chr(0), ""))
'''            .Title = Trim(Replace(.Title, Chr(0), ""))
'''
'''            If Trim(.Artist) = "" Then
'''              If Trim(.Title) = "" Then
'''                bId3V2Found = False
'''              Else
'''                sListName = Trim(.Title)
'''                sTitle = Trim(.Title)
'''              End If
'''            Else
'''              sListName = Trim(.Artist) & " - " & Trim(.Title)
'''              sArtist = Trim(.Artist)
'''              sTitle = Trim(.Title)
'''            End If
'''          End With
'''        End If
'''      End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Close the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Close
      
      'No tags found, use file name for display...
      If sArtist = "" And sTitle = "" Then
        For i = Len(strFilename) To 1 Step -1
          If Mid(strFilename, i, 1) = "\" Then
            sListName = Mid(strFilename, i + 1)
            sListName = Left(sListName, Len(sListName) - 4) 'Remove the extention from the listname...
            sListName = Replace(sListName, "www.livingelectro.com", "")
            sArtist = ""
            sTitle = sListName
            Exit For
          End If
        Next i
      End If
      
      Id3TagArr(0) = Replace(Replace(sListName, Chr(0), ""), "www.livingelectro.com", "")
      Id3TagArr(1) = Replace(Replace(sTitle, Chr(0), ""), "www.livingelectro.com", "")
      Id3TagArr(2) = Replace(Replace(sArtist, Chr(0), ""), "www.livingelectro.com", "")

    Exit Function
        
errorHandler:
    'MsgBox "Error reading file"
    Err.Clear
    Close
    Resume Next
End Function

Public Function GetID3v2Tag1(Tag2 As String) As Boolean

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
             b = Asc(Mid$(Tag2, i + 9))
             If (b And 128) = True Or (b And 64) = True Then GoTo ReadAlbum
          End If
          TagV2.Title = Mid$(Tag2, i + FieldOffset, FieldSize)
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
          b = Asc(Mid$(Tag2, i + 9))
          If (b And 128) = 128 Or (b And 64) = 64 Then GoTo ReadArtist
       End If
       TagV2.Album = Mid$(Tag2, i + FieldOffset, FieldSize)
       
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
         b = Asc(Mid$(Tag2, i + 9))
         If (b And 128) = 128 Or (b And 64) = 64 Then GoTo ReadYear
      End If
      TagV2.Artist = Mid$(Tag2, i + FieldOffset, FieldSize)
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
         b = Asc(Mid$(Tag2, i + 9))
         If (b And 128) = 128 Or (b And 64) = 64 Then GoTo ReadGenre
      End If
      TagV2.Year = Mid$(Tag2, i + FieldOffset, FieldSize)
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
         b = Asc(Mid$(Tag2, i + 9))
         If (b And 128) = 128 Or (b And 64) = 64 Then GoTo ReadTrackNbr
      End If
      
      s = Mid$(Tag2, i + FieldOffset, FieldSize)
      If Left$(s, 1) = "(" Then
        TagV2.Genre = Val(Mid$(s, 2, 2))
'        cboGenre2.ListIndex = Val(txtGenre.Text) + 1
        
      Else
         'i = InStr(gsGenres, s & Space$(22 - Len(s)))
         TagV2.Genre = i
         
'         cboGenre2.ListIndex = i
         If i > 0 Then
            TagV2.Genre = Int(i / 22)
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
         b = Asc(Mid$(Tag2, i + 9))
         If (b And 128) = 128 Or (b And 64) = 64 Then GoTo Done
      End If
      TagV2.TrackNr = Mid$(Tag2, i + FieldOffset, FieldSize)
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

