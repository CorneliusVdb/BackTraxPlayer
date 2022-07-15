Attribute VB_Name = "Module1"

Public i As Integer
Public strEmptyString As String
Public B As Byte
Public s As String

Public Type ID3v1Tag
  id As String * 3
  title As String * 30
  Artist As String * 30
  Album As String * 30
  Year As String * 4
  Comment As String * 30
  Genre As Byte
End Type
Public Tag1 As ID3v1Tag

Public Type ID3v2Tag
  id As String
  title As String
  Artist As String
  Album As String
  Year As String
  Comment As String
  Genre As Integer
  TrackNr As String
End Type
Public TagV2 As ID3v2Tag

Private Version As Byte

Public Function ReadFile(strFilename As String)

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

On Error GoTo errorhandler

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' use the filename to get ID3 info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Dim strFilename As String
    Dim lngFilesize As Long
    
    'strFilename = Me.Caption

    Dim fn As Integer
    Dim lngHeaderPosition As Long
    
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
    bId3V1Found = False
    If Tag1.id = "TAG" Then 'If "TAG" is present, then we have a valid ID3v1 tag and will extract all available ID3v1 info from the file
        Get #fn, , Tag1.title   'Always limited to 30 characters
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
        If Trim(Tag1.Artist) = "" Then
          If Trim(Tag1.title) = "" Then
            bId3V1Found = False
          Else
            sListName = Trim(Tag1.title)
            sTitle = Tag1.title
          End If
        Else
          sListName = Trim(Tag1.Artist) & " - " & Trim(Tag1.title)
          sArtist = Tag1.Artist
          sTitle = Tag1.title
          bTitleArtistLoaded = True
        End If

      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Proceed to extract the ID3v2 tag info if any exists
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If Not bId3V1Found Then
        bId3V2Found = False
        If Tag2 <> strEmptyString Then
  ''''''''            frmID3.chkv2.value = 1
          GetID3v2Tag1 (Tag2) 'Pass the Id3v2 TagId to the GetID3v2Tag1 function
          
         'My own code
          bId3V2Found = True
          With TagV2
            If Trim(.Artist) = "" Then
              If Trim(.title) = "" Then
                bId3V2Found = False
              Else
                sListName = Trim(.title)
                sTitle = .title
              End If
            Else
              sListName = Trim(.Artist) & " - " & Trim(.title)
              sArtist = Trim(.Artist)
              sTitle = Trim(.title)
            End If
          End With
        End If
      End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Close the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Close
      
      'No tags found, use file name for display...
      If sArtist = "" And sTitle = "" Then
        For i = Len(sFileToLoad) To 1 Step -1
          If Mid(sFileToLoad, i, 1) = "\" Then
            sListName = Mid(sFileToLoad, i + 1)
            iPos = InStr(1, sListName, ".")
            sListName = Mid(sListName, 1, iPos - 1)
            sArtist = ""
            sTitle = sListName
            Exit For
          End If
        Next i
      End If
      
      Id3TagArr(0) = sListName
      Id3TagArr(1) = sTitle
      Id3TagArr(2) = sArtist

    Exit Function
        
errorhandler:
    'MsgBox "Error reading file"
    Err.Clear
    Close
    Resume Next
End Function

Public Function GetID3v2Tag1(Tag2 As String) As Boolean

On Error GoTo errorhandler

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
          TagV2.title = Mid$(Tag2, i + FieldOffset, FieldSize)
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
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadYear
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
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadGenre
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
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadTrackNbr
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
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo Done
      End If
      TagV2.TrackNr = Mid$(Tag2, i + FieldOffset, FieldSize)
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We're done looking for ID3v2 info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Done:
   
   Exit Function

errorhandler:
   Err.Clear
   Resume Next
End Function

