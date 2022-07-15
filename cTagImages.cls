VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTaggedImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class attempts to extract image formats from an MP3 file

' Using 3 separate parsing routines, one for each version of the id3 tags; though versions are similar
' id3v2.x uses 3 byte length values & 3 character frame types & no option for extended headers
' id3v3.x uses 4 byte length values & 4 character frame types. Extended header size is not sync-safe
' id3v4.x similar to id3v3 except Extended header size is sync-safe & additional frame flag values
'   id3v4 offers 4 types of text encoding vs just the 2 types that id3v3 offered
'   id3v4 allows individual frames to be unsynchronized vs. just the entire tag

' Parsing routines will undo unsynchronization for all 3 versions. See pvSynchronizeTag for more
' If any compression is used (all versions) or encryption (id3v4) then frames are skipped
' Though I can include decompression routines, I haven't found any mp3s that use zLib compression

' unicode supported file-related APIs
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1&

' unicode supported string-related APIs
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long

Private Type MP3Offsets     ' used to track image data found within the passed tag
    Start As Long           ' where image starts in the tag
    Length As Long          ' number of bytes for the image data
    Category As Long        ' what category the image belongs to
    MIME As String          ' image type
End Type
Private m_Stream() As Byte  ' cached tag
Private m_Frames() As MP3Offsets
Private m_Count As Long
Private m_TagLen As Long

Public Function LoadTagged_Images(ByVal FileName As String) As Long

    ' Function passes off parsing to individual routines
    ' Return value is number of image(s) found
    
    Dim fNr As Long, lSize As Long, lRead As Long
    Dim sGUID As String, GUID(0 To 3) As Long
    Const tagWMAheader As String = "{75B22630-668E-11CF-A6D9-00AA0062CE6C}"
    
    m_Count = 0&
    m_TagLen = 0&
    Erase m_Stream()
    Erase m_Frames()
    
    fNr = pvGetFileHandle(FileName)
    If fNr = INVALID_HANDLE_VALUE Then Exit Function
    
    lSize = GetFileSize(fNr, 0&)
    If lSize < 30& Then
        CloseHandle fNr
        Exit Function
    End If
    
    ReDim m_Stream(0 To 29)
    SetFilePointer fNr, 0&, 0&, 0&
    ReadFile fNr, m_Stream(0), 30&, lRead, ByVal 0&
    If lRead = 30& Then
    
        If (m_Stream(0) = 73 And m_Stream(1) = 68 And m_Stream(2) = 51) Then ' valid tags start with ID3
            If m_Stream(3) > 1 And m_Stream(3) < 5 Then         ' supports versions 2,3,4
                If m_Stream(4) < 255 Then                       ' revision will never be 255
                    m_TagLen = pvSyncSafeLong(m_Stream, 6&) ' length of tag
                    If m_TagLen > lSize Then  ' can't be larger than entire file
                        m_TagLen = 0&                             ' failure flag
                    Else
                        ReDim Preserve m_Stream(0 To m_TagLen - 1&)
                        If m_TagLen > 30& Then
                            ReadFile fNr, m_Stream(30), m_TagLen - 30&, lRead, ByVal 0&
                            If lRead < m_TagLen - 30& Then m_TagLen = 0& ' couldn't access the data?
                        End If
                    End If
                End If
            End If
        Else                                        ' check for wma vs mp3
            sGUID = String$(38, vbNullChar)
            If StringFromGUID2(VarPtr(m_Stream(0)), StrPtr(sGUID), 39) > 0& Then
                If sGUID = tagWMAheader Then
                    CopyMemory m_TagLen, m_Stream(20), 4&
                    If m_TagLen = 0& Then ' else the file is either very huge or malformatted wma
                        CopyMemory m_TagLen, m_Stream(16), 4&
                        If m_TagLen > 60& And m_TagLen < lSize Then
                            If m_Stream(29) = 2 Then
                                ' This field must be set to the value 0x02.
                                ' If the value is different when read, the application should fail to source the content.
                                ReDim Preserve m_Stream(0 To m_TagLen - 1&)
                                If m_TagLen > 30& Then
                                    ReadFile fNr, m_Stream(30), m_TagLen - 30&, lRead, ByVal 0&
                                    If lRead < m_TagLen - 30& Then m_TagLen = 0& ' couldn't access the data?
                                End If
                            Else
                                m_TagLen = 0&
                            End If
                        Else
                            m_TagLen = 0&
                        End If
                    End If
                End If
            End If
        End If
    End If
    CloseHandle fNr

    If m_TagLen = 0& Then
        Erase m_Stream()
    Else
        Select Case m_Stream(3) ' version
            Case 2: Call pvParse3v2
            Case 3: Call pvParse3v3
            Case 4: Call pvParse3v4
            Case Else
                If sGUID = tagWMAheader Then pvParseWMA
        End Select
    End If
    If m_Count = 0& Then Erase m_Frames()  ' failed to parse any images
    
    LoadTagged_Images = m_Count

End Function

Public Property Get ImageCount() As Long
    ImageCount = m_Count
End Property

Public Property Get ImageType(Index As Long) As String
    ' 1 based index
    If Index > 0 And Index <= m_Count Then
        ImageType = m_Frames(Index - 1).MIME
    End If
End Property

Public Property Get ImageCategory(Index As Long) As Long
    ' 1 based index
    If Index > 0 And Index <= m_Count Then
        ImageCategory = m_Frames(Index - 1).Category
    End If
' Categories:
'0   Other
'1  32x32 pixels 'file icon' (PNG only)
'2  Other file icon
'3   Cover (front)
'4   Cover (back)
'5   Leaflet Page
'6  Media (e.g. label side of CD)
'7  Lead artist/lead performer/soloist
'8  Artist/performer
'9   Conductor
'10  Band/Orchestra
'11  Composer
'12  Lyricist/text writer
'13  recording Location
'14  During recording
'15  During performance
'16  Movie/video screen capture
'17  A bright coloured fish
'18  Illustration
'19  Band/artist logotype
'20  Publisher/Studio logotype

End Property

Public Function ExtractImageData(ByVal Index As Long, outArray() As Byte) As Boolean
    ' 1 based index
    ' These bytes can be PNG, GIF, JPG, BMP, or possibly anything really
    ' The ImgeType property may let you know what type it is; assuming the info is correct
    ' Suggest sending the bytes to GDI+ API: GdipLoadImageFromStream
    '       that API does not need to know image format & can load most common image formats
    If Index > 0 And Index <= m_Count Then
        With m_Frames(Index - 1)
            ReDim outArray(0 To .Length - 1)
            CopyMemory outArray(0), m_Stream(.Start), .Length
        End With
        ExtractImageData = True
    End If

End Function

Public Function ExtractTag(outArray As Byte) As Long

    ' returns the complete ID3 tag. Return value is length of tag
    If m_TagLen Then
        outArray = m_Stream()
        ExtractTag = m_TagLen
    End If
    
End Function

Private Function pvGetFileHandle(ByVal FileName As String) As Long

    ' Function uses APIs to create a file handle to read/write files with unicode support

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&
    
    Dim Flags As Long, Access As Long
    Dim Disposition As Long, Share As Long
    
    Access = GENERIC_READ
    Share = FILE_SHARE_READ
    Disposition = OPEN_EXISTING
    Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
            Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    
    If IsWindowUnicode(GetDesktopWindow) Then
        pvGetFileHandle = CreateFileW(StrPtr(FileName), Access, Share, ByVal 0&, Disposition, Flags, 0&)
    Else
        pvGetFileHandle = CreateFile(FileName, Access, Share, ByVal 0&, Disposition, Flags, 0&)
    End If
    
End Function


Private Function pvReverseLong4(inStream() As Byte, lOffset As Long) As Long

    ' Convert 4 byte big endian to little endian
    pvReverseLong4 = inStream(lOffset + 3&) Or inStream(lOffset + 2&) * &H100& _
        Or inStream(lOffset + 1&) * &H10000 Or (inStream(lOffset) And &H7F) * &H1000000
    If (inStream(lOffset) And &H80) Then pvReverseLong4 = pvReverseLong4 Or &H80000000

End Function

Private Function pvReverseLong3(inStream() As Byte, lOffset As Long) As Long

    ' Convert 3 byte big endian to little endian
    pvReverseLong3 = inStream(lOffset + 2&) Or inStream(lOffset + 1&) * &H100& Or inStream(lOffset) * &H10000

End Function

Private Function pvSyncSafeLong(inStream() As Byte, lOffset As Long) As Long

    ' Convert 4 byte big endian sync-safe to 4 byte little endian
    ' Sync-safe values are 28 bits vs 32 bits. The high bit of each byte is zero/ignored
    ' -- concatenate the 1st 7 bits of each byte to create a long value
    pvSyncSafeLong = (inStream(lOffset + 3&) And &H7F) Or (inStream(lOffset + 2&) And &H7F) * &H80& _
        Or (inStream(lOffset + 1&) And &H7F) * &H4000& Or (inStream(lOffset) And &H7F) * &H200000

End Function

Private Function pvStringFromArray(inStream() As Byte, Offset As Long, nrChars As Long) As String

    Dim sText As String
    Dim c As Long
    For c = Offset To Offset + nrChars - 1
        If inStream(c) < 32 Then Exit For
        sText = sText & Chr$(inStream(c))
    Next
    pvStringFromArray = sText

End Function

Private Function pvSynchronizeTag(inStream() As Byte, Offset As Long, nrBytes As Long) As Long

    Dim X As Long, aPtr As Long, zPtr As Long
    Dim lResult As Long, aPtrAnchor As Long

    ' if unsync'd, then need to fix it up before parsing
    ' unsynchronizing a tag was done for much older mp3 players, that didn't recognize ID3 tags
    
    ' 1. the unsync encoding (done by whoever created the id3 tag) looks for any 11 consecutive bits that are all on
    ' 2. then it splits the 1st 8 bits from the next 3 bits by inserting a null byte
    ' 3. but the splitting affects any naturally occurring 255 0 byte combinations, so it first must fix those
    '   by inserting a null byte between them: so 11111111 00000000 will end up as 11111111 00000000 00000000
    ' 4. Now for sync'ing:
    '   these 2 bytes: 11111111 111xxxxx will end up as 3 bytes: 11111111 00000000 111xxxxx
    ' 5. to remove the scheme, we need to simply remove any null byte that follows a 255 byte
    
    ' id3v4 has ability to unsync individual frames, less their 10 byte frame header
    ' earlier versions applied unsync to entire tag, less the 6/10 byte tag header
    '   id3v4 can also apply this to the entire tag, same logic
    
    ' Note: Any array passed to this function is owned by this control. Am not messing with user data
    
    zPtr = Offset: aPtrAnchor = Offset
    For aPtr = Offset + 1& To Offset + nrBytes - 1&
        If inStream(aPtr) = 0& Then                 ' see if previous byte was 255
            If inStream(aPtr - 1&) = 255 Then       ' shift array left
                If lResult Then CopyMemory inStream(zPtr), inStream(aPtrAnchor), aPtr - aPtrAnchor
                zPtr = zPtr + aPtr - aPtrAnchor     ' adjust position for next write
                aPtrAnchor = aPtr + 1&              ' adjust position for first inclusive byte
                lResult = lResult + 1&              ' keep track of number of bytes removed
            End If
        End If
    Next
    If (aPtrAnchor < aPtr) And (lResult > 0&) Then  ' perform last shift if needed
        CopyMemory inStream(zPtr), inStream(aPtrAnchor), aPtr - aPtrAnchor
        inStream(Offset + nrBytes - lResult - 1&) = 0 ' ensure adjusted tag terminates with zero
    End If
    pvSynchronizeTag = lResult

End Function

Private Sub pvParse3v2()

    ' http://www.id3.org/id3v2-00
    
    ' id3v2.3.0 header block format
    ' offset 0 : 3 byte magic number: ID3
    ' offset 3 : 1 byte major version
    ' offset 4 : 1 byte revision
    ' offset 5 : 1 byte flag   (bit 7=unsync'd, 6=compression)
    ' offset 6 : 4 byte sync-safe length

    Dim lMax As Long, aPtr As Long, lSize As Long
    Dim lValue As Long, imgOffset As Long, sType As String
    Const tagPIC As Long = 4409680                      ' only care about these frames
    
    ' frame header (6 bytes always)
    ' aPtr + 0 : 3 byte header type (i.e., PIC)
    ' aPtr + 3 : 3 byte header size as big endian long
    If (m_Stream(5) And &H40) Then Exit Sub             ' standard compression was never implemented
    
    aPtr = 10&
    lMax = pvSyncSafeLong(m_Stream(), 6&)
    If (m_Stream(5) And &H80) Then lMax = lMax - pvSynchronizeTag(m_Stream(), aPtr, lMax - aPtr)
    If lMax < aPtr Then Exit Sub
    
    Do While aPtr < lMax
        If m_Stream(aPtr) = 0& Then Exit Do             ' padding; done
        For lValue = aPtr To aPtr + 2&                  ' validate header type (ASCII A-Z 0-9)
            Select Case m_Stream(lValue)
            Case 65 To 90, 48 To 57
            Case Else
                Exit Do                                 ' malformatted id3 tag
            End Select
        Next
        CopyMemory lValue, m_Stream(aPtr), 3&           ' get frame type
        lSize = pvReverseLong3(m_Stream(), aPtr + 3&)   ' and its overall size
        If lSize + aPtr <= lMax Then
            If (lValue And &HFFFFFF) = tagPIC And lSize > 0& Then ' zero-length frames are not uncommon
                ' PIC header
                ' aPtr + 6      : 1 byte text type (0=ASCII, 1=unicode)
                ' aPtr + 7      : 3 byte image format (PNG, JPG, etc)
                ' aPtr + 10     : 1 byte image category (i.e., album front, back, etc)
                ' aPtr + 11     : t bytes image description (depends on text type byte)
                ' aptr + 11+t   : binary data for the image
                '   Note: if image format is "-->" then image data is a URL & excluded from processing here
                If Not (m_Stream(aPtr + 7&) = 45 And m_Stream(aPtr + 8&) = 45) Then ' else dashes in the image type field
                    sType = pvStringFromArray(m_Stream(), aPtr + 7, 3)
                    ' determine length of image description (max of 64)
                    Select Case m_Stream(aPtr + 6&)
                        Case 0: imgOffset = lstrlen(VarPtr(m_Stream(aPtr + 11&))) + 1&
                        Case 1: imgOffset = lstrlenW(VarPtr(m_Stream(aPtr + 11&))) * 2& + 2&
                        Case Else: imgOffset = 0&
                    End Select
                    If imgOffset Then                     ' else malformed tag
                        imgOffset = imgOffset + aPtr + m_Stream(aPtr + 6&) + 11&
                        ' ^^ +11 = 6-10 in header above + 6 byte frame header
                        ReDim Preserve m_Frames(0 To m_Count)
                        With m_Frames(m_Count)
                            .Category = m_Stream(aPtr + 10&)
                            .Start = imgOffset
                            .Length = lSize - (imgOffset - aPtr - 6&) ' determine actual size of just image data
                            .MIME = sType
                            If .Length + .Start > aPtr + lSize + 6& Then m_Count = m_Count - 1& ' abort this frame
'                            Debug.Print "(3v2)PIC found "
'                            Debug.Print vbTab; "Size: "; .Length; " found at pos: "; .Start; m_Stream(aPtr + 9); m_Stream(5)
                        End With
                        m_Count = m_Count + 1&
                    End If
                End If
            End If
        End If
        aPtr = aPtr + lSize + 6&
    Loop

End Sub

Private Sub pvParse3v3()

    ' http://www.id3.org/id3v2.3.0
    
    ' id3v2.3.0 header block format
    ' offset 0 : 3 byte magic number: ID3
    ' offset 3 : 1 byte major version
    ' offset 4 : 1 byte revision
    ' offset 5 : 1 byte flag   (bit 7=unsync'd, 6=extended header)
    ' offset 6 : 4 byte sync-safe length

    Dim lMax As Long, aPtr As Long, lSize As Long
    Dim lValue As Long, imgOffset As Long, sType As String
    Const tagAPIC As Long = 1128878145                  ' only care about these frames
    
    ' frame header (10 bytes always)
    ' aPtr + 0 : 4 byte header type (i.e., APIC)
    ' aPtr + 4 : 4 byte header size as big endian long
    ' aPtr + 8 : 1 byte flag (not important for this parser)
    ' aPtr + 9 : 1 byte flag (includes whether frame is encrypted and/or comnpressed)
    
    aPtr = 10&
    If (m_Stream(5) And &H40) Then          ' extended header (length is not sync-safe)
        aPtr = aPtr + pvReverseLong4(m_Stream(), aPtr)
        If aPtr < 10& Then Exit Sub         ' indicates malformed or my parsing engine is wrong
    End If
    lMax = pvSyncSafeLong(m_Stream(), 6&)
    If (m_Stream(5) And &H80) Then lMax = lMax - pvSynchronizeTag(m_Stream(), aPtr, lMax - 10&)
    If lMax < aPtr Then Exit Sub
    
    Do While aPtr < lMax
        If m_Stream(aPtr) = 0& Then Exit Do             ' padding; done
        For lValue = aPtr To aPtr + 3&                  ' validate header type (ASCII A-Z 0-9)
            Select Case m_Stream(lValue)
            Case 65 To 90, 48 To 57
            Case 0
                ' these shouldn't really be allowed. Per specs frame names are 4 characters, not 3 + null
                ' but since we are not processing them anyway; just let it go
            Case Else
                Exit Do                                 ' malformatted id3 tag
            End Select
        Next
        CopyMemory lValue, m_Stream(aPtr), 4&           ' get frame type
        lSize = pvReverseLong4(m_Stream(), aPtr + 4&)   ' and its overall size
        If lSize < 0& Then Exit Do                      ' something is corrupted shouldn't have negative values
        If lSize + aPtr <= lMax Then
            If lValue = tagAPIC And lSize > 0& Then     ' zero-length frames are not uncommon
                ' APIC header
                ' aPtr + 10     : 1 byte text type (0=ASCII, 1=unicode with BOM)
                ' aPtr + 11     : n bytes image MIME type (null terminated non-unicode string)
                ' aPtr + 11+n   : 1 byte image category (i.e., album front, back, etc)
                ' aPtr + 12+n   : t bytes image description (depends on text type byte)
                ' aptr + 12+n+t : binary data for the image
                '   Note: if MIME type is "-->" then image data is a URL & excluded from processing here
                If (m_Stream(aPtr + 9&) And &HC0) = 0& Then             ' else compressed and/or encrypted
                    imgOffset = lstrlen(VarPtr(m_Stream(aPtr + 11&)))   ' get length of MIME type
                    If imgOffset = 3& Then                              ' malformed? should be MIME type or --> or nothing
                        If m_Stream(aPtr + 12&) = 45 Then imgOffset = -1& ' do not process URL links
                    End If
                    If imgOffset = 0& Or imgOffset > 2& Then
                        sType = pvStringFromArray(m_Stream(), aPtr + 11, imgOffset)
                        imgOffset = imgOffset + aPtr + 13&   ' 13 = 10 header, 1 null terminator, 1 text type, 1 category
                        ' determine length of image description (should be max of 64)
                        Select Case m_Stream(aPtr + 10&)
                            Case 0: lValue = lstrlen(VarPtr(m_Stream(imgOffset))) + 1&
                            Case 1: lValue = lstrlenW(VarPtr(m_Stream(imgOffset))) * 2& + 2&
                            Case Else: imgOffset = 0&
                        End Select
                        If imgOffset Then                               ' else malformated id3 tag
                            ReDim Preserve m_Frames(0 To m_Count)
                            With m_Frames(m_Count)
                                .Category = m_Stream(imgOffset - 1&)
                                .Start = imgOffset + lValue
                                .Length = lSize - (.Start - aPtr - 10&) ' determine actual size of just image data
                                .MIME = sType
                                If .Length + .Start > aPtr + lSize + 10& Then m_Count = m_Count - 1& ' abort this frame
'                                Debug.Print "(3v3)APIC found "
'                                Debug.Print vbTab; "Size: "; .Length; " found at pos: "; .Start; m_Stream(aPtr + 9); m_Stream(5)
                            End With
                            m_Count = m_Count + 1&
                        End If
                    End If
                End If
            End If
        End If
        aPtr = aPtr + lSize + 10&
    Loop

End Sub

Private Sub pvParse3v4()

    ' http://www.id3.org/id3v2.4.0-structure
    
    ' id3v2.4.0 header block format
    ' offset 0 : 3 byte magic number: ID3
    ' offset 3 : 1 byte major version
    ' offset 4 : 1 byte revision
    ' offset 5 : 1 byte flag   (bit 7=unsync'd, 6=extended header, 4=footer)
    ' offset 6 : 4 byte sync-safe length

    Dim lMax As Long, aPtr As Long, lSize As Long, lSizeOffset As Long
    Dim lValue As Long, imgOffset As Long, sType As String
    Const tagAPIC As Long = 1128878145                  ' only care about these frames
    
    ' frame header (10 bytes always)
    ' aPtr + 0 : 4 byte header type (i.e., APIC)
    ' aPtr + 4 : 4 byte header size as big endian long
    ' aPtr + 8 : 1 byte flag (not important for this parser)
    ' aPtr + 9 : 1 byte flag (includes whether frame is unsync'd, encrypted and/or comnpressed + other stuff)
    '   if data length bit set then 4 bytes added to end of frame
    '   if group bit set then 1 byte added to end of frame
    '   if encryption bit set then 1 byte added to end of frame (note that encrypted frames are not processed here)
    
    aPtr = 10&
    If (m_Stream(5) And &H40) Then          ' extended header (length is sync-safe)
        aPtr = aPtr + pvSyncSafeLong(m_Stream(), aPtr)
    End If
    lMax = pvSyncSafeLong(m_Stream(), 6&)
    If (m_Stream(5) And &H10) Then lMax = lMax - 10& ' exclude processing the footer
    If (m_Stream(5) And &H80) Then lMax = lMax - pvSynchronizeTag(m_Stream(), aPtr, lMax - 10&)
    If lMax < aPtr Then Exit Sub
    
    Do While aPtr < lMax
        If m_Stream(aPtr) = 0& Then Exit Do             ' padding; done
        For lValue = aPtr To aPtr + 3&                  ' validate header type (ASCII A-Z 0-9)
            Select Case m_Stream(lValue)
            Case 65 To 90, 48 To 57
            Case 0
                ' these shouldn't really be allowed. Per specs frame names are 4 characters, not 3 + null
                ' but since we are not processing them anyway; just let it go
            Case Else
                Exit Do                                 ' malformatted id3 tag
            End Select
        Next
        
        CopyMemory lValue, m_Stream(aPtr), 4&           ' get frame type
        lSize = pvSyncSafeLong(m_Stream(), aPtr + 4&)   ' and its overall size
        If lSize < 0& Then Exit Do                      ' something is corrupted shouldn't have negative values
        If lSize + aPtr <= lMax Then
            If lValue = tagAPIC And lSize > 0& Then         ' zero-length frames are not uncommon
                ' APIC header (no change from v2.0)
                ' aPtr + 10     : 1 byte text type (0=ASCII, 1=unicode with BOM, 2=unicode w/o BOM, 3=UTF-8)
                ' aPtr + 11     : n bytes image MIME type (null terminated non-unicode string)
                ' aPtr + 11+n   : 1 byte image category (i.e., album front, back, etc)
                ' aPtr + 12+n   : t bytes image description (depends on text type byte)
                ' aptr + 12+n+t : binary data for the image
                '   Note: if MIME type is "-->" then image data is a URL & excluded from processing here
                If (m_Stream(aPtr + 9&) And &HC) = 0& Then  ' else compressed and/or encrypted
                    If (m_Stream(aPtr + 9&) And &H2) Then   ' frame is unsync'd
                        lSizeOffset = pvSynchronizeTag(m_Stream(), aPtr + 10&, lSize)
                    Else
                        lSizeOffset = 0&
                    End If
                    If (m_Stream(aPtr + 9&) And &H40) Then  ' presensce of grouping flag; extra byte added to frame
                        lSizeOffset = lSizeOffset - 1&
                    End If
                    If (m_Stream(aPtr + 9&) And &H1) Then   ' data length flag; extra 4 byets added to frame
                        lSizeOffset = lSizeOffset - 4&
                    End If
                    
                    imgOffset = lstrlen(VarPtr(m_Stream(aPtr + 11&)))   ' get length of MIME type
                    If imgOffset = 3& Then                              ' malformed? should be MIME type or --> or nothing
                        If m_Stream(aPtr + 12&) = 45 Then imgOffset = -1& ' do not process URL links
                    End If
                    If imgOffset = 0& Or imgOffset > 2& Then            ' else can't be an official MIME value
                        sType = pvStringFromArray(m_Stream(), aPtr + 11, imgOffset)
                        imgOffset = imgOffset + aPtr + 13&   ' 13 = 10 header, 1 null terminator, 1 text type, 1 category
                        ' determine length of image description (max of 64)
                        Select Case m_Stream(aPtr + 10&)
                        Case 0, 3                                       ' single null terminator
                            lValue = lstrlen(VarPtr(m_Stream(imgOffset))) + 1&
                        Case 1, 2                                       ' double null terminator
                            lValue = lstrlenW(VarPtr(m_Stream(imgOffset))) * 2& + 2&
                        Case Else
                            imgOffset = 0&
                        End Select
                        If imgOffset Then
                            ReDim Preserve m_Frames(0 To m_Count)
                            With m_Frames(m_Count)
                                .Category = m_Stream(imgOffset - 1&)
                                .Start = imgOffset + lValue
                                .Length = lSize - (.Start - aPtr - 10&) - lSizeOffset ' determine actual size of just image data
                                .MIME = sType
                                If .Length + .Start > aPtr + lSize + 10& - lSizeOffset Then m_Count = m_Count - 1& ' abort this frame
'                                Debug.Print "(3v4)APIC found "
'                                Debug.Print vbTab; "Size: "; .Length; " found at pos: "; .Start; m_Stream(aPtr + 9); m_Stream(5), lSizeOffset; lSize
                            End With
                            m_Count = m_Count + 1&
                        End If
                    End If
                End If
            End If
        End If
        aPtr = aPtr + lSize + 10&
    Loop

End Sub

Private Sub pvParseWMA()

    ' WMA mini-parser
    ' http://msdn.microsoft.com/en-us/library/bb643323.aspx

    Dim aPtr As Long, sGUID As String
    Dim lSize As Long, lMax As Long, nrHdrObj As Long
    
    ' ASF header introduction
    ' offset 0  : 16 byte guid
    ' offset 16 : 8 byte block size
    
    ' First header is the Object Header
    ' 24 byte introduction
    ' offset 24 : 4 byte number of blocks within this header
    ' offset 28 : 1 byte reserved (value should always be 1)
    ' offset 29 : 1 byte reserved (value should always be 2)
    ' -- of all the types of blocks, we are only concerned with the Extended header & Extended content
    
    
    ' WMA format has GUIDs as block/section headers. Here are one's were intereseted in
    Const tagASF_Header_Extension_Object As String = "{5FBF03B5-A92E-11CF-8EE3-00C00C205365}"
    Const tagASF_Extended_Content_Description_Object = "{D2D0A440-E307-11D2-97F0-00A0C95EA850}"
    
    CopyMemory lMax, m_Stream(16), 4&           ' the overall size of the header
    CopyMemory nrHdrObj, m_Stream(24), 4&
    If nrHdrObj < 1& Then Exit Sub
    
    aPtr = 30&: lMax = lMax - 24&
    sGUID = String$(38, vbNullChar)
    For nrHdrObj = 0 To nrHdrObj - 1&
        ' get the GUID and the size of the block. Abort if any blocks are over 2gb
        If StringFromGUID2(VarPtr(m_Stream(aPtr)), StrPtr(sGUID), 39&) = 0& Then Exit For
        CopyMemory lSize, m_Stream(aPtr + 20&), 4&: If lSize <> 0& Then Exit For
        CopyMemory lSize, m_Stream(aPtr + 16&), 4& ' size of the block
        'Debug.Print "guid "; sGUID; lSize
        
        Select Case sGUID
        
        Case tagASF_Extended_Content_Description_Object
            If pvParseWMA_ExtContent(lSize + aPtr, aPtr + 24&) = False Then Exit For
            
        Case tagASF_Header_Extension_Object
            Call pvParseWMA_ExtHeader(lSize + aPtr, aPtr + 24&)
            Exit For    ' any blocks after the mandatory extended header are sound data; no more metadata possible
            
        Case Else
            
        End Select
        
        aPtr = aPtr + lSize
        If aPtr > lMax Then Exit For
    Next

End Sub

Private Function pvParseWMA_ExtContent(Ceiling As Long, Offset As Long) As Boolean


    ' Extended Content header
    ' 24 byte introduction
    ' + 0 bytes : 2 byte count of descriptors
    ' each descriptor will be in this format. No padding, all are consecutive
    ' + 0 : 2 byte description length
    ' + 2 : unicode descriptor name
    ' + 2 + name length : 2 byte data type
    ' + 4 + name length : 2 byte data length
    ' + 8 + name length : the data in format defned by parsed data type
    '   -- of all the possible records, the only ones we care about are WM/Picture
    
    ' The WM/Picture record looks like this
    '   + 0 : 1 byte image category. Same 20+ categories that ID3 uses for MP3s
    '   + 1 : 4 byte image data length
    '   + 5 : double null terminated, unicode MIME type
    '   + 5 + MIME size : double null terminated, unicode image description
    '   + 5 + MIME size + Image description : the image binary data
    
    Dim lRecs As Long, lSize As Long
    Dim sDescriptor As String
    Const tagPicture As String = "WM/Picture" & vbNullChar
    
    CopyMemory lRecs, m_Stream(Offset), 2&
    Offset = Offset + 2&
    For lRecs = 0& To lRecs - 1&
        lSize = 0&: CopyMemory lSize, m_Stream(Offset), 2&   ' record's name length
        sDescriptor = String$(lSize \ 2&, vbNullChar)     ' get the name
        CopyMemory ByVal StrPtr(sDescriptor), m_Stream(Offset + 2&), lSize
        Offset = Offset + lSize + 2&
        If sDescriptor = tagPicture Then                  ' is it WM/Picture?
            CopyMemory lSize, m_Stream(Offset), 2&    ' get data type
            If lSize = 1& Then
            CopyMemory lSize, m_Stream(Offset + 2&), 2& ' data size
                Offset = Offset + 4&
                ReDim Preserve m_Frames(0 To m_Count)
                With m_Frames(m_Count)
                    .Category = m_Stream(Offset)          ' cache image category (same as MP3)
                    CopyMemory .Length, m_Stream(Offset + 1&), 4& ' cache length of image
                    ' cache the MIME, double null terminated string
                    lSize = lstrlenW(ByVal VarPtr(m_Stream(Offset + 5&))) * 2&
                    If lSize Then
                        .MIME = String$(lSize \ 2&, vbNullChar)
                        CopyMemory ByVal StrPtr(.MIME), m_Stream(Offset + 5&), lSize
                    End If
                    Offset = Offset + lSize + 7& ' 5 = cat byte + size bytes + 2 byte terminator
                    ' skip the image description, double null terminated string
                    lSize = lstrlenW(ByVal VarPtr(m_Stream(Offset))) * 2 + 2&
                    .Start = Offset + lSize               ' cache where image starts
                    Offset = .Start + .Length             ' move pointer to end of block
                    If Offset > Ceiling Then m_Count = m_Count - 1&
                End With
                lSize = -1&
                m_Count = m_Count + 1&
            End If
        End If
        If lSize > -1& Then ' skipping this metadata record
            CopyMemory lSize, m_Stream(Offset + 2&), 2& ' record's data length
            Offset = Offset + lSize + 4&
        End If
        If Offset > Ceiling Then Exit Function
    Next
    pvParseWMA_ExtContent = True
    
End Function

Private Function pvParseWMA_ExtHeader(Ceiling As Long, Offset As Long) As Boolean

    ' Extended header
    ' 24 byte introduction
    ' + 0 : 16 byte GUID & must be 5FBF03B5-A92E-11CF-8EE3-00C00C205365
    ' + 16 : 2 byte reserved (value should always be 6)
    ' + 18 : 4 byte data length
    ' + 22 : the extended header data
    '   - this data consists of 0 or more consecutive blocks, in any order
    '   - of all the blocks that can exist, we only care about the MetaData & Extended Content blocks
    
    ' MetaData header
    ' 24 byte introduction
    ' + 0 bytes : 2 byte record count
    ' each metadata record will be in this format. No padding, all records are consecutive
    '   + 0  : 2 byte language list index
    '   + 2  : 2 byte stream number
    '   + 4  : 2 byte metadata name length
    '   + 6  : 2 byte data type flag (byte array, integer, long, etc)
    '   + 8  : 4 byte data length
    '   + 12 : metadata name where number of bytes defined @ offset 4
    '   + 12 + name length : the metadata in format defined @ offset 6
    '   total record length is 12 + value @ offset 4 + value @ offset 8
    '   -- of all the possible records, the only ones we care about are WM/Picture
    
    ' The WM/Picture metadata record looks like this
    '   + 0 : 1 byte image category. Same 20+ categories that ID3 uses for MP3s
    '   + 1 : 4 byte image data length
    '   + 5 : double null terminated, unicode MIME type
    '   + 5 + MIME size : double null terminated, unicode image description
    '   + 5 + MIME size + Image description : the image binary data
    
    Dim sGUID As String, sTag As String, lExtSize As Long
    Dim lSize As Long, lRecs As Long, lRecSize As Long
    
    Const tagASF_Reserved_1 As String = "{ABD3D211-A9BA-11CF-8EE6-00C00C205365}"
    Const tagASF_Metadata_Library_Object As String = "{44231C94-9498-49D1-A141-1D134E457054}"
    Const tagASF_Extended_Content_Description_Object = "{D2D0A440-E307-11D2-97F0-00A0C95EA850}"
    ' not a header, but a section 'name' within a metadata record
    Const tagPicture As String = "WM/Picture" & vbNullChar
    

    sGUID = String$(38, vbNullChar)
    Call StringFromGUID2(VarPtr(m_Stream(Offset)), StrPtr(sGUID), 39&)
    If sGUID <> tagASF_Reserved_1 Then Exit Function  ' per specs, this must be the next 16 bytes
    CopyMemory lExtSize, m_Stream(Offset + 18&), 4&   ' size of extended header
    Offset = Offset + 22&                             ' move pointer amount of bytes read in block
    lExtSize = lExtSize + Offset                      ' set max read size
    
    Do While Offset < lExtSize
        If StringFromGUID2(VarPtr(m_Stream(Offset)), StrPtr(sGUID), 39&) = 0& Then Exit Function
        CopyMemory lSize, m_Stream(Offset + 20&), 4&: If lSize <> 0& Then Exit Function
        CopyMemory lSize, m_Stream(Offset + 16&), 4&  ' size of extended header sub-block
'       Debug.Print vbTab; sGUID; lSize
        
        If sGUID = tagASF_Metadata_Library_Object Then
            CopyMemory lRecs, m_Stream(Offset + 24&), 2&: If lRecs = 0& Then Exit Function
            lRecSize = Offset + lSize                   ' set max read size & read nr of meta records
            Offset = Offset + 26&                       ' move pointer amount of bytes read in block
            For lRecs = 0& To lRecs - 1&
                lSize = 0&: CopyMemory lSize, m_Stream(Offset + 6&), 2&   ' data type
                If lSize = 1& Then          ' byte array which is what we are after
                    CopyMemory lSize, m_Stream(Offset + 4&), 2&     ' record's name length
                    If lSize Then
                        sTag = String$(lSize \ 2&, vbNullChar)     ' get the name
                        CopyMemory ByVal StrPtr(sTag), m_Stream(Offset + 12&), lSize
                        If sTag = tagPicture Then                  ' is it WM/Picture?
                            Offset = Offset + 12& + lSize           ' move pointer to where it starts
                            
                            ReDim Preserve m_Frames(0 To m_Count)
                            With m_Frames(m_Count)
                                .Category = m_Stream(Offset)        ' cache image category (same as MP3)
                                CopyMemory .Length, m_Stream(Offset + 1&), 4& ' cache length of image
                                ' skip the MIME, double null terminated string
                                lSize = lstrlenW(ByVal VarPtr(m_Stream(Offset + 5&))) * 2&
                                If lSize Then
                                    .MIME = String$(lSize \ 2&, vbNullChar)
                                    CopyMemory ByVal StrPtr(.MIME), m_Stream(Offset + 5&), lSize
                                End If
                                Offset = Offset + lSize + 7& ' 5 = cat byte + size bytes + 2 byte terminator
                                ' skip the image description, double null terminated string
                                lSize = lstrlenW(ByVal VarPtr(m_Stream(Offset))) * 2 + 2&
                                .Start = Offset + lSize             ' cache where image starts
                                Offset = .Start + .Length           ' move pointer to end of block
                                If Offset > lRecSize Then m_Count = m_Count - 1&
                            End With
                            
                            m_Count = m_Count + 1&
                            lSize = -1&                             ' flag indicating block 100% processed
                        End If
                    End If
                End If
                If lSize > -1& Then ' skipping this metadata record
                    lSize = 0&: CopyMemory lSize, m_Stream(Offset + 4&), 2& ' record's name length
                    Offset = Offset + lSize + 12&
                    CopyMemory lSize, m_Stream(Offset - lSize - 4&), 4& ' record's data length
                    Offset = Offset + lSize
                End If
                If Offset > lRecSize Then Exit Function
            Next
        Else
            Offset = Offset + lSize
            ' haven't seen any extended content contained within the extended header; but
            ' I couldn't find any documentation to say it is not possible -- so, just in case:
            If sGUID = tagASF_Extended_Content_Description_Object Then
                If pvParseWMA_ExtContent((Offset), Offset - lSize + 24&) = False Then Exit Function
            End If
        End If
        If Offset > Ceiling Then Exit Function
    Loop
    pvParseWMA_ExtHeader = True

End Function
