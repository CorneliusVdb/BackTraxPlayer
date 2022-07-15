Attribute VB_Name = "modFileInfo"
'''Declare Function FindFirstFile Lib "kernel32" Alias _
'''   "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
'''   As WIN32_FIND_DATA) As Long
'''
'''   Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
'''   (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'''
'''   Declare Function GetFileAttributes Lib "kernel32" Alias _
'''   "GetFileAttributesA" (ByVal lpFileName As String) As Long
'''
'''   Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) _
'''   As Long
'''
'''   Declare Function FileTimeToLocalFileTime Lib "kernel32" _
'''   (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'''
'''   Declare Function FileTimeToSystemTime Lib "kernel32" _
'''   (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'''
'''   Public Const MAX_PATH = 260
'''   Public Const MAXDWORD = &HFFFF
'''   Public Const INVALID_HANDLE_VALUE = -1
'''   Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
'''   Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
'''   Public Const FILE_ATTRIBUTE_HIDDEN = &H2
'''   Public Const FILE_ATTRIBUTE_NORMAL = &H80
'''   Public Const FILE_ATTRIBUTE_READONLY = &H1
'''   Public Const FILE_ATTRIBUTE_SYSTEM = &H4
'''   Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
'''
'''   Type FILETIME
'''     dwLowDateTime As Long
'''     dwHighDateTime As Long
'''   End Type
'''
'''   Type WIN32_FIND_DATA
'''     dwFileAttributes As Long
'''     ftCreationTime As FILETIME
'''     ftLastAccessTime As FILETIME
'''     ftLastWriteTime As FILETIME
'''     nFileSizeHigh As Long
'''     nFileSizeLow As Long
'''     dwReserved0 As Long
'''     dwReserved1 As Long
'''     cFileName As String * MAX_PATH
'''     cAlternate As String * 14
'''   End Type
'''
'''   Type SYSTEMTIME
'''     wYear As Integer
'''     wMonth As Integer
'''     wDayOfWeek As Integer
'''     wDay As Integer
'''     wHour As Integer
'''     wMinute As Integer
'''     wSecond As Integer
'''     wMilliseconds As Integer
'''   End Type

Public Const vbLoadingLeft As Integer = 8295   '5790  '3855
Public Const iCounter As Long = 25    'Counter for loading display

   Public Function StripNulls(OriginalStr As String) As String
      If (InStr(OriginalStr, Chr(0)) > 0) Then
         OriginalStr = Left(OriginalStr, _
          InStr(OriginalStr, Chr(0)) - 1)
      End If
      StripNulls = OriginalStr
   End Function
   
Public Function ShowSpaceInfo(drvpath) As String
Dim FSO, D, s
Dim fAttrib As FileSystemObject, fFile

On Error Resume Next
Set fAttrib = CreateObject("Scripting.FileSystemObject")
Set fFile = fAttrib.GetFile(drvpath)
ShowSpaceInfo = "File Size: " & FormatNumber(fFile.Size / 1024, 0) & " Kbytes"

End Function

Public Function FileExists(sPath) As Boolean
Dim FSO, D, s
Dim fAttrib As FileSystemObject, fFile

On Error Resume Next
Set fAttrib = CreateObject("Scripting.FileSystemObject")
Set fFile = fAttrib.GetFile(sPath)

FileExists = fAttrib.FileExists(sPath)

End Function

Public Function DirExists(SDir) As Boolean
Dim FSO, D, s
Dim fAttrib As FileSystemObject, fFile

On Error Resume Next
Set fAttrib = CreateObject("Scripting.FileSystemObject")
Set fFile = fAttrib.GetFolder(SDir)
DirExists = fAttrib.FolderExists(SDir)

End Function

Public Sub LoadFileList(grdList As VSFlexGrid, sStartPath As String, ExclIcons As Boolean, Optional IsDrive As Variant)
Dim fsMain As New FileSystemObject
Dim fsDrives As Drives
Dim fsDrive As Drive
Dim fsFolder As Folder
Dim fsFile As file
Dim sExt As String
Dim iPicIndex As Integer
Dim bEmptyFolder As Integer   '0=Empty, 1=no music, 2 = music
Dim lCnt As Long
Dim iCntMusic As Long
Dim iRow As Long
Const iRowH As Integer = 600

On Error GoTo ErrHandler


'TYPE : 0=DIR, 1=FOLDER, 2=FILE
grdList.Enabled = True
'grdList.SelectedItem.EnsureVisible
frmExplorer.lblCnt.Caption = 0
frmExplorer.lblCnt.Visible = True
'DoEvents
lCnt = 0
iCntMusic = 0
iRow = 0

grdList.Rows = 1
grdList.Cols = 7
grdList.Clear
DoEvents
grdList.Redraw = False
'DoEvents

grdList.TextMatrix(0, 1) = ""
grdList.RowHeight(0) = 0
grdList.Row = 0
DoEvents


'=====================================================
' Load DRIVES, Network Drives and other Drives FIRST '
'=====================================================
If IsMissing(IsDrive) Then
  If sStartPath = "" Then
    For Each fsDrive In fsMain.Drives
      If fsDrive.DriveType = 1 Or fsDrive.DriveType = 2 Or fsDrive.DriveType = 3 Or fsDrive.DriveType = 4 Then '1=MemCard, 2=HD, 3=Netwerk,4=CD-rom
         If fsDrive.Path = "A:" Then GoTo ReadNextDrive
         iRow = iRow + 1
         grdList.Rows = iRow + 1
         grdList.Row = iRow
         grdList.RowHeight(iRow) = iRowH
        If ExclIcons Then
          grdList.TextMatrix(iRow, 1) = fsFile.name
        Else
          grdList.TextMatrix(iRow, 1) = fsDrive.Path & "\"
        End If
        grdList.RowData(iRow) = grdList.TextMatrix(iRow, 1)
         iCntMusic = grdList.Rows
         'grdList.Col = 0
         'Set grdList.CellPicture = frmExplorer.ImageList2.ItemPicture(iDrawImage)
         Select Case fsDrive.DriveType
            Case 1
               iDrawImage = 18
            Case 2
               iDrawImage = 16
            Case 3
               iDrawImage = 1
            Case 4
               iDrawImage = 17
         End Select
      '   grdList.TextMatrix(iRow, 5) = iDrawImage
      '   grdList.TextMatrix(iRow, 5) = ""
         
         grdList.TextMatrix(iRow, 6) = "0"
         grdList.TextMatrix(iRow, 2) = "0"
         grdList.TextMatrix(iRow, 3) = fsDrive.Path & "\"
         If iRow < 500 Then
            grdList.Col = 0
            Set grdList.CellPicture = frmExplorer.ImageList2.ItemPicture(iDrawImage)
         End If
      End If
ReadNextDrive:
    Next fsDrive
     If grdList.Row > 0 Then
        grdList.RowHeight(0) = 0
        grdList.Row = 1
        grdList.TopRow = 1
     End If
'     'Load the image now (only after all the columns have been loaded with the text...
'     For iRow = 1 To grdList.Rows - 1
'      grdList.Row = iRow
'      iDrawImage = Val(grdList.TextMatrix(iRow, 5))
'      grdList.Col = 0
'      Set grdList.CellPicture = frmExplorer.ImageList2.ItemPicture(iDrawImage)
'     Next iRow
    grdList.Redraw = True
    Exit Sub
  End If
End If


'==================================================================================
'FSO Attributes                                                                   '
'Normal     0     Normal file. No attributes are set.                             '
'ReadOnly   1     Read-only file. Attribute is read/write.                        '
'Hidden     2     Hidden file. Attribute is read/write.                           '
'system     4     System file. Attribute is read/write.                           '
'volume     8     Disk drive volume label. Attribute is read-only.                '
'Directory  16    Folder or directory. Attribute is read-only.                    '
'Archive    32    File has changed since last backup. Attribute is read/write.    '
'Alias      1024  Link or shortcut. Attribute is read-only.                       '
'Compressed 2048  Compressed file. Attribute is read-only.                        '
'                                                                                 '
'thus 2064 = Compressed + Directory                                               '
'thus 2096 = Compressed + Archived + Directory                                    '
'==================================================================================

'==================================
'Load FOLDERS into the list first '
'==================================
Set fsFolder = fsMain.GetFolder(sStartPath)
  For Each fsFolder In fsMain.GetFolder(sStartPath).SubFolders
    'If fsFolder.Attributes = Directory Or fsFolder.Attributes = 48 Or fsFolder.Attributes = 2064 Or fsFolder.Attributes = 2096 Or fsFolder.Attributes = 17 Then
      If fsFolder.Attributes = 22 Then GoTo ReadNextFolder  'Skip SYSTEM FOLDERS

''''         bEmptyFolder = 1
''''      'If fsFolder.Files.Count > 0 Then 'Skip folders thats empty...
''''         For Each fsFile In fsFolder.Files
''''           bEmptyFolder = 1  '14
''''           sExt = Right(fsFile.name, 4)
''''           If InStr(1, Filter, sExt) > 0 Then
''''             bEmptyFolder = 1  '15
''''             Exit For
''''           End If
''''            lCnt = lCnt + 1
''''            If lCnt Mod iCounter = 0 Then
''''               frmExplorer.lblCnt.Caption = lCnt
''''               frmExplorer.lvFiles.Refresh
''''            End If
''''           'DoEvents
''''         Next fsFile


      'Now load directory with propper icon
      lCnt = lCnt + 1
      If lCnt Mod iCounter = 0 Then
         frmExplorer.lblCnt.Caption = lCnt
         DoEvents
      End If
      iRow = iRow + 1
      grdList.Rows = iRow + 1
      grdList.Row = iRow
      grdList.RowHeight(iRow) = iRowH
      If ExclIcons Then
        iDrawImage = 3
         grdList.TextMatrix(iRow, 1) = fsFile.name
      Else
         iDrawImage = 2
         grdList.TextMatrix(iRow, 1) = fsFolder.name
      End If
      grdList.RowData(iRow) = grdList.TextMatrix(iRow, 1)
      iCntMusic = grdList.Rows
         
      grdList.TextMatrix(iRow, 2) = "1"
      grdList.TextMatrix(iRow, 3) = fsFolder.Path
      'grdList.TextMatrix(iRow, 5) = ""  'iDrawImage
      grdList.TextMatrix(iRow, 6) = "0"

      If iRow < 500 Then
         grdList.Col = 0
         Set grdList.CellPicture = frmExplorer.ImageList2.ItemPicture(iDrawImage)
      End If
      
      
ReadNextFolder:
  Next fsFolder
  
  '==============================================
  'Load FILES for the CURRENTLY selected FOLDER '
  '==============================================
  Set fsFolder = fsMain.GetFolder(sStartPath)
  For Each fsFile In fsFolder.Files
    sExt = Right(fsFile.name, 4)
    If InStr(1, UCase(Filter), UCase(sExt)) > 0 Then
      Select Case UCase(sExt)
        Case ".MP3"
          iPicIndex = 3
        Case ".WAV"
          iPicIndex = 4
        Case ".WMA"
          iPicIndex = 5
        Case ".MP4", ".MP2", ".MP1"
          iPicIndex = 6
        Case ".FLA", "FLAC"
          iPicIndex = 7
        Case ".AIF", "AIFF"
          iPicIndex = 8
        Case ".OGG", ".OGA"
          iPicIndex = 9
        Case ".APE"
         iPicIndex = 10
        Case ".AAC"
          iPicIndex = 11
        Case ".M4A"
          iPicIndex = 12
        Case Else
          iPicIndex = 15
      End Select
      
      If ExclIcons Then
         iDrawImage = 3
      Else
         iDrawImage = iPicIndex
      End If
      
      lCnt = lCnt + 1
      If lCnt Mod iCounter = 0 Then
         frmExplorer.lblCnt.Caption = CStr(lCnt)
         'frmExplorer.lblCnt.Refresh
         DoEvents
      End If
      iRow = iRow + 1
      grdList.Rows = iRow + 1
      grdList.Row = iRow
      grdList.RowHeight(iRow) = iRowH
      grdList.TextMatrix(iRow, 1) = fsFile.name
      grdList.RowData(iRow) = grdList.TextMatrix(iRow, 1)
      grdList.TextMatrix(iRow, 2) = "2"
      grdList.TextMatrix(iRow, 3) = fsFolder.Path
      'grdList.TextMatrix(iRow, 5) = ""  'iDrawImage
      grdList.TextMatrix(iRow, 6) = "13"
'      If iRow Mod 50 = 0 Then
'         Debug.Print "Cnt:" & iRow & "  File:" & fsFile.name
'      End If
      'It seems that the grid has some problem loading too many images. We limit images to 500, but still load the rest of the data...
      If iRow < 500 Then
         grdList.Col = 0
         Set grdList.CellPicture = frmExplorer.ImageList2.ItemPicture(iDrawImage)
      End If
      iCntMusic = grdList.Rows
    End If
  Next fsFile
  
 ' grdList.ScrollBars = flexScrollBarVertical
  If grdList.Row > 0 Then
     grdList.RowHeight(0) = 0
     grdList.Row = 1
     grdList.TopRow = 1
  End If
'  'Load the image now (only after all the columns have been loaded with the text...
'  For iRow = 1 To grdList.Rows - 1
'   grdList.Row = iRow
'   iDrawImage = Val(grdList.TextMatrix(iRow, 5))
'   grdList.Col = 0
'   Set grdList.CellPicture = frmExplorer.ImageList2.ItemPicture(iDrawImage)
''''   If Val(grdList.TextMatrix(iRow, 6)) > 0 Then
''''      grdList.Col = 4
''''      Set grdList.CellPicture = frmExplorer.ImageList2a.ItemPicture(1)
''''   End If
'
'  Next iRow
  
  grdList.Redraw = True
  grdList.Refresh
  
  DoEvents
  
  Exit Sub

ErrHandler:

If Err.Number = 76 Then
   MsgBox "No CD was found in drive.", vbInformation, "Load Error"
ElseIf Err.Number <> 70 Then
   MsgBox Err.Description
End If

grdList.Redraw = True
grdList.Refresh

'Resume Next

End Sub

Function CheckMusicfile(sPath) As Boolean
Dim MusicFile
CheckMusicfile = False
MusicFile = Dir(sPath, vbNormal)   ' Retrieve the first entry.
If MusicFile <> "" Then
  If InStr(1, Filter, MusicFile) > 0 Then
    CheckMusicfile = True
  End If
End If

End Function

Public Function ListSubDirs(Path) As Boolean

ListSubDirs = False
On Error Resume Next

Dim Count, D(), i, DirName      ' Declare variables.
DirName = Dir(Path, 16)         ' Get first directory name.
Do While DirName <> ""
    ' A file or directory name was returned
    If DirName <> "." And DirName <> ".." Then
        ' Not a parent or current directory entry so process it
        If GetAttr(Path + DirName) = 16 Then
            ' This is a directory
            ListSubDirs = True
            If (Count Mod 10) = 0 Then
                ' Resize the array
                ReDim Preserve D(Count + 10)
            End If
            Count = Count + 1   ' Increment counter.
            D(Count) = DirName  ' Add directory name to array
        End If
    End If
    DirName = Dir   ' Get another directory name.
Loop
    
End Function

