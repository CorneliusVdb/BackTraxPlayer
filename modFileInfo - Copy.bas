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

Public Sub LoadFileList(lvList As ListView, sStartPath As String, ExclIcons As Boolean, Optional IsDrive As Variant)
Dim fsMain As New FileSystemObject
Dim fsDrives As Drives
Dim fsDrive As Drive
Dim fsFolder As Folder
Dim fsFile As file
Dim sExt As String
Dim iPicIndex As Integer
Dim bEmptyFolder As Integer   '0=Empty, 1=no music, 2 = music

Screen.MousePointer = vbCustom
Screen.MouseIcon = frmExplorer.lblCursorPlaceHolder.MouseIcon
DoEvents


'TYPE : 0=DIR, 1=FOLDER, 2=FILE
lvList.ListItems.Clear

If IsMissing(IsDrive) Then
  If sStartPath = "" Then
    For Each fsDrive In fsMain.Drives
      If fsDrive.DriveType = 1 Or fsDrive.DriveType = 2 Then '1=MemCard, 2=HD, 3=Netwerk,4=CD-rom
        If ExclIcons Then
          Set mItem = lvList.ListItems.Add(, , fsFile.name)
        Else
          Set mItem = lvList.ListItems.Add(, , fsDrive.Path & "\", 0, 11)
        End If
        mItem.SubItems(1) = "0"
        mItem.SubItems(2) = fsDrive.Path & "\"
      End If
    Next fsDrive
    Screen.MousePointer = vbDefault
    DoEvents
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


'If sStartPath <> "" And sFolderPath = "" Then
Set fsFolder = fsMain.GetFolder(sStartPath)
  For Each fsFolder In fsMain.GetFolder(sStartPath).SubFolders
    If fsFolder.Attributes = Directory Or fsFolder.Attributes = 2064 Or fsFolder.Attributes = 2096 Then
      bEmptyFolder = 1
      'Set mItem = lvList.ListItems.Add(, , fsFolder.name, 0, 1)
      For Each fsFile In fsFolder.Files
        bEmptyFolder = 14
        sExt = Right(fsFile.name, 4)
        If InStr(1, filter, sExt) > 0 Then
          bEmptyFolder = 15
          Screen.MousePointer = vbDefault
          DoEvents
          Exit For
        End If
      Next fsFile
      'Now load directory with propper icon
      If ExclIcons Then
        Set mItem = lvList.ListItems.Add(, , fsFile.name)
      Else
        Set mItem = lvList.ListItems.Add(, , fsFolder.name, 0, bEmptyFolder)
      End If
      mItem.SubItems(1) = "1"
      mItem.SubItems(2) = fsFolder.Path
          
    End If
  Next fsFolder
  
  'Also check if any music files in the current folder
  Set fsFolder = fsMain.GetFolder(sStartPath)
  For Each fsFile In fsFolder.Files
    sExt = Right(fsFile.name, 4)
    If InStr(1, filter, sExt) > 0 Then
      Select Case UCase(sExt)
        Case ".MP3"
          iPicIndex = 2
        Case ".WAV"
          iPicIndex = 3
        Case ".WMA"
          iPicIndex = 4
        Case ".M4A"
          iPicIndex = 5
        Case ".AAC"
          iPicIndex = 6
        Case "FLAC"
          iPicIndex = 7
        Case ".MP4"
          iPicIndex = 8
        Case ".OOG"
          iPicIndex = 9
        Case ".AIF"
          iPicIndex = 10
        Case Else
          iPicIndex = 13
      End Select
      If ExclIcons Then
        Set mItem = lvList.ListItems.Add(, , fsFile.name)
      Else
        Set mItem = lvList.ListItems.Add(, , fsFile.name, 0, iPicIndex)
      End If
      mItem.SubItems(1) = "2"
      mItem.SubItems(2) = fsFolder.Path
    End If
  Next fsFile
    
    Screen.MousePointer = vbDefault
    DoEvents
    
  Exit Sub
'End If

End Sub

Function CheckMusicfile(sPath) As Boolean
Dim MusicFile
CheckMusicfile = False
MusicFile = Dir(sPath, vbNormal)   ' Retrieve the first entry.
If MusicFile <> "" Then
  If InStr(1, filter, MusicFile) > 0 Then
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

