Attribute VB_Name = "modGlobal"
Option Explicit

Public MainApp As Boolean

Public PlArr()  'Start at 1 with both dimensions
Public Enum PLA
   eTtle = 0    'Title
   efTtle = 1   'Original Title
   eFN = 2      'Filename
   eVol = 3     'Volume
   eAve = 4     'Average
   eClr = 5     'Color
   eLs = 6      'Leading Silences
   eTp = 7      'Time Played
End Enum

Public Const bDoEq As Boolean = False ' True

Public iPageno As Integer
Public sOverwritePalet As String
Public bOverwritePalet As Boolean
Public ClearPalet As Boolean

Public gColor As Integer
Public ApplyStandardTheme As Boolean

Public mItem  As ListItem
Public PaletteName As String
Public bSavePalette As Boolean
Public FSO As New FileSystemObject
Public lDeviceNo As Long
Public lDeviceSingle As Long
Public lOldDeviceSingle As Long
Public Resp As Integer
Public RepeatSong As Boolean
Public Fadeout As Single
Public Const iDuration As Long = 1500
Public Const iInterval As Integer = 100
Public Const MaxLevelVal As Integer = 30473

Public SerialNumber As String
Public bSkipValidation As Boolean
Public iCntDemo As Integer

Public ListOfPlugins As String

Dim sChr(11) As String
Dim sSeed(11) As String

Public aEQ() As Integer              'Storage for Each Button's Eq settings
Public Const MaxEqs As Integer = 3  '14  'EQ sliders
Public EqFreq(20)                    'Contains the Frequenzy values for each band
Public iEqBands As Integer           'Counter for Eq Bands
Public iButCnt As Integer            'Counter for Buttons on screen
Public floatable As Long             ' floating-point channel support
Public fxEQ As Long

Public cStartPos As Double
  
Public LoadARR() As String

Public BusyPlaying As Boolean

Public fx(15) As Long        ' 10 eq band      '+ reverb


'======================================================
' Change this to TRUE when a DEMO system is created
'Public Const DemoFlag As Boolean = True
Public DemoFlag As Boolean
Public Const DemoMsg1 As String = "This is a DEMO system."
Public Const DemoMsg2 As String = "Playback is limited to 20 seconds."
Public Const DemoMsg3 As String = "Playback is limited to 5 songs ONLY."
Public Const DemoMsg4 As String = "Playback is limited to 5 songs ONLY and 20 seconds per song."
Public Const DemoHeading As String = "DEMO MODE"
Public Const DemoMax As Integer = 5    '5 songs limit
Public DemoCnt As Integer
Public Const DemoTime As Long = 20     '20 seconds
'======================================================

Public chan(6) As Long
Public iPlayingChan As Long
Public OutDev(1) As Long
Public Filter As String

Public FilenameToLoad As String
Public bTagsUpdated  As Boolean
Public Duration(6) As Single
Public MaxWidth As Long
Public IncrVal As Long

Public aRnd() As Integer

Public bSinglePlayer As Boolean

Public LeftMChan As Integer
Public RightMChan As Integer

Public Left1Chan As Integer
Public Right1Chan As Integer
Public Left2Chan As Integer
Public Right2Chan As Integer
Public Stop1 As Boolean
Public Stop2 As Boolean


'Global variables for volume control
Public lVolume As Long
Public fVolume As Single
Public lBass As Integer
Public lMid As Integer
Public lHigh As Integer

'Defines Time Class
Public bassTime As cbass_time   ' Class module Handle

Public Const vbDefaultBack As Long = &HFEC5CA      '&H808080
Public Const vbDefaultProg As Long = &H404040

Public Const vbSelected As Long = &HE7DB49

Public Const vbOrange As Long = &H64FF&       '&H27CF7      ' &H80FF&
Public Const vbDGreen As Long = &HC000&

Public Const vbButRed As Long = &H3D02A0      '&HF0863
Public Const vbButBlue As Long = &HD02400
Public Const vbButGreen As Long = &H76120
Public Const vbButPurple As Long = &H7B0096
Public Const vbButGold As Long = &HD657D

Public Const vbButSelRed As Long = &H938BF0       '&H5E56BB     '&H7269CA
Public Const vbButSelBlue As Long = &HE8B785      '&HB77951    '&HC69361
Public Const vbButSelGreen As Long = &H8BECB0     '&H55BB73   '&H67CA8C
Public Const vbButSelPurple As Long = &HF2B2ED    '&HBE6EB5  '&HCA7FC3
Public Const vbButSelGold As Long = &H81EAF3      '&H5BC2DB     '&H73CFE2

Public Const vbPlaylistSelColor As Long = vbGreen  'vbYellow   '&H80C0FF
Public Const vbPlayListBackColor As Long = &H1B4F0F
Public Const vbPlayListSelBackColor As Long = vbRed

Public Const vbProgressRed As Long = &HFFFF&
Public Const vbProgressGreen As Long = &HFB91&   '&HC000C0
Public Const vbProgressBlue As Long = &HFF&
Public Const vbProgressPurple As Long = &HFF00&
Public Const vbProgressGold As Long = &H80FF&



Public Const vbProgressBackRed As Long = &H40&
Public Const vbProgressBackGreen As Long = &H390C&
Public Const vbProgressBackBlue As Long = &H400000
Public Const vbProgressBackPurple As Long = &H400040
Public Const vbProgressBackGold As Long = &H4040&

Public Const vbDirectionColor As Long = &HE7DB49      '&HE4C761

'New colors.
Public vbNDefault As Long   '= &HE2FFB3     '&H400000    ' &H404000     '&HA4ABBB     '&H79533C     '&H2F2F2F      '&HE1FFFF     '&HFEC5CA     '&H404040         '&H404040       'Default (0)
Public vbNDefaultFore As Long '= &H80000008

Public Const vbNColor1 As Long = &H808080
Public Const vbNColor2 As Long = &HC000C0
Public Const vbNColor3 As Long = &HFF3C4A
Public Const vbNColor4 As Long = &H565810
Public Const vbNColor5 As Long = &H7D0D&
Public Const vbNColor6 As Long = &H8080&
Public Const vbNColor7 As Long = &H40C0&
Public Const vbNColor8 As Long = &H80&
Public Const vbNColor9 As Long = &H400482
Public Const vbNColor10 As Long = &HFB1594
Public Const vbNColor11 As Long = &HF89F0C
Public Const vbNColor12 As Long = &H808000
Public Const vbNColor13 As Long = &H5128&
Public Const vbNColor14 As Long = &H4301F1
Public Const vbNColor15 As Long = &H7CEC&
Public Const vbNColor16 As Long = &H1022FC

Public Const vbNDYellow As Long = vbYellow   '&H24DEFF    '&H007DBF7B&
Public Const vbNCompleted As Long = &H260F35
Public Const vbFadeOut As Long = &H80FF&       '&H24DEFF    '&H007DBF7B&

Public vbKeepColor As Integer

Public iDrawImage As Integer

'Variable to hold/limit the buttons on the palette
Public iMaxBut As Integer  ' = 25

Public Const vbButArtist As Long = &HC0FFFF
Public Const vbButSelArtist As Long = &H3A3332

'Registry keys
Public Const regMainKey = "LilacPro Systems"
Public Const regSubKey = "Options"
Public regString              As String
Public gregMainKey            As String

Public InitY As Single
Public EndY As Single
Public VelosityY As Single
Public VelosityStart As String
Public VelosityEnd As String
Public StartIndex As Integer
Public EndIndex As Integer
Public LongPress As Integer
Public XPos As Single
Public YPos As Single
Public SelPlayerIndex As Integer

Public bClearButton As Boolean
Public bLoadButton As Boolean
Public bExitSetup As Boolean
Public bDucking As Boolean
Public bTagEdit As Boolean
Public bTagEditMP3 As Boolean

Public ButLeft As Long
Public ButTop As Long
Public iButtonDirection As Integer
Public iButtonStreams As Integer
Public iButtonMaxSelected As Integer
Public iButtonRemoveSilence As Integer
Public iButtonPlayStopPause As Integer
Public iButtonDefaultColor As Integer
Public iAdjustVol As Integer
Public iSecureMode As Integer
Public sSecurePWD As String
Public iAutoAdvance As Integer
Public ScreenOptions(0 To 10) As Integer
   
Public Const ButDirection As Integer = 0
Public Const ButPlaystop As Integer = 1
Public Const ButAutoAdvance As Integer = 2
Public Const ButRemSilence As Integer = 3
Public Const ButStreams As Integer = 4
Public Const ButDefColor As Integer = 5
Public Const ButSecure As Integer = 6
Public Const ButMaxSel As Integer = 7
Public Const ButAdjVol As Integer = 8

Public Const LongPressCnt As Integer = 15  'Timer is 50, and we want to do this 30 times to get to 1500 (1.5 seconds)


'Public iButtonStreams As Integer

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Access the GetCursorPos function in user32.dll
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

' GetCursorPos requires a variable declared as a custom data type
' that will hold two integers, one for x value and one for y value
Type POINTAPI
   X_Pos As Long
   Y_Pos As Long
End Type

' Dimension the variable that will hold the x and y cursor positions
Public HoldCursorPos As POINTAPI

'Declarations
Public X                      As Double
Public GVersion               As String
'=====================================================================
'  These values can be found in the VERSION.INI file on the server   '
'  To customize for each project, check the VERSION.INI file and     '
'  make the appropriate entries there-in.                            '
'=====================================================================
Public strAppPath             As String
Public strDistrib             As String
Public strSetupPath           As String
Public strVersionPath         As String
'The path to the VERSION.INI
Public strVersionFile         As String ' = "\\ASSUPRIMARY\AsiaInst\Version.ini"

'Declare the functions to read ini files
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Private Type SYSTEMTIME
   wYear                As Integer
   wMonth               As Integer
   wDayOfWeek           As Integer
   wDay                 As Integer
   wHour                As Integer
   wMinute              As Integer
   wSecond              As Integer
   wMilliseconds        As Integer
End Type

Public Const MaxLits = 17
Public Enum MsgLits
  Ready = 0
  Start = 1
  VerifyLic = 2
  MainPlugin = 3
  Otherplugins = 4
  Init = 5
  ChkDisk = 6
  LoadEnv = 7
  GetReg = 8
  FormatLayout = 9
  LoadSoundCards = 10
  OpenPlayList = 11
  Finalise = 12
  ProdName = 13 '"BackTrax Player Professional"
  Running = 14 '"BackTrax Player is already running."
  Copyrite = 15 '"Copyright © # Lilac Productions. All Rights Reserved."
  Version = 16 '"Version"
  InitSoundCard = 17
End Enum

Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

' Registry manipulation API's (32-bit)
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_NO_MORE_ITEMS = 259&

Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const HKeyVersion = "SOFTWARE\LilacPro\VersionCount"
'==============================================================================
' Add the following line in the startup procedure (ie. Form_load or Sub Main) '
'==============================================================================
'  CheckVersion                                                               '
'==============================================================================

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Private Declare Function SHBrowseForFolder Lib "shell32" _
'   (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type LV_Item
   mask As Long
   iItem As Long
   iSubItem As Long
   State As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   lParam As Long
   iIndent As Long
End Type

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE As Long = (LVM_FIRST + 40)
Private Const LVM_SETITEMSTATE As Long = (LVM_FIRST + 43)
Private Const LVIS_FOCUSED As Long = &H1
Private Const LVIS_SELECTED As Long = &H2
Private Const LVIF_STATE As Long = &H8
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVS_EX_SUBITEMIMAGES As Long = &H2&
Private Const LVIF_IMAGE As Long = &H2&
Private Const LVM_SETITEM As Long = (LVM_FIRST + 6)



Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type Mp3IDTag
      Songname As String * 30
      Artist As String * 30
      Album As String * 30
      Year As String * 4
      Comment As String * 30
      Genre As String * 1
End Type

Const WM_USER = &H400
Const CCM_FIRST = &H2000&
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Const PBM_SETBARCOLOR = (WM_USER + 9)


Public Id3TagArr(5) As String

Private m_cID3v1 As New cMP3ID3v1
Private m_cID3v2 As New cMP3ID3v2
'Private objTag As ID3v23x.clsID3v2


Public Type SYSTEM_INFO
   dwOemID As Long
   dwPageSize As Long
   lpMinimumApplicationAddress As Long
   lpMaximumApplicationAddress As Long
   dwActiveProcessorMask As Long
   dwNumberOrfProcessors As Long
   dwProcessorType As Long
   dwAllocationGranularity As Long
   dwReserved As Long
End Type

Public Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Sub GetNativeSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)


Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
                                                                  ByVal dwMaxNameChars As Long, _
                                                                  ByVal pszColorBuff As Long, _
                                                                  ByVal cchMaxColorChars As Long, _
                                                                  ByVal pszSizeBuff As Long, _
                                                                  ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeFilename Lib "uxtheme.dll" _
   (ByVal hTheme As Long, _
    ByVal iPartId As Long, _
    ByVal iStateId As Long, _
    ByVal iPropId As Long, _
    pszThemeFileName As Long, _
    ByVal cchMaxBuffChars As Long _
   ) As Long

Private m_iTheme As Long


Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_BOTH = 3

Public Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)


Public Const hlpButtons As Integer = 1
Public Const hlpFeedback As Integer = 1000
Public Const hlpIntro As Integer = 1001
Public Const hlpLoadNewMusic As Integer = 1002
Public Const hlpLoadSong As Integer = 1003
Public Const hlpPlaysong As Integer = 1004

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Long) As Long
Public MonCount As Long


Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwnewlong As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Long, ByVal bErase As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Public Enum enuTBType
    enuTB_FLAT = 1
    enuTB_STANDARD = 2
End Enum

Private Const GCL_HBRBACKGROUND = (-10)

Public palletArr(100)
Public DiskSpaceTot As Currency
Public DiskSpaceFree As Currency
Public DiskSpaceFreeMB As String

Private Declare Function SHGetDiskFreeSpace Lib "shell32" Alias "SHGetDiskFreeSpaceA" (ByVal pszVolume As String, pqwFreeCaller As Currency, pqwTot As Currency, pqwFree As Currency) As Long

Public Sub WriteLog(sText As String)
Dim ff
Dim sFilename As String

'Exit Sub

   ff = FreeFile
   sFilename = App.Path & "\Loading.log"
   If FSO.FileExists(sFilename) Then
    Open sFilename For Append As #ff
   Else
    Open sFilename For Output As #ff
   End If
   Print #ff, sText
      
   Close #ff

End Sub

Public Sub UpdateSongFreqEq(FreqIndex As Integer, GainVal As Integer)

'Dim p As BASS_DX8_PARAMEQ

    Dim eq As BASS_BFX_PEAKEQ

    eq.lBand = FreqIndex    ' Band values you would like to get

    Call BASS_FXGetParameters(fxEQ, eq)
    eq.fGain = GainVal * -1
    Call BASS_FXSetParameters(fxEQ, eq)
    
    

'For i = 1 To MaxEqs
'  Call BASS_FXGetParameters(fx(FreqIndex), p)
'  p.fGain = 10# - GainVal
'  Call BASS_FXSetParameters(fx(FreqIndex), p)
'Next i
'    Else
'        Dim p1 As BASS_DX8_REVERB
'        Call BASS_FXGetParameters(fx(b), p1)
'        p1.fReverbMix = -0.012 * v * v * v
'        Call BASS_FXSetParameters(fx(b), p1)
   ' End If
    
    
End Sub

Public Sub SetupEqFx(chanEQ As Long)
    ' setup the effects
  '  Dim p As BASS_DX8_PARAMEQ
    Dim i As Integer
    Dim eq As BASS_BFX_PEAKEQ
  
    ' set peaking equalizer effect with no bands
    fxEQ = BASS_ChannelSetFX(chanEQ, BASS_FX_BFX_PEAKEQ, 0)
    eq.fBandwidth = 2.5
    eq.fQ = 0#
    eq.fGain = 0#
    eq.lChannel = BASS_BFX_CHANALL
    
    For i = 1 To MaxEqs
      eq.lBand = i
      eq.fCenter = EqFreq(i)
      'Call BASS_FXSetParameters(fx(I), p)
      Call BASS_FXSetParameters(fxEQ, eq)
    Next i
    
End Sub


Public Sub SetButtonEqArrayValue(Index As Integer, FreqBand As Integer, FreqBandVal As Integer)

'ALWAYS ONLY do one button(song's) values, I pass the button index into this function to populate the eq settings array
'this is to get the values LATER when I open the EQ screen for a button

'For iButCnt = 1 To 30         'We can only do 30 buttons
  'For iEqBands = 1 To MaxEqs
    aEQ(Index, FreqBand) = FreqBandVal    'Slider value at default (middle is 10)
  'Next iEqBands
'Next iButCnt

End Sub

'Public Sub GetEqArrayValues(Index As Integer)
''For iButCnt = 1 To 30         'We can only do 30 buttons
'  'For iEqBands = 1 To MaxEqs
'    aEQ(Index, FreqBand) = FreqBandVal    'Slider value at default (middle is 10)
'  'Next iEqBands
''Next iButCnt
'
'
'End Sub

Public Sub ClearEqArrayFreq(Index As Integer, Band As Integer)

ReDim aEQ(3, 30) 'Second dimension = value, First dimension = Indicator (0 = index of song, 1 = band to set, 3 = freq of band)


End Sub


Public Function Win3(T As String) As String
  'We use win3 to convert the path "C:\version5\" to "C:\version5" or
  'you will get "Path Not Found" errors on "Chdir" function

  T$ = Trim$(T$)
    
  If Right$(T$, 1) = "\" Then
    Win3 = Left$(T$, Len(T$) - 1)
  Else
    Win3 = T$
  End If

End Function

Public Function GetMonitorCount() As Long
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, GetMonitorCount
End Function

Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hDCMonitor As Long, ByVal lprcMonitor As Long, dwData As Long) As Long
    dwData = dwData + 1 'increase monitor count
    'Debug.Print dwData, hMonitor
    MonitorEnumProc = 1
End Function


Public Sub GetFreeSpace(DrvToCheck As String)
Dim FreeCaller As Currency
Dim Tot As Currency
Dim Free As Currency

On Error Resume Next

SHGetDiskFreeSpace DrvToCheck, FreeCaller, Tot, Free
'SHGetDiskFreeSpace "C:\", FreeCaller, Tot, Free

DiskSpaceTot = Round((Tot * 10000) / 1024 / 1024 / 1024, 1)
DiskSpaceFree = Round((Free * 10000) / 1024 / 1024 / 1024, 1)
DiskSpaceFreeMB = Format(Round((Free * 10000) / 1024 / 1024, 1), "###,###,##0")


'DiskSpaceTot = Format$(Tot * 10000, "###,###,###,##0")
'DiskSpaceFree = Format$(Free * 10000, "###,###,###,##0")

'MsgBox "Free space to caller: " + Format$(FreeCaller * 10000, "###,###,###,##0") + vbCrLf + _
'"Total space: " + Format$(Tot * 10000, "###,###,###,##0") + vbCrLf + _
'"Free space: " + Format$(Free * 10000, "###,###,###,##0")

End Sub

Function CheckNumbers(CheckValue As Integer) As Boolean
Dim iChk As Integer
CheckNumbers = True

For iChk = 1 To 9
   If sChr(iChk) = "" Then Exit For
   If sChr(iChk) = CheckValue Then
      CheckNumbers = False
      Exit For
   End If
Next iChk

End Function

Function CheckSeed(CheckValue As Integer) As Boolean
Dim iChk As Integer
CheckSeed = True

For iChk = 1 To 9
   If sSeed(iChk) = "" Then Exit For
   If sSeed(iChk) = CheckValue Then
      CheckSeed = False
      Exit For
   End If
Next iChk

End Function

Public Function GenerateNewSerial() As String
Dim lSerialKey As Integer
Dim sSerialNum As String
Dim sWork As String
Dim iSerialSeed As String
Dim i As Integer
Dim iChars As Integer
Dim iChk As Integer
Dim iLoop() As String


GenerateNewSerial = ""
Erase sSeed
Erase sChr

'RandomInteger = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
'0=48
'9=57
'a=65
'z=90

Randomize

'Generate the ALPHA characters
For i = 1 To 4
GenNumber:
   iChars = GenRndNumber(65, 90)
   If Not CheckNumbers(iChars) Then
      GoTo GenNumber
   End If
   sChr(i) = iChars
   'Debug.Print "Int:" & sChr(i) & "   Char:" & Chr(sChr(i))
Next i

'Generate the Number characters
For i = 5 To 9
GenNumber1:
   iChars = GenRndNumber(48, 57)
   If Not CheckNumbers(iChars) Then
      GoTo GenNumber1
   End If
   sChr(i) = iChars
   'Debug.Print "Int:" & sChr(i) & "   Char:" & Chr(sChr(i))
Next i

'Scramble the characters using a random index to load the text from the array
For i = 1 To 9
GenNumber2:
   iChars = GenRndNumber(1, 9)
   If Not CheckSeed(iChars) Then
      GoTo GenNumber2
   End If
   sSeed(i) = iChars
   'Debug.Print "Number order (" & i & ") : " & sSeed(i)
   sWork = sWork & iChars & "~"
Next i

sWork = Left(sWork, Len(sWork) - 1)
sWork = "~" & sWork

'Debug.Print "Number order " & sWork
iLoop = Split(sWork, "~")
sWork = ""
For i = 1 To UBound(iLoop)
   sWork = sWork & Chr(sChr(iLoop(i)))
Next i

'Debug.Print "Number order " & sWork
Randomize 9
GetNewKey:
lSerialKey = Rnd * 9

If lSerialKey = 0 Then GoTo GetNewKey

Select Case lSerialKey
   Case 1, 4
      iSerialSeed = "X"
   Case 2, 5
      iSerialSeed = "4"
   Case 3, 6
      iSerialSeed = "F"
   Case 4, 7
      iSerialSeed = "H"
   Case 5, 8
      iSerialSeed = "7"
   Case Else
      iSerialSeed = "D"
End Select

sSerialNum = ""
 
If lSerialKey > 1 Then
   sSerialNum = Mid(sWork, 1, lSerialKey - 1) & iSerialSeed & Mid(sWork, lSerialKey)
Else
   sSerialNum = iSerialSeed & Mid(sWork, lSerialKey)
End If

sSerialNum = CStr(lSerialKey) & sSerialNum

GenerateNewSerial = sSerialNum

'If ValidateSerial(GenerateNewSerial) Then
'   MsgBox "Valid"
'End If

End Function

Function GenRndNumber(iLowerBound As Integer, iUpperBound As Integer) As Integer

GenRndNumber = Int((iUpperBound - iLowerBound + 1) * Rnd + iLowerBound)

End Function

Public Function ValidateSerial(sSerial As String) As Boolean
Dim CheckDiget As Integer
Dim sWork As String
Dim sSerialSeed As String

On Error Resume Next

ValidateSerial = False
If sSerial = "" Then Exit Function
If Len(sSerial) <> 11 Then Exit Function
'If Not IsNumeric(sSerial) Then Exit Function

sWork = Mid(sSerial, 2)

CheckDiget = IIf(CInt(Left(sSerial, 1)) = 0, 1, CInt(Left(sSerial, 1)))

Select Case CheckDiget
   Case 1, 4
      sSerialSeed = "X"
   Case 2, 5
      sSerialSeed = "4"
   Case 3, 6
      sSerialSeed = "F"
   Case 4, 7
      sSerialSeed = "H"
   Case 5, 8
      sSerialSeed = "7"
   Case Else
      sSerialSeed = "D"
End Select

If Mid(sWork, CheckDiget, 1) = sSerialSeed Then ValidateSerial = True

End Function

Public Function GetFadeOutValue() As Single

'Public Fadeout As Integer
'Public Const iDuration As Integer = 5000
'Public Const iInterval As Integer = 100

GetFadeOutValue = CSng(100 / (iDuration / iInterval))

End Function

Public Sub ConvertOldPalettes()
Dim iPos As Integer
Dim iInner As Integer
Dim FD
Dim FileToOpen As String
Dim sHeading As String
Dim sStr As String
Dim sArrH()
Dim sKeepHeading As String
Dim bHeadingFound As Boolean
Dim bEnd As Boolean
Dim strPath As String
Dim tmpI As Integer
Dim sArr() As String
Dim sTemp As String
Dim FD1
Dim sNow As String
Dim iCntRows As Integer
Dim iMid As Integer

Dim ConvertThisFile As Boolean

On Error GoTo err1

FD = FreeFile
ConvertThisFile = False

'Loop through files in palets dir, and load filenames into each. Also read first line to determine order for sort...
Dim newFolder As Folder
Dim NewFile As file
Dim newFileName As String

ReDim sArrH(200, 2)


Set newFolder = FSO.GetFolder(App.Path & "\Palets")
ReadFiless:
For Each NewFile In newFolder.Files
  newFileName = NewFile
  Open NewFile For Input As FD
   
  Line Input #FD, sHeading
  If InStr(1, sHeading, ":") = 3 Then 'Means old file, so convert this
    FileToOpen = NewFile
    ConvertThisFile = True
    ReDim sArr(180)
    tmpI = 0
    'set the first line, since we just read this
    sNow = Format(Now, "YYYYMMDDHHmmSS")
    sArr(tmpI) = "000:" & sNow
    
'    iInner = 0
    'Read through file and load content into an array
    Do Until (EOF(FD) = True)
       Line Input #FD, sTemp
       tmpI = tmpI + 1
       sArr(tmpI) = Format(tmpI, "00") & ":" & Mid(sTemp, 4)
''       iInner = iInner + 1
''       If iInner > 16 Then iInner = 1
''       'Re-Create EMPTY entries if nothing in file
''       If sTemp = "" Or Len(sTemp) < 4 Then
''         If tmpI = 0 Then
''           sArr(tmpI) = "000:" & sNow
''         ElseIf tmpI < 17 Then
''           sArr(tmpI) = CStr(1) & Format(iInner, "00") & ":"
''         ElseIf tmpI < 33 Then
''           sArr(tmpI) = CStr(2) & Format(iInner, "00") & ":"
''         ElseIf tmpI < 49 Then
''           sArr(tmpI) = CStr(3) & Format(iInner, "00") & ":"
''         ElseIf tmpI < 65 Then
''           sArr(tmpI) = CStr(4) & Format(iInner, "00") & ":"
''         ElseIf tmpI < 81 Then
''           sArr(tmpI) = CStr(5) & Format(iInner, "00") & ":"
''         ElseIf tmpI < 97 Then
''           sArr(tmpI) = CStr(6) & Format(iInner, "00") & ":"
''         End If
''       Else
''          If tmpI < 17 Then
''            sArr(tmpI) = CStr(1) & Format(iInner, "00") & ":" & Mid(sTemp, 4)
''          ElseIf tmpI < 31 Then
''            sArr(tmpI) = CStr(2) & Format(iInner, "00") & ":" & Mid(sTemp, 4)
''          End If
''       End If
    Loop
   
  ElseIf InStr(1, sHeading, ":") = 4 Then   'Previous version file (with Page numbers)  OR  the very old format where there are 16 - 30 lines
    'First loop through file and check lines
    iCntRows = 1
    Do Until (EOF(FD) = True)
      Line Input #FD, sTemp
      If Trim(sTemp) <> "" Then iCntRows = iCntRows + 1
    Loop
    Close FD
    Open NewFile For Input As FD
    Line Input #FD, sHeading
    
    Line Input #FD, sTemp  'Read next line, if this line has a 3 didget number, it means we have to convert this as well...
    If iCntRows < 40 Or InStr(1, sTemp, ":") = 3 Or InStr(1, sTemp, ":") = 4 Then
      iMid = InStr(1, sTemp, ":") + 1
      FileToOpen = NewFile
      ConvertThisFile = True
      ReDim sArr(180)
      tmpI = 0
      'set the first line, since we just read this
      sNow = Format(Now, "YYYYMMDDHHmmSS")
      sArr(tmpI) = "000:" & sNow
      
      tmpI = tmpI + 1
      sArr(tmpI) = Format(tmpI, "000") & ":" & Mid(sTemp, iMid)
      'Read through file and load content into an array
      Do Until (EOF(FD) = True)
         Line Input #FD, sTemp
         If Trim(sTemp) <> "" Then
            tmpI = tmpI + 1
            sArr(tmpI) = Format(tmpI, "000") & ":" & Mid(sTemp, iMid)
         End If
      Loop
      If tmpI < 180 Then
         'Loop until we have loaded valid values all the way to end
         For tmpI = tmpI + 1 To 180
            sArr(tmpI) = Format(tmpI, "000") & ":"
         Next tmpI
      End If
    End If
  End If
  
  Close FD
  If ConvertThisFile Then 'Means old file, so convert this
    'Re-Create file with new structure
    GoSub RecreateFile
    ConvertThisFile = False
  End If
  
Next NewFile
 
Set newFolder = Nothing

  
Exit Sub
  
  
  
RecreateFile:
'Write new file from array
FD1 = FreeFile
Open FileToOpen For Output As FD1
For i = 0 To UBound(sArr)
  'If Trim(sArr(i)) <> "" Then
  Print #FD, sArr(i)
Next i
Close FD

Return

err1:
MsgBox "Converting old Playlist FAILED !!!" & Chr(13) & Chr(13) & "Please check the following file for flaws." & Chr(13) & Chr(13) & "***  " & FileToOpen & "  *** ", vbExclamation, "ERROR conversting Playlist"
WriteLog "modGlobal : ConvertOldPallets on file " & FileToOpen & " "
ConvertThisFile = False
Resume Next

End Sub

Public Sub SavePaleteLine(PaleteName As String, ButtonNo As Integer)
'To code later  ... :-)

End Sub

Public Sub SavePalete(pHeading As String, PageNo As Integer)
Dim iInner As Integer
Dim FD
Dim FileToOpen As String
Dim sHeading As String
Dim sStr As String
Dim sArr() As String
Dim sKeepHeading As String
Dim bHeadingFound As Boolean
Dim bEnd As Boolean
Dim bEmptyFile As Boolean
Dim sNow As String
Dim sTemp As String
Dim iP As Integer
Dim iValid As Integer
Dim bValid As Boolean
Dim iMax As Integer
Dim newFolder As Folder
Dim NewFile As file
Dim newFileName As String
Dim strPath As String
Dim bCreateNewFile As Boolean
Dim tmpIMaxBut As Integer
Dim tmpIbutStart As Integer
Dim tmpI As Integer
Dim iButIncrease As Integer

   

On Error Resume Next

'Set FSO = New FileSystemObject
bCreateNewFile = True
strPath = App.Path & "\Palets"
Set newFolder = FSO.GetFolder(strPath)

'Create the Palets Directory, if NOT exists
If Err.Number = 76 Then  'Dir does NOT exists
   MkDir App.Path & "\Palets"
   Err.Clear
   Set newFolder = FSO.GetFolder(strPath)
End If

'For Each NewFile In newFolder.Files
'   newFileName = NewFile
'  'pHeading
'   If UCase(Trim$(NewFile.ShortName)) = Trim(UCase(pHeading)) & ".DAT" Then
''  If UCase(Right$(Trim$(NewFile.ShortName), 3)) = "DAT" Then
'      bCreateNewFile = False
'  End If
'Next NewFile
 
sNow = Format(Now, "YYYYMMDDHHmmSS")
FD = FreeFile
iMax = 9
iInner = -1
tmpI = -1
bEmptyFile = False
ReDim Preserve sArr(181)  'Dimension to 97, so we can load all the songs, starting at 1

'We will save each list into its own file.

'----------------------------------------------------------------------------------------
' The first portion will load valid data into an array.
' We use this array to either change current entries, or
' to add new entries from the current screen'
'-----------------------------------------------------------

'If we use the clear button, we use the TMP001 file, so we kill it off before we start, so the rest of the code will create a clean file
If ClearPalet And Trim(UCase(pHeading)) = "TMP001" Then
  Kill strPath & "\" & Trim(UCase(pHeading)) & ".DAT"
End If


'Create empty file if nothing exists
FileToOpen = strPath & "\" & Trim(UCase(pHeading)) & ".DAT"
'---------------------------------------------
' FILE DOES NOT EXIST, CREATE NEW            '
' ALSO, create the array with empty entries  '
'---------------------------------------------
If Not FSO.FileExists(FileToOpen) Then
   Open FileToOpen For Output As FD
   'Close now without doing anything. This created an empty file and populated the array with default entries...
   Close FD
    For tmpI = 0 To 180
'      iInner = iInner + 1
'      If iInner > 16 Then iInner = 1
      If tmpI = 0 Then
         sArr(tmpI) = "000:" & sNow
      Else
         sArr(tmpI) = Format(tmpI, "000") & ":"
         'sArr(tmpI) = Format(iInner, "00") & ":"
'      ElseIf tmpI < 17 Then
'        sArr(tmpI) = CStr(1) & Format(iInner, "00") & ":"
'      ElseIf tmpI < 33 Then
'        sArr(tmpI) = CStr(2) & Format(iInner, "00") & ":"
'      ElseIf tmpI < 49 Then
'        sArr(tmpI) = CStr(3) & Format(iInner, "00") & ":"
'      ElseIf tmpI < 65 Then
'        sArr(tmpI) = CStr(4) & Format(iInner, "00") & ":"
'      ElseIf tmpI < 81 Then
'        sArr(tmpI) = CStr(5) & Format(iInner, "00") & ":"
'      ElseIf tmpI < 97 Then
'        sArr(tmpI) = CStr(6) & Format(iInner, "00") & ":"
      End If
    Next tmpI
Else
  '----------------------------------------------------------------------------
  'File DOES EXIST, so we check if we have to overwrite, from previous screen '
  '----------------------------------------------------------------------------
  'FileToOpen = strPath & "\" & Trim(UCase(pHeading)) & ".DAT"  'Set the file if nothing needs to be changed
'''  If bOverwritePalet Then
'''    If sOverwritePalet <> pHeading Then 'If we overwrite some other file
'''      FileToOpen = strPath & "\" & Trim(UCase(sOverwritePalet)) & ".DAT"
'''    Else  'Overwrite myself, so no need for the overwrite flag anymore
'''      bOverwritePalet = False
'''    End If
'''  End If
  
  Open FileToOpen For Input As FD
  'Read through file and load content into an array
  Do Until (EOF(FD) = True)
     Line Input #FD, sTemp
     tmpI = tmpI + 1
     If tmpI > 181 Then Exit Sub
'     iInner = iInner + 1
'     If iInner > 16 Then iInner = 1
     'Re-Create EMPTY entries if nothing in file
     If sTemp = "" Then
       If tmpI = 0 Then
         sArr(tmpI) = "000:" & sNow
       Else
         sArr(tmpI) = Format(tmpI, "000") & ":"
       End If
     Else
       sArr(tmpI) = sTemp
     End If
  Loop
  Close FD
  FileToOpen = strPath & "\" & Trim(UCase(pHeading)) & ".DAT" 'We need to make sure that we do save/overwrite the correct file
End If

   Select Case iMaxBut
      Case 9
         iButIncrease = 8
      Case 16
         iButIncrease = 15
      Case 20
         iButIncrease = 19
      Case 30
         iButIncrease = 29
   End Select
   
   Select Case PageNo
     Case 1 '1-16
       tmpIbutStart = 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
     Case 2 '17-32
       'tmpIbutStart = 17
       tmpIbutStart = iMaxBut + 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 32
     Case 3 '33-48
       tmpIbutStart = iMaxBut + iMaxBut + 1
       'tmpIbutStart = 33
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 48
     Case 4 '49-64
       'tmpIbutStart = 49
       tmpIbutStart = iMaxBut + iMaxBut + iMaxBut + 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 64
       
     Case 5 '65-80
       'tmpIbutStart = 65
       tmpIbutStart = iMaxBut + iMaxBut + iMaxBut + iMaxBut + 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 80
       
     Case 6 '81-96
       'tmpIbutStart = 81
       tmpIbutStart = iMaxBut + iMaxBut + iMaxBut + iMaxBut + iMaxBut + 1
       tmpIMaxBut = tmpIbutStart + iButIncrease  '16
       'tmpIMaxBut = 96
       
   End Select
   
tmpI = 0
      

'Now loop through Player form, and re-write the ARRAY with new content
sArr(0) = "000:" & sNow               'Check for first entry only, and write the CURRENT DateTime here
For i = tmpIbutStart To tmpIMaxBut  ' - 1
   If i = 0 Then
      sArr(i) = "000:" & sNow               'Check for first entry only, and write the CURRENT DateTime here
   Else
      tmpI = tmpI + 1
      If tmpI > tmpIMaxBut Then tmpI = 1
      If tmpI <= iMaxBut Then 'Maxbutton will be set according to the selecion on the player...
         If frmPlayer.sspSongTitle(tmpI).TagVariant <> "" Then
            sArr(i) = Format(GetNextButtonNumber + tmpI, "000") & ":" & AddBlank(frmPlayer.sspSongTitle(tmpI).TagVariant, 2) & "|" & frmPlayer.sspSongTitle(tmpI).Tag & "|" & frmPlayer.lblVol(tmpI).Caption & "|" & frmPlayer.cmdSong(tmpI).TagVariant & "|" & frmPlayer.sspProgress(tmpI).Tag
            'sArr(i) = CStr(PageNo) & Format(tmpI, "00") & ":" & AddBlank(frmPlayer.sspSongTitle(tmpI).TagVariant, 2) & "|" & frmPlayer.sspSongTitle(tmpI).Tag & "|" & frmPlayer.lblVol(tmpI).Caption & "|" & frmPlayer.cmdSong(tmpI).TagVariant & "|" & frmPlayer.sspProgress(tmpI).Tag
         Else
            sArr(i) = Format(GetNextButtonNumber + tmpI, "000") & ":"
            'sArr(i) = CStr(PageNo) & Format(tmpI, "00") & ":"
         End If
      Else
         'This will leave the rest of the data as original, since we ONLY change whats visible on screen
         If sArr(i) <> "" Then
          sArr(i) = Format(i, "000") & ":"
          'sArr(i) = CStr(PageNo) & Format(i, "00") & ":"
         End If
      End If
   End If
Next i

RecreateFile:
'Write new file from array
Open FileToOpen For Output As FD
For i = 0 To UBound(sArr) - 1
  'If Trim(sArr(i)) <> "" Then
  Print #FD, sArr(i)
Next i
Close FD


Set newFolder = Nothing
SaveSetting regMainKey, regSubKey, "Palette Name", pHeading

'ReloadPlayListArray strPath & "\" & Trim(UCase(pHeading)) & ".DAT"

frmPlayer.Caption = pHeading

End Sub

Public Sub ReloadPlayListArray(PaletteFile As String)
Dim FD
Dim sTemp As String
Dim iRow As Integer
Dim iPos As Integer
Dim iButNum As Integer
Dim sArr()
Dim sVol As Integer

FD = FreeFile
iRow = 0

  Open PaletteFile For Input As FD
  'Read through file and load content into an array
  Do Until (EOF(FD) = True)
      Line Input #FD, sTemp
      
      If Left(sTemp, 3) = "000" Then
        Line Input #FD, sTemp
      End If
      iRow = iRow + 1

      iPos = InStr(3, sTemp, "|")
      If iPos > 0 Then
         '==================================================================================
         'For Demo system, only allow load of 5 songs
         If DemoFlag Then
            DemoCnt = DemoCnt + 1
            If DemoCnt > DemoMax Then
               'GoTo ReadNext
               Exit Do
            End If
         End If
         '==================================================================================

          sArr = Split(sTemp, "|")
          'Check if file exists on the HD still...
          'Get the Title
          sArr(0) = Mid(sArr(0), 5)
          PlArr(PLA.efTtle, iRow) = sArr(0)               'Keep Full title here
          PlArr(PLA.eTtle, iRow) = FixSongTitle(CStr(sArr(0)))    'Fix the above to show nice title ('Determine if there are "-" in the title array, if so, split the 2)
          PlArr(PLA.eFN, iRow) = sArr(1)                  'Keep the filename here
          sVol = 70                                       'Set Default value for Volume = 70%
          If Val(sArr(2)) > 0 Then sVol = Val(sArr(2))    'Get the volume
          If sVol > 100 Then sVol = 99.999
          PlArr(PLA.eVol, iRow) = sVol
          PlArr(PLA.eAve, iRow) = sArr(3)                 'Use this to keep the Average when song is loaded...
          PlArr(PLA.eClr, iRow) = Val(sArr(4))            'Color

          Call BASS_StreamFree(chan(3))        'free the old stream
          chan(3) = BASS_StreamCreateFile(BASSFALSE, StrPtr(PlArr(PLA.eFN, iRow)), 0, 0, 0)
    
          Dim Bytes As Long
          Bytes = BASS_ChannelGetLength(chan(3), BASS_POS_BYTE)
          Dim time As Long
          time = BASS_ChannelBytes2Seconds(chan(3), Bytes)
         
          'Get the starting position by finding the silence in the front and set start after silence
          PlArr(PLA.eLs, iRow) = CStr(ScanForLeadingSilences(CStr(PlArr(PLA.eFN, iRow)), iRow))
          PlArr(PLA.eTp, iRow) = Trim(Format((time \ 60), "00") & ":" & Format(time Mod 60, "00"))
      End If
  Loop


End Sub

Public Function FixSongTitle(sData As String) As String
Dim sTitle() As String

Dim sTemp As String
Dim MaxChars As Integer

Const sNameChrs As String = "**"

Select Case iMaxBut
   Case 9, 16
      'iChr13 = 2
      MaxChars = 40
   Case 20
      'iChr13 = 1
      MaxChars = 30
   Case 30
      'iChr13 = 1
      MaxChars = 28

End Select

sTemp = sData

If InStr(1, sTemp, "-") = 0 Then
   If InStr(1, sTemp, "_") > 0 Then
      sTemp = Replace(sTemp, "_", "-")
   End If
End If
sTemp = Replace(sTemp, "/", " / ")
sTitle = Split(Replace(sTemp, "&", "&&"), "-")



'Select Case UBound(sTitle)
'   Case 1
'      If iMaxBut = 30 And Len(Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs) > 70 Then
'         FixSongTitle = Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'      Else
'         FixSongTitle = String(iChr13, Chr(13)) & Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'      End If
'   Case 2
'      FixSongTitle = String(iChr13, Chr(13)) & Trim(sTitle(1)) & " " & Trim(sTitle(2)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'   Case 3
'      FixSongTitle = String(1, Chr(13)) & Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "]" & Chr(13) & Trim(sTitle(2)) & Chr(13) & Trim(sTitle(3))
'   Case 4
'      FixSongTitle = sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'   Case Else
'      If iMaxBut = 9 Or iMaxBut = 25 Then
'         FixSongTitle = String(iChr13 + 1, Chr(13)) & Trim(sTitle(0))
'      Else
'         FixSongTitle = String(iChr13, Chr(13)) & Trim(sTitle(0))
'      End If
'End Select

Select Case UBound(sTitle)
   Case Is > 1
      For i = 1 To UBound(sTitle)
         FixSongTitle = FixSongTitle & Trim(sTitle(i)) & "-"
      Next i
      If Right(FixSongTitle, 1) = "-" Then FixSongTitle = Mid(FixSongTitle, 1, Len(FixSongTitle) - 1)
      FixSongTitle = FixSongTitle & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
   Case 1
      If Len(Trim(sTitle(1))) > MaxChars Then sTitle(1) = FixTitle(sTitle(1), MaxChars)
      If Len(Trim(sTitle(0))) > MaxChars Then sTitle(0) = FixTitle(sTitle(0), MaxChars)
      
      FixSongTitle = Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
   Case Else
      If Len(Trim(sTitle(0))) > MaxChars Then sTitle(0) = FixTitle(sTitle(0), MaxChars)
      FixSongTitle = sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
   
End Select
'If Trim(sTitle(1)) <> "" And Trim(sTitle(0)) <> "" Then  'Both present
'   FixSongTitle = Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'ElseIf Trim(sTitle(1)) = "" Then
'   FixSongTitle = sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'Else
'   FixSongTitle = Trim(sTitle(1)) & Chr(13) & sNameChrs & "  " & Trim(sTitle(0)) & "  " & sNameChrs
'End If


'Debug.Print "sData : " & sData & "  (" & Len(sData) & ")" & Chr(13) & "sTitle(0) : " & sTitle(0) & "  (" & Len(sTitle(0)) & ")" & Chr(13) & "sTitle(1) : " & sTitle(1) & "  (" & Len(sTitle(1)) & ")" & Chr(13) & "FIXED : " & FixSongTitle & "  (" & Len(FixSongTitle) & ")"


End Function

Function FixTitle(sTitle As String, MaxChars As Integer) As String
Dim iChr13 As Integer
Dim iCnt As Integer

   For iCnt = MaxChars To 1 Step -1
      If Mid(sTitle, iCnt, 1) = " " Then
         iChr13 = iCnt  'Get last position before end where we have a space so I can wrap it there
         Exit For
      End If
   Next iCnt
   FixTitle = Left(sTitle, iChr13) & Chr(13) & Mid(sTitle, iChr13 + 1)
      
End Function

Public Function GetNextButtonNumber() As Integer

If iButtonMaxSelected = 2 Then   '16 per page
   Select Case iPageno
      Case 1
         GetNextButtonNumber = 0
      Case 2
         GetNextButtonNumber = 16
      Case 3
         GetNextButtonNumber = 32
      Case 4
         GetNextButtonNumber = 48
      Case 5
         GetNextButtonNumber = 64
      Case 6
         GetNextButtonNumber = 80
   End Select
ElseIf iButtonMaxSelected = 3 Then '20 per page
   Select Case iPageno
      Case 1
         GetNextButtonNumber = 0
      Case 2
         GetNextButtonNumber = 20
      Case 3
         GetNextButtonNumber = 40
      Case 4
         GetNextButtonNumber = 60
      Case 5
         GetNextButtonNumber = 80
      Case 6
         GetNextButtonNumber = 100
   End Select
ElseIf iButtonMaxSelected = 4 Then  '30 per page
   Select Case iPageno
      Case 1
         GetNextButtonNumber = 0
      Case 2
         GetNextButtonNumber = 30
      Case 3
         GetNextButtonNumber = 60
      Case 4
         GetNextButtonNumber = 90
      Case 5
         GetNextButtonNumber = 120
      Case 6
         GetNextButtonNumber = 150
   End Select
End If
      
End Function

Public Sub ChangeTBBack(TB As Object, PNewBack As Long, pType As enuTBType)
Dim lTBWnd      As Long

    Select Case pType
        
        Case enuTB_FLAT     'FLAT Button Style Toolbar
            DeleteObject SetClassLong(TB.hWnd, GCL_HBRBACKGROUND, PNewBack) 'Its Flat, Apply directly to TB Hwnd
        
        Case enuTB_STANDARD 'STANDARD Button Style Toolbar
''''            lTBWnd = FindWindowEx(TB.hwnd, 0, "msvb_lib_toolbar", vbNullString) 'Standard, find Hwnd first
''''            DeleteObject SetClassLong(lTBWnd, GCL_HBRBACKGROUND, PNewBack)      'Set new Back
    End Select
End Sub

'Sets row height of lv ListView. hgt in pixels
 Public Sub SetLVRowHeight(lv As ListView, ilImage As ImageList, ByVal hgt As Long)
    
    If hgt <= 0 Then Exit Sub
    Set lv.SmallIcons = Nothing
    ilImage.ListImages.Clear
    ilImage.ImageHeight = hgt
'    ilImage.ListImages.Add , ,  Me.Icon
    Set lv.SmallIcons = ilImage
    
 End Sub
 

Public Sub SetLVSubImages(lv As ListView, ByVal Index, ByVal Column As Long, ByVal Image As Long, ByVal SubImagesOn As Boolean)
    
    Dim lvStyle As Long
    Dim LVItem As LV_Item
    
    lvStyle = SendMessageLong(lv.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    
    If SubImagesOn Then
        lvStyle = lvStyle Or LVS_EX_SUBITEMIMAGES
    Else
        lvStyle = lvStyle And Not LVS_EX_SUBITEMIMAGES
    End If
    
    Call SendMessageLong(lv.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lvStyle)
    
    If SubImagesOn Then
        With LVItem
            .mask = LVIF_IMAGE
            .iSubItem = Column
            .iItem = Index - 1
            .iImage = Image
        End With
        
        Call SendMessage(lv.hWnd, LVM_SETITEM, Index - 1, LVItem)
    End If
    
   ' LV.Refresh
    
End Sub



Public Function GetSysTime(ByVal bReturnUTC As Boolean) As String
   Dim oSYS    As SYSTEMTIME
   
   If bReturnUTC Then
      GetSystemTime oSYS
   Else
  '    GetLocalTime oSYS
   End If
   
   With oSYS
'''      GetSysTime = DateSerial(.wYear, .wMonth, .wDay) + _
'''            TimeSerial(.wHour, .wMinute, .wSecond) + _
'''            .wMilliseconds / 86400000
      GetSysTime = .wHour & ":" & .wMinute & ":" & .wSecond & "." & .wMilliseconds  '/ 86400000
   End With
   
End Function
 
Public Function FormatDateMS(ByVal dat As Date) As String
   Dim dDate As Date
   Dim dMilli As Double
   Dim lMilli As Long
   
   dDate = DateSerial(Year(dat), Month(dat), Day(dat)) + TimeSerial(hour(dat), Minute(dat), Second(dat))
   dMilli = dat - dDate
   lMilli = dMilli * 86400000
   If lMilli < 0 Then
      lMilli = lMilli + 1000 'was rounded so add 1 second
      dDate = DateAdd("s", -1, dDate) 'was rounded so subtract 1 second
   End If
   FormatDateMS = Format(dDate, "YYYY-MM-DD HH:NN:SS") & Format(lMilli / 1000, ".000")
   
End Function

Public Sub LoadDataIntoFile(DataName As Integer, FileName As String)
    Dim myArray() As Byte
    Dim myFile As Long
    On Error Resume Next
    Kill FileName
    If Dir(FileName) = "" Then
        myArray = LoadResData(DataName, "CUSTOM")
        myFile = FreeFile
        Open FileName For Binary Access Write As #myFile
         Put #myFile, , myArray
         Close #myFile
     End If
     
 End Sub



''''Public Sub InitTheme(ByVal hWnd As Long)
''''    Dim hTheme As Long
''''    Dim lPtrColorName As Long
''''    Dim lPtrThemeFile As Long
''''    Dim sThemeFile As String
''''    Dim sColorName As String
''''    Dim sShellStyle As String
''''    Dim hRes As Long
''''    Dim iPos As Long
''''    Dim lhWndD As Long
''''    Dim lhDCC As Long
''''    Dim lBitsPixel As Long
''''
''''    If (IsXp) Then
''''        On Error Resume Next
''''        hTheme = OpenThemeData(hWnd, StrPtr("ExplorerBar"))
''''        If Not (hTheme = 0) Then
''''
''''            ReDim bThemeFile(0 To 260 * 2) As Byte
''''            lPtrThemeFile = VarPtr(bThemeFile(0))
''''            ReDim bColorName(0 To 260 * 2) As Byte
''''            lPtrColorName = VarPtr(bColorName(0))
''''            hRes = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)
''''
''''            sThemeFile = bThemeFile
''''            iPos = InStr(sThemeFile, vbNullChar)
''''            If (iPos > 1) Then sThemeFile = Left(sThemeFile, iPos - 1)
''''            sColorName = bColorName
''''            iPos = InStr(sColorName, vbNullChar)
''''            If (iPos > 1) Then sColorName = Left(sColorName, iPos - 1)
''''
''''            Select Case sColorName
''''                Case "NormalColor"
''''                    m_iTheme = 1
''''                Case "Metallic"
''''                    m_iTheme = 2
''''                Case "Homestead"
''''                    m_iTheme = 3
''''                Case Else
''''                    m_iTheme = 0
''''            End Select
''''
''''            CloseThemeData hTheme
''''        End If
''''    End If
''''
''''End Sub




Public Function GetWindowStyle(ByVal hWnd As Long, ByVal extended_style As Boolean) As Long
'Public Function SetWindowStyle(ByVal hwnd As Long, ByVal extended_style As Boolean, ByVal style_value As Long, ByVal new_value As Boolean, ByVal brefresh As Boolean)
   Dim style_type As Long
   Dim style As Long
   
   If extended_style Then
       style_type = GWL_EXSTYLE
   Else
       style_type = GWL_STYLE
   End If
   
   ' Get the current style.
   GetWindowStyle = GetWindowLong(hWnd, style_type)
   
'   ' Add or remove the indicated value.
'   If new_value Then
'       style = style Or style_value
'   Else
'       style = style And Not style_value
'   End If
   
'   ' Hide Window if Changing ShowInTaskBar
'   If brefresh Then
'       ShowWindow hwnd, SW_HIDE
'   End If
   
'   ' Set the style.
'   SetWindowLong hwnd, style_type, style
'
'   ' Show Window if Changing ShowInTaskBar
'   If brefresh Then
'       ShowWindow hwnd, SW_SHOW
'   End If
'
'   ' Make the window redraw.
'   SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
   
    
End Function

Sub ctrfrm(Fm As Form)
'A procedure to

'Programer     Date    Description
'========== ========== =====================================================================
'Conversion 07/04/2000 Converted form 16 bit to 32 bit.

  If frmPlayer.WindowState <> 0 Then
    Fm.Left = (Screen.Width - Fm.Width) \ 2
    Fm.Top = (Screen.Height - Fm.Height) \ 2
  Else
    Fm.Left = frmPlayer.Left + (frmPlayer.Width - Fm.Width) \ 2
    If Fm.Left < 0 Then Fm.Left = 0
    Fm.Top = frmPlayer.Top + (frmPlayer.Height - Fm.Height) \ 2
    If Fm.Top < 0 Then Fm.Top = 0
  End If

End Sub

Sub CtrFrm2(Fm As Form)
'A procedure to

'Programer     Date    Description
'========== ========== =====================================================================
'Conversion 07/04/2000 Converted form 16 bit to 32 bit.

  Dim TTop As Integer
  
  Fm.Left = (Screen.Width - Fm.Width) / 2
  TTop = ((Screen.Height - Fm.Height) / 2) + 200
  If Fm.Left < 0 Then Fm.Left = 0
  If Fm.Top < 0 Then Fm.Top = 0
    Fm.Top = TTop + 300

End Sub

'''Public Function GetPeakLevel(ByVal sFile As String, Index As Integer) As Single
'''
''''   Dim buf(10000) As Byte
''''   Dim Count As Long
'''   Dim chan As Long
'''   Dim cPeak As Single
'''   Dim cLevel As Single
'''   Dim MonoLevel As Long
'''   Dim iCnt As Integer
'''   Dim aPeak() As Double
'''   Dim il As Long
'''   Dim lDuration As Long
'''   Dim lStreamHandle As Long
'''
'''   cPeak = 0
'''   'cLevel = 0
''''   ReDim aPeak(1)
'''
'''   chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, BASS_STREAM_DECODE)     ' create decoding channel
'''   If chan = 0 Then
'''      GetPeakLevel = -1
'''      Exit Function
'''   End If
'''   lDuration = Format(bassTime.GetDuration(chan), "0")
'''   SetLoadIndicator Index, 1, lDuration
'''
'''   Sleep 100
''''   Do While (BASS_ChannelIsActive(chan))
''''      'MonoLevel = BASS_ChannelGetLevelEx(chan, cLevel, 1, BASS_LEVEL_RMS)
''''      MonoLevel = BASS_ChannelGetLevelEx(chan, cLevel, 0.5, BASS_LEVEL_MONO)
''''      If cPeak < cLevel Then
''''         cPeak = cLevel ' found a higher peak BASS_LEVEL_RMS
''''      End If
''''      iCnt = iCnt + 1
'''''      ReDim Preserve aPeak(iCnt + 1)
'''''      aPeak(iCnt) = cLevel
''''      If iCnt Mod 100 = 0 Then
''''         SetLoadIndicator Index, iCnt, lDuration
''''      End If
''''   Loop
'''
'''   'Round lDuration to nearest 100
'''   lDuration = Round(lDuration / 100, 0) * 100
'''
'''   For iCnt = 1 To lDuration / 10
'''      SetLoadIndicator Index, iCnt, lDuration / 10
'''      Sleep 10
'''   Next iCnt
'''
'''
''''   SetLoadIndicator Index, iCnt, lDuration
''''   DoEvents
'''
'''   Call BASS_StreamFree(chan)
'''
'''   cPeak = 100
'''
'''   GetPeakLevel = cPeak
'''   DoEvents
'''
'''  'frmPlayer.picProgress(Index).Width = 0
'''  'frmPlayer.picProgress(Index).Visible = False
'''  'frmPlayer.sspProgress(Index).FloodPercent = 0
'''  'frmPlayer.sspProgress(Index).ProValue = 0
'''  frmPlayer.sspProgress(Index).value = 0
'''
'''  DoEvents
'''
'''
'''End Function

Public Function GetSongVolume(Index As Integer) As Single
  
  GetSongVolume = BASS_GetVolume()
  DoEvents
  
  frmPlayer.sspProgress(Index).value = 0
  DoEvents
    
End Function

Public Sub SetSongVolume(iVol As Single)

  Dim bSetVol As Boolean
  bSetVol = BASS_SetVolume(iVol)
    
End Sub

Public Function GetAverageLevel(ByVal sFile As String, Index As Integer) As Single

'   Dim buf(10000) As Byte
'   Dim Count As Long
   Dim chan As Long
   Dim cPeak As Single
   Dim cLPeak As Single
   Dim cLevel As Single
   Dim MonoLevel As Boolean
   Dim iCnt As Integer
   Dim aPeak() As Double
   Dim il As Long
   Dim lDuration As Long
   Dim lStreamHandle As Long
   Dim cDblLevel As Double
   
   cPeak = 0
   ReDim aPeak(1)
      
   chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, BASS_STREAM_DECODE)     ' create decoding channel
   If chan = 0 Then
      GetAverageLevel = -1
      Exit Function
   End If
   lDuration = Format(bassTime.GetDuration(chan), "0")
          
   Do While (BASS_ChannelIsActive(chan)) And iCnt < lDuration
      'MonoLevel = BASS_ChannelGetLevelEx(chan, cLevel, 1, BASS_LEVEL_RMS)
      MonoLevel = BASS_ChannelGetLevelEx(chan, cLevel, 0.5, BASS_LEVEL_MONO)   'Every half second
     ' MonoLevel = BASS_ChannelGetLevelEx(chan, cLevel, 0.5, BASS_LEVEL_RMS)    ' BASS_LEVEL_MONO)
      If cPeak < cLevel Then cPeak = Round(cLevel, 5) ' found a higher peak BASS_LEVEL_RMS
      iCnt = iCnt + 1
      ReDim Preserve aPeak(iCnt + 1)
      aPeak(iCnt) = cLevel
      If iCnt Mod 10 = 0 Then
         SetLoadIndicator Index, iCnt, lDuration
      End If
   Loop
  'Debug.Print "Count of loops : " & iCnt
   SetLoadIndicator Index, iCnt, lDuration
   
   cPeak = 0
   il = 0
   For iCnt = 1 To UBound(aPeak) - 1
      
'      If aPeak(iCnt) > cPeak Then cPeak = aPeak(iCnt)
'      If aPeak(iCnt) > 0 Then
'        If cLPeak = 0 And aPeak(iCnt) > 0 Then cLPeak = aPeak(iCnt)
'        If cLPeak > aPeak(iCnt) Then cLPeak = aPeak(iCnt)
'      End If
      
      If Round(aPeak(iCnt), 5) > 0.00027 Then
         cPeak = cPeak + aPeak(iCnt)
         il = il + 1
      End If
   
   Next iCnt
   
 '  cDblLevel = cLPeak
   
   If cPeak > il Then
      cPeak = 1
   Else
      If il = 0 Then
         cPeak = 1
      Else
         cPeak = cPeak / il
      End If
   End If
   
   If cPeak > 1 Then cPeak = 1
   
   Call BASS_StreamFree(chan)
   
   GetAverageLevel = cPeak
   
   DoEvents
   
  'frmPlayer.picProgress(Index).Width = 0
  'frmPlayer.picProgress(Index).Visible = False
  'frmPlayer.sspProgress(Index).FloodPercent = 0
  'frmPlayer.sspProgress(Index).ProValue = 0
  frmPlayer.sspProgress(Index).value = 0
  DoEvents
  

End Function

Public Function ScanForLeadingSilences(ByVal sFile As String, Index As Integer) As String
   Dim decode As Long
   Dim cLevel As Single
   Dim MonoLevel As Boolean
   Dim iSeconds As Integer
   Dim Ti As Double
      
   decode = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, BASS_STREAM_DECODE)     ' create decoding channel
   If decode = 0 Then
      ScanForLeadingSilences = -1
      Exit Function
   End If

   Ti = 0
   cStartPos = 0
   
   Do While (BASS_ChannelIsActive(decode)) And iSeconds < 300   'Only first 30 seconds
      MonoLevel = BASS_ChannelGetLevelEx(decode, cLevel, 0.1, BASS_LEVEL_MONO)   'Every tenth of second
      Ti = Ti + 0.1
      If cStartPos = 0 Then
        If Round(CDbl(cLevel), 5) > 0.00027 Then
          If Ti > 1 Then
            'cStartPos = Round(Ti - 0.5, 1)
            cStartPos = Round(Ti - 0.25, 1)
          Else
            cStartPos = 0
          End If
          Exit Do
        End If
      End If
      iSeconds = iSeconds + 1
   Loop

   Call BASS_StreamFree(decode)
   
   ScanForLeadingSilences = cStartPos
   
   DoEvents
  
  
End Function

Public Function GetPeakLevel(ByVal sFile As String, Index As Integer) As Single

'   Dim buf(10000) As Byte
'   Dim Count As Long
   Dim chan As Long
   Dim cPeak As Single
   Dim cLevel As Single
   Dim MonoLevel As Boolean
   Dim iCnt As Integer
   Dim aPeak() As Double
   Dim il As Long
   Dim lDuration As Long
   Dim lStreamHandle As Long
   Dim cDblLevel As Double
   
   
   cPeak = 0
   ReDim aPeak(1)
      
   chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, BASS_STREAM_DECODE)     ' create decoding channel
   If chan = 0 Then
      GetPeakLevel = -1
      Exit Function
   End If
   lDuration = Format(bassTime.GetDuration(chan), "0")
   
     
   Do While (BASS_ChannelIsActive(chan))
      'MonoLevel = BASS_ChannelGetLevelEx(chan, cLevel, 1, BASS_LEVEL_RMS)
      MonoLevel = BASS_ChannelGetLevelEx(chan, cLevel, 0.01, BASS_LEVEL_MONO)
     ' MonoLevel = BASS_ChannelGetLevelEx(chan, cLevel, 0.5, BASS_LEVEL_RMS)    ' BASS_LEVEL_MONO)
      If cPeak < cLevel Then cPeak = Round(cLevel, 5) ' found a higher peak BASS_LEVEL_RMS
      iCnt = iCnt + 1
      ReDim Preserve aPeak(iCnt + 1)
      aPeak(iCnt) = cLevel
      If iCnt Mod 50 = 0 Then
         SetLoadIndicator Index, iCnt, lDuration
      End If
      
   Loop
   SetLoadIndicator Index, iCnt, lDuration
   
   cPeak = 0
   il = 0
   For iCnt = 1 To UBound(aPeak) - 1
      If aPeak(iCnt) > cPeak Then
        cPeak = aPeak(iCnt)
      End If
   Next iCnt
   
   If cPeak > 1 Then cPeak = 1
   
   Call BASS_StreamFree(chan)
   
   GetPeakLevel = cPeak
   
   DoEvents
   
  frmPlayer.sspProgress(Index).value = 0
  DoEvents
  

End Function



Sub SetLoadIndicator(Index As Integer, sPos As Integer, lDuration As Long)
Dim pos As Single
Dim NewPerc As Single
Dim NewStart As Integer

'TimeElapsedPerc = Val(pos) * 100 / Val(Duration(1))
'TimeLeft = (TimeElapsedPerc / 100) * MaxWidth
''Show progress bar value
'picProgress(Player1Index).Width = TimeLeft




'NewPerc = (sPos / MaxWidth) * 100  'Percentage of where I need to start
NewPerc = (sPos / lDuration) * 100  'Percentage of where I need to start
'NewStart = ((NewPerc * MaxWidth) / 100)

'If Duration(CLng(cmdSong(Index).Tag)) - NewStart > 10 Then
'  frmPlayer.picProgress(Index).Visible = True
'  frmPlayer.picProgress(Index).Width = NewStart
If NewPerc > 100 Then NewPerc = 100
  'frmPlayer.sspProgress(Index).ProValue = NewPerc
  frmPlayer.sspProgress(Index).value = NewPerc
  'frmPlayer.sspProgress(Index).FloodPercent = NewPerc
  DoEvents
  
'End If

End Sub

Public Function GetSilenceLength(ByVal file As String) As Long
    
   Dim buf(10000) As Byte
   Dim Count As Long
   Dim chan As Long
   
   chan = BASS_StreamCreateFile(BASSFALSE, file, 0, 0, BASS_STREAM_DECODE)  ' create decoding channel
   Do While (BASS_ChannelIsActive(chan))
       Dim a As Long, b As Long
       b = BASS_ChannelGetData(chan, buf(0), 10000) ' decode some data
       a = 0
       Do While ((a < b) And (buf(a) = 0))
           a = a + 1
       Loop
       Count = Count + a 'add number of silent bytes
       If (a < b) Then Exit Do    'sound has begun!
   Loop
   Call BASS_StreamFree(chan)
   GetSilenceLength = Count
    
End Function


Public Function GetId3Tags(sFileToLoad As String) As String
Dim bId3V1Found As Boolean
Dim bId3V2Found As Boolean
Dim bTitleArtistLoaded As Boolean
Dim sArtist As String
Dim sTitle As String
Dim sListName As String
Dim i As Integer
Dim iPos As Integer

'Set objTag = New ID3v23x.clsID3v2

Set m_cID3v1 = New cMP3ID3v1
Set m_cID3v2 = New cMP3ID3v2

bId3V1Found = False
bId3V2Found = False
bTitleArtistLoaded = False
' Get ID3v1 Tag Information:
With m_cID3v1
  .MP3File = sFileToLoad
  If .HasID3v1Tag Then
    bId3V1Found = True
    If Trim(.Artist) = "" Then
      If Trim(.Title) = "" Then
        bId3V1Found = False
      Else
        sListName = Trim(.Title)
        sTitle = .Title
      End If
    Else
      sListName = Trim(.Artist) & " - " & Trim(.Title)
      sArtist = .Artist
      sTitle = .Title
      bTitleArtistLoaded = True
    End If
  End If
End With

If Not bTitleArtistLoaded Then
  With m_cID3v2
    .MP3File = sFileToLoad
    If .HasID3v2Tag Then
      bId3V2Found = True
      If Trim(.Artist) = "" Then
        If Trim(.Title) = "" Then
          bId3V2Found = False
        Else
          sListName = Trim(.Title)
          sTitle = .Title
        End If
      Else
        sListName = Trim(.Artist) & " - " & Trim(.Title)
        sArtist = .Artist
        sTitle = .Title
      End If
    End If
  End With
End If

''''objTag.ReadTag sFileToLoad
''''
''''With objTag
'''''    txtTrack.Text = .GetFrameValue(eTrack)
''''    sTitle = .GetFrameValue(eTitle)
''''    sArtist = .GetFrameValue(eArtist)
''''    sArtist = .GetFrameValue(eArtist)
'''''    txtAlbum.Text = .GetFrameValue(eAlbum)
'''''    txtYear.Text = .GetFrameValue(eYear)
'''''    txtGenre.Text = .GetFrameValue(eGenre)
'''''    txtComments.Text = .GetFrameValue(eComment)
'''''    txtComposer.Text = .GetFrameValue(eComposer)
'''''    txtOrigArtist.Text = .GetFrameValue(eOrigArtist)
'''''    txtCopyright.Text = .GetFrameValue(eCopyright)
'''''    txtURL.Text = .GetFrameValue(eURL)
'''''    txtEncodedBy.Text = .GetFrameValue(eEncodedBy)
''''End With


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

End Function

Sub LoadMsgLits()
'  Ready = 0
'  Start = 1
'  VerifyLic = 2
'  ChkSndMods = 3
'  VerfyMusicLibs = 4
'  Init = 5
'  ChkDisk = 6
'  LoadEnv = 7
'  GetReg = 8
'  FormatLayout = 9
'  LoadSoundCards = 10
'  OpenPlayList = 11
'  Finalise = 12
'  Msg1 = 13 '"BackTrax Player Professional"
'  Msg2 = 14 '"BackTrax Player is already running."
'  Msg3 = 15 '"Copyright © # Lilac Productions. All Rights Reserved."
'  Msg4 = 16 '"Version"

ReDim LoadARR(25)
LoadARR(MsgLits.Start) = "Starting BackTrax Player Version"
LoadARR(MsgLits.VerifyLic) = "Verifying License..."
LoadARR(MsgLits.MainPlugin) = "Validating Main Plug-in..."
LoadARR(MsgLits.Otherplugins) = "Loading Music Libraries..."
LoadARR(MsgLits.Init) = "Prepare Main Interface..."
LoadARR(MsgLits.ChkDisk) = "Checking available disk space..."
LoadARR(MsgLits.LoadEnv) = "Loading Environment Variables..."
LoadARR(MsgLits.GetReg) = "Retreiving Registry entries..."
LoadARR(MsgLits.FormatLayout) = "Format layouts..."
LoadARR(MsgLits.InitSoundCard) = "Initializing Default Sound Card..."
LoadARR(MsgLits.LoadSoundCards) = "Loading Available Sound Cards ..."
LoadARR(MsgLits.OpenPlayList) = "Loading Last Opened Playlist..."
LoadARR(MsgLits.Finalise) = "Finalising..."
LoadARR(MsgLits.Ready) = "System Ready..."

LoadARR(MsgLits.ProdName) = "BackTrax Player Professional"
LoadARR(MsgLits.Running) = "BackTrax Player is already running."
LoadARR(MsgLits.Copyrite) = "Copyright © # Lilac Productions. All Rights Reserved."
LoadARR(MsgLits.Version) = "Version"



End Sub

Public Sub Main()
Dim sFilter1 As String
Dim errCnt As Integer
'================================================================
' Check the compile parameters to see which app to run          '
'================================================================
'''#If PlayerOnly = -1 Then  'COMPILE without the main form to show
'''  MainApp = False
'''#Else
'''  MainApp = True
'''#End If
On Error GoTo err1
errCnt = 0
bSkipValidation = False
Dim FSO As New FileSystemObject
WriteLog "MAIN : STARTING..."


''''BASS_SetConfig BASS_CONFIG_DEV_DEFAULT, 1
''''BASS_Init -1, 44100, 0, 0, 0



'Dimension the array, and set values in each (default will be 10, which is the middle of the sliders when we want to set EQ
ReDim aEQ(30, MaxEqs + 1) 'Second dimension = value, First dimension = Indicator (0 = index of song,2 = freq of band)
For iButCnt = 1 To 30         'We can only do 30 buttons
  For iEqBands = 1 To MaxEqs
    aEQ(iButCnt, iEqBands) = 0 'Slider value at default (middle is 10)
  Next iEqBands
Next iButCnt

'Load EQ frequencies
EqFreq(1) = 125   '80
EqFreq(2) = 1000  '125
EqFreq(3) = 8000  '170

EqFreq(4) = 250
EqFreq(5) = 310
EqFreq(6) = 450
EqFreq(7) = 630
EqFreq(8) = 1000
EqFreq(9) = 2500
EqFreq(10) = 4000
EqFreq(11) = 6300
EqFreq(12) = 10000
EqFreq(13) = 12500
EqFreq(14) = 16000



If FSO.FileExists(App.Path & "\Loading.log") Then Kill App.Path & "\Loading.log"

'Check for multiple monitors
MonCount = GetMonitorCount
'Load the literals
WriteLog "MAIN : Call LoadMsgLits"
LoadMsgLits
'Show the loader form so we have at least something to start with ??
frmLoader.Show
Sleep 50
DoEvents

frmLoader.ShowLoading MsgLits.Start, 150 'Validate serial
Sleep 500
DoEvents


frmLoader.ShowLoading MsgLits.VerifyLic, 150 'Validate serial
'First Determine if DEMO
If FSO.FileExists(App.Path & "\omed.txt") Then
   SaveSetting regMainKey, regSubKey, "SerialNumber", "DEMO"
   SaveSetting regMainKey, regSubKey, "DemoUsed", 0
   Kill App.Path & "\omed.txt"
End If
Set FSO = Nothing

'Get the setting from the registry - If above test was done, there should be a DEMO in the serial number
CheckSerial:
Screen.MousePointer = vbHourglass
DoEvents
WriteLog "MAIN : Get Serial number from Registry..."
SerialNumber = GetSetting(regMainKey, regSubKey, "SerialNumber")

WriteLog "MAIN : Check for DEMO..."
If SerialNumber = "DEMO" Then
   'Add 1 to count of times used
   iCntDemo = Val(GetSetting(regMainKey, regSubKey, "DemoUsed"))
   iCntDemo = iCntDemo + 1
   
   If iCntDemo > 5 Then
      MsgBox "        **********************************************   " & Chr(13) & "                             The DEMO option has EXPIRED " & Chr(13) & "        **********************************************   " & Chr(13) & Chr(13) & "Please obtain a valid SERIAL NUMBER to continue using the package.", vbInformation, "DEMO Expired"
      Screen.MousePointer = vbDefault
      frmSerial.Show vbModal
      Screen.MousePointer = vbHourglass
      If Not bSkipValidation Then
         GoTo CheckSerial
      Else
         End
      End If
   Else
      'SaveSetting regMainKey, regSubKey, "SerialNumber", ""
      SaveSetting regMainKey, regSubKey, "DemoUsed", iCntDemo
   End If
   DemoFlag = True
   
   ShowDemoMsg
Else
   If Not ValidateSerial(SerialNumber) Then
      frmSerial.Show vbModal
      If Not bSkipValidation Then
         GoTo CheckSerial
      Else
         End
      End If
   End If
End If
   
Screen.MousePointer = vbHourglass
DoEvents
If DemoFlag Then
   Sleep 500
Else
   frmLoader.sspDemo.Visible = False
   DoEvents
   Sleep 1500
End If
DoEvents

frmLoader.ShowLoading MsgLits.MainPlugin, 150 'Checking Add-on libs
Sleep 250
DoEvents
   
errCnt = errCnt + 1  '7

'Force a Change directory so we can find the BASS modules each time I load form...
WriteLog "MAIN : ChDir to APP.Path"
ChDrive App.Path
ChDir App.Path

errCnt = errCnt + 1 '8

' check the correct BASS was loaded
WriteLog "MAIN : Check BASS version..."
If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
    Call MsgBox("An incorrect version of BASS.DLL was loaded." & Chr(13) & Chr(13) & "HiWord(BASS_GetVersion) Version = " & HiWord(BASS_GetVersion) & Chr(13) & "BASSVERSION = " & BASSVERSION, vbCritical)
    End
End If


errCnt = errCnt + 1 '9
WriteLog "MAIN : Set bassTime = New cbass_time..."
Set bassTime = New cbass_time

Dim iPlugCnt As Integer
Dim sFilterPerGroup As String
Dim sPlugIns() As String
Dim sWork As String


frmLoader.ShowLoading MsgLits.Otherplugins, 150 'Checking Add-on libs
Sleep 150
DoEvents

WriteLog "MAIN : Load List of BASS DLL's..."
sWork = "bass.dll|"

sFilter1 = "*.mp3;*.mp2;*.mp1;*.wav;*.aif;*.aiff"
ListOfPlugins = "- bass.dll" & "  (" & Replace(Replace(sFilter1, "*.", ""), ";", " | ") & ")" & Chr(13)

errCnt = errCnt + 1 '10

' look for plugins (in the executable's directory)
Dim fh As String
fh = Dir("bass*.dll")   ' find 1st file

frmLoader.Label1.Visible = True

iPlugCnt = 0
WriteLog "MAIN : Look for plugins (in the executable's directory)..."

Do While (fh <> "")
   Dim plug As Long
   plug = BASS_PluginLoad(fh, 0)   ' plugin loaded...
   If (plug) Then
         
      Dim pinfo As BASS_PLUGININFO
      pinfo = BASS_PluginGetInfo(plug) ' get plugin info to add to the file selector filter...
      Dim a As Long
      sFilterPerGroup = ""
      For a = 0 To pinfo.formatc - 1
          sFilter1 = sFilter1 & ";" & VBStrFromAnsiPtr(BASS_PluginGetInfoFormat(plug, a).exts)
          sFilterPerGroup = sFilterPerGroup & " | " & Replace(Replace(VBStrFromAnsiPtr(BASS_PluginGetInfoFormat(plug, a).exts), "*.", ""), ";", " | ")
      Next a
      If Left(Trim(sFilterPerGroup), 1) = "|" Then
         sFilterPerGroup = Trim(sFilterPerGroup)
         sFilterPerGroup = Trim(Mid(sFilterPerGroup, 2))
      End If
      ListOfPlugins = ListOfPlugins & "- " & fh & "  (" & sFilterPerGroup & ")" & Chr(13)
      sWork = sWork & fh & "|"
      
'      If frmLoader.lstPlugins.Caption = "" Then
'         frmLoader.lstPlugins.Caption = fh
'      Else
'         'If iPlugCnt Mod 4 = 0 Then
'         '   frmLoader.lstPlugins.Caption = frmLoader.lstPlugins.Caption & Chr(13) & fh
'         'Else
'            frmLoader.lstPlugins.Caption = frmLoader.lstPlugins.Caption & "  |  " & fh
'         'End If
'      End If
'      iPlugCnt = iPlugCnt + 1
   End If
   fh = Dir()  ' get next file
Loop

sWork = Left(sWork, Len(sWork) - 1)
sPlugIns = Split(sWork, "|")
iPlugCnt = 0

WriteLog "MAIN : Write Plugins to screen..."
For i = 0 To UBound(sPlugIns)
'   If iPlugCnt > 2 Then
'      iPlugCnt = 0
'   End If
   frmLoader.lstPlugins(iPlugCnt).Caption = frmLoader.lstPlugins(iPlugCnt).Caption & IIf(sPlugIns(i) = "bass.dll", sPlugIns(i) & "(" & HiWord(BASS_GetVersion) & ")", sPlugIns(i)) & Chr(13)
'   iPlugCnt = iPlugCnt + 1
Next i


Sleep 250
DoEvents

errCnt = errCnt + 1 '11

Filter = LCase(sFilter1) 'Make sure we always have lCase, because we need to test exact mathes later

Sleep 500

errCnt = errCnt + 1  '12
WriteLog "MAIN : Load frmPlayer..."
Load frmPlayer
DoEvents
WriteLog "MAIN : SHOW frmPlayer..."
frmPlayer.Show
DoEvents
Sleep 1000
   
   

'================================
' Shared loads and tests...     '
'================================

frmLoader.ShowLoading MsgLits.Init, 150 'Initialise main interface
Sleep 250
DoEvents



'
'    ' check the correct BASS_FX was loaded
'    If (HiWord(BASS_FX_GetVersion) <> BASSVERSION) Then
'        Call MsgBox("An incorrect version of BASS_FX.DLL was loaded (2.4 is required)", vbCritical)
'        End
'    End If



frmLoader.ShowLoading MsgLits.Init, 150 'Initialise main interface
Sleep 500
DoEvents


'==================================================================================
' Condtional compile test, so we can run with both screens, or as a player only   '
'==================================================================================
''If MainApp Then
'   frmLoader.ShowLoading 4, 150
'   Sleep 250
   DoEvents
   
   errCnt = errCnt + 1  '12
'''   WriteLog "MAIN : Load frmPlayer..."
'''   Load frmPlayer
'''   DoEvents
   WriteLog "MAIN : SHOW frmPlayer..."
   frmPlayer.Show
   DoEvents
   Sleep 1000

'''Else
'''  frmSinglePlayer.Show
'''End If

'frmTestMp3Pics.Show
   Exit Sub

err1:
MsgBox "ERROR " & Chr(13) & Chr(13) & "The following error has occurred : " & Chr(13) & Chr(13) & "ERROR : " & Err.Description & " (" & Err.Number & ")", vbExclamation, "MAIN"
Resume
Err.Clear
End

End Sub

Sub ShowDemoMsg()

      frmLoader.sspDemo.Visible = True
      frmLoader.sspDemo.Caption = "DEMO MODE"
      
      If 5 - iCntDemo = 1 Then
         frmLoader.sspDemo.ForeColor = vbYellow
         'MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3 & Chr(13) & Chr(13) & "The DEMO system may only be used 1 more time.", vbExclamation, DemoHeading
         If MsgBox(vbTab & DemoMsg1 & Chr(13) & vbTab & "==============" & Chr(13) & Chr(13) & vbTab & "-   " & DemoMsg3 & Chr(13) & vbTab & "-   " & "The DEMO system may be used 1 more times." & Chr(13) & Chr(13) & vbTab & "Would you like to Register now??", vbQuestion + vbYesNo, DemoHeading) = vbYes Then
            GoTo EnterSerial
         End If
      ElseIf 5 - iCntDemo = 0 Then
         frmLoader.sspDemo.ForeColor = vbRed
         'MsgBox DemoMsg1 & Chr(13) & Chr(13) & DemoMsg3 & Chr(13) & Chr(13) & "The DEMO system may be used for the LAST time.", vbExclamation, DemoHeading
         If MsgBox(vbTab & DemoMsg1 & Chr(13) & vbTab & "==============" & Chr(13) & Chr(13) & vbTab & "-   " & DemoMsg3 & Chr(13) & vbTab & "-   " & "The DEMO system may be used for the LAST time." & Chr(13) & Chr(13) & vbTab & "Would you like to Register now??", vbQuestion + vbYesNo, DemoHeading) = vbYes Then
            GoTo EnterSerial
         End If
      Else
         frmLoader.sspDemo.ForeColor = vbGreen
         If MsgBox(vbTab & DemoMsg1 & Chr(13) & vbTab & "==============" & Chr(13) & Chr(13) & vbTab & "-   " & DemoMsg3 & Chr(13) & vbTab & "-   " & "The DEMO system may be used " & 5 - iCntDemo & " more times." & Chr(13) & Chr(13) & vbTab & "Would you like to Register now??", vbQuestion + vbYesNo, DemoHeading) = vbYes Then
            GoTo EnterSerial
         End If
      End If
   
   On Error Resume Next
   
   Exit Sub
   
EnterSerial:
   frmSerial.Show vbModal
   
End Sub

Public Function ConvertTwipsToPixels(lngTwips As Long, bytDirection As Byte) As Long
   
   Dim lngRetVal As Long
   
   If (bytDirection = 0) Then       'Horizontal
      lngRetVal = lngTwips / Screen.TwipsPerPixelX
   Else                            'Vertical
      lngRetVal = lngTwips / Screen.TwipsPerPixelY
   End If
   ConvertTwipsToPixels = lngRetVal

End Function
    
Public Function ConvertPixelsToTwips(lngPixels As Long, bytDirection As Byte) As Long

   Dim lngRetVal As Long
   
   If (bytDirection = 0) Then       'Horizontal
      lngRetVal = lngPixels * Screen.TwipsPerPixelX
   Else                            'Vertical
      lngRetVal = lngPixels * Screen.TwipsPerPixelY
   End If
   ConvertPixelsToTwips = lngRetVal

End Function


Public Sub MakeFormRound(pform As Form, lValue As Long)
'Esme - About Form
 
Dim lret As Long
Dim l As Long
Dim llWidth As Long
Dim llHeight As Long
 
'Get Form size in pixels
llWidth = pform.Width / Screen.TwipsPerPixelX
llHeight = pform.Height / Screen.TwipsPerPixelY

'Create Form with Rounded Corners
lret = CreateRoundRectRgn(0, 0, llWidth, llHeight, lValue, lValue)
l = SetWindowRgn(pform.hWnd, lret, True)
 
End Sub

Public Sub LVDragDropSingle(ByRef lvList As ListView, ByVal X As Single, ByVal Y As Single)
'Item being dropped
Dim objDrag As ListItem
'Item being dropped on
Dim objDrop As ListItem
'Item being readded to the list
Dim objNew As ListItem
'Subitem reference in dropped item
Dim objSub As ListSubItem
'Drop position
Dim intIndex As Integer
Dim intRememberSort As Integer

'Retrieve the original items
Set objDrop = lvList.HitTest(X, Y)
Set objDrag = lvList.SelectedItem
If (objDrop Is Nothing) Or (objDrag Is Nothing) Then
    Set lvList.DropHighlight = Nothing
    Set objDrop = Nothing
    Set objDrag = Nothing
    Exit Sub
End If


'Retrieve the drop position
intIndex = objDrop.Index
'intRememberSort = lvList.ListItems(intIndex).SubItems(4)

'Remove the dragged item
lvList.ListItems.Remove objDrag.Index

'Add it back into the dropped position
Set objNew = lvList.ListItems.Add(intIndex, objDrag.key, objDrag.text, objDrag.Icon, objDrag.SmallIcon)
'Copy the original subitems to the new item
If objDrag.ListSubItems.Count > 0 Then
    For Each objSub In objDrag.ListSubItems
      If objSub.Index = 4 Then
        objNew.ListSubItems.Add objSub.Index, objSub.key, "99999", objSub.ReportIcon, objSub.ToolTipText
      Else
        objNew.ListSubItems.Add objSub.Index, objSub.key, objSub.text, objSub.ReportIcon, objSub.ToolTipText
      End If
    Next
End If
Dim iRows As Long

For iRows = 1 To lvList.ListItems.Count
    lvList.ListItems.Item(iRows).SubItems(4) = Format(iRows, "000")
Next iRows


'Reselect the item
objNew.Selected = True

'Destroy all objects
Set objNew = Nothing
Set objDrag = Nothing
Set objDrop = Nothing
Set lvList.DropHighlight = Nothing


End Sub

Public Function Change_pb_ForeColor(ByVal hWnd As Long, ByVal lColor As Long)
  SendMessage hWnd, PBM_SETBARCOLOR, 0, ByVal lColor
End Function

Public Function Change_pb_Color(ByVal hWnd As Long, ByVal lColor As Long)
  SendMessage hWnd, PBM_SETBKCOLOR, 0, ByVal lColor
End Function

Public Function RandomNumber(ByVal MaxValue As Integer) As Integer
'Dim iNew As Integer
' Dim intResult As Integer
'  On Error Resume Next
'
'  ReDim Preserve aRnd(MaxValue + 1)
'Redo:
'  Randomize
'  intResult = Int(MaxValue * Rnd) + 1
'
'  For iNew = 1 To MaxValue
'    If aRnd(iNew) = intResult Then
'      GoTo Redo
'      Exit For
'    End If
'    If aRnd(iNew) = 0 Then
'      aRnd(iNew) = intResult
'      Exit For
'    End If
'  Next iNew
'
'
'     '// Initializes the random-number generator, otherwise each time you run your
'     '// program, the sequence of numbers will be the same
'     Randomize
'     intResult = Int((MaxValue * Rnd) + 1) '// Generate random value between 1 and 6.
'     MsgBox "Number: " & intResult '// Display result
     
  

End Function

Public Sub Sort2Array(ByRef DArray(), Element As Integer)
    Dim gap As Integer, doneflag As Integer, SwapArray()
    Dim Index As Integer, acol As Integer, CNT As Integer
    ReDim SwapArray(2, UBound(DArray, 1), UBound(DArray, 2))
    'Gap is half the records
    gap = Int(UBound(DArray, 2) / 2)
    Do While gap >= 1
        Do
            doneflag = 1
            For Index = 0 To (UBound(DArray, 2) - (gap + 1))
                'Compare 1st 1/2 to 2nd 1/2
                If DArray(Element, Index) > DArray(Element, (Index + gap)) Then
                    For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
                        SwapArray(0, acol, Index) = DArray(acol, Index)
                        SwapArray(1, acol, Index) = DArray(acol, Index + gap)
                    Next
                    For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
                        DArray(acol, Index) = SwapArray(1, acol, Index)
                        DArray(acol, Index + gap) = SwapArray(0, acol, Index)
                    Next
                    CNT = CNT + 1
                    doneflag = 0
                End If
            Next
        Loop Until doneflag = 1
        gap = Int(gap / 2)
    Loop
    
End Sub

Public Sub SortArray(ByRef aArr() As String)
Dim i As Long
Dim j As Long
Dim minimum As Long
Dim swapValue As Long
Dim upperBound As Long
Dim lowerBound As Long

lowerBound = LBound(aArr)
upperBound = UBound(aArr)
For i = lowerBound To upperBound
  minimum = i
  For j = i + 1 To upperBound
    'Search for the smallest remaining item in the array
    If aArr(j) < aArr(minimum) Then
      'A smaller value has been found, remember the position in the array
      minimum = j
    End If
  Next j
  If minimum <> i Then
    'Swap array Values
    swapValue = aArr(minimum)
    aArr(minimum) = aArr(i)
    aArr(i) = swapValue
  End If
Next i

End Sub

Public Function GetTag(FileName As String, ITag As Mp3IDTag) As Boolean
Dim f As Integer
Dim Tagg As String * 3
Dim temp As String
Dim i As Integer
Dim CMD As String

f = FreeFile
On Error Resume Next
Open FileName For Binary As #f
Get #f, FileLen(FileName) - 127, Tagg
If Tagg = "TAG" Then
    GetTag = True
    Get #f, , ITag.Songname
    Get #f, , ITag.Artist
    Get #f, , ITag.Album
    Get #f, , ITag.Year
    Get #f, , ITag.Comment
    Get #f, , ITag.Genre
    ITag.Songname = Replace(ITag.Songname, Chr(0), " ", , , vbBinaryCompare)
    ITag.Artist = Replace(ITag.Artist, Chr(0), " ", , , vbBinaryCompare)
    ITag.Album = Replace(ITag.Album, Chr(0), " ", , , vbBinaryCompare)
    ITag.Year = Replace(ITag.Year, Chr(0), " ", , , vbBinaryCompare)
    ITag.Comment = Replace(ITag.Comment, Chr(0), " ", , , vbBinaryCompare)
    ITag.Genre = Replace(ITag.Genre, Chr(0), " ", , , vbBinaryCompare)
Else
    GetTag = False
    temp = FileName
    If InStr(temp, "\") Then
        Do
        temp = Right(temp, Len(temp) - 1)
        Loop Until InStr(temp, "\") = 0
        
        'If UCase(Right(temp, 4)) = ".MP3" Then
        temp = Left(temp, Len(temp) - 4)
        
        If InStr(temp, "-") Then
            ITag.Artist = Left(temp, InStr(temp, "-") - 1)
            ITag.Songname = Mid(temp, InStr(temp, "-") + 1)   'Right(temp, Len(temp) - InStr(temp, "-") - 1)
        Else
            ITag.Songname = temp
        End If
    End If
End If

If Trim(ITag.Songname) = "" Then
    GetTag = False
    temp = FileName
    If InStr(temp, "\") Then
        Do
        temp = Right(temp, Len(temp) - 1)
        Loop Until InStr(temp, "\") = 0
        
        'If UCase(Right(temp, 4)) = ".MP3" Then
        temp = Left(temp, Len(temp) - 4)
        
        If InStr(temp, "-") Then
            ITag.Artist = Left(temp, InStr(temp, "-") - 1)
            ITag.Songname = Mid(temp, InStr(temp, "-") + 1)   'Right(temp, Len(temp) - InStr(temp, "-") - 1)
        Else
            ITag.Songname = temp
        End If
    End If

End If

If Trim(ITag.Songname) > 38 Then ITag.Songname = Left(ITag.Songname, 36) & "..."
If Trim(ITag.Artist) > 38 Then ITag.Artist = Left(ITag.Artist, 36) & "..."

Close #f
End Function

Public Function SetItemFocusA(ByRef ctlListview As MSComctlLib.ListView, ByVal iIndex As Long, Optional iVisibleIndex = 2) As Boolean
On Error GoTo Hell

Dim lv As LV_Item
Dim lvItemsPerPage As Long
Dim lvNeededItems As Long
Dim lvCurrentTopIndex As Long

    With ctlListview
        ' Since this is a multi-select list, we want to unselect all items before selecting the current track.
        With lv
            .mask = LVIF_STATE
            .State = False
            .stateMask = LVIS_SELECTED
        End With
        Call SendMessage(.hWnd, LVM_SETITEMSTATE, -1, lv) ' Poof
        
        ' Select and set the focus rectangle on the item.
        With lv
            .mask = LVIF_STATE
            .State = True
            .stateMask = LVIS_SELECTED Or LVIS_FOCUSED
        End With
        Call SendMessage(.hWnd, LVM_SETITEMSTATE, iIndex - 1, lv) ' Listview index is 0-based in the API world
        
        ' Determine if desired index + number of items in view will exceed total items in the control
        lvCurrentTopIndex = SendMessage(.hWnd, LVM_GETTOPINDEX, 0&, ByVal 0&)
        lvItemsPerPage = SendMessage(.hWnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)
        
        ' Do we even need to scroll? Not if the selected track is already in view
        'If (lvCurrentTopIndex >= iIndex) Or (iIndex > lvCurrentTopIndex + lvItemsPerPage) Then
        
            ' Is 'x' above or below target index?
            If lvCurrentTopIndex >= iIndex Then ' Going UP
                If iIndex > iVisibleIndex Then
                    .ListItems((iIndex - iVisibleIndex + 1)).EnsureVisible ' Drops the highlighted item down a few so it's not hidden
                                                            ' behind the Column header.
                Else
                    .ListItems((iIndex)).EnsureVisible
                End If
            
            Else ' Going DOWN
                ' Are there sufficient items to set to the topindex
                If (iIndex + lvItemsPerPage) > .ListItems.Count Then
               
                   ' Can't be set to the top as the control has insufficient
                   ' items, so just scroll to the end of listview
                   .ListItems(.ListItems.Count).EnsureVisible
                   
                Else
                
                  ' It is below, and since a listview always moves the item just into view,
                  ' have it instead move to the top by faking item we want to 'EnsureVisible'
                  ' the item lvItemsPerPage -1(or -3) below the actual index of interest.
                    If iIndex > iVisibleIndex Then
                        .ListItems((iIndex + lvItemsPerPage) - iVisibleIndex).EnsureVisible
                    Else
                        .ListItems((iIndex + lvItemsPerPage) - 1).EnsureVisible
                    End If
                End If
            End If
        'End If
    End With

    SetItemFocusA = True

Hell:
End Function


Public Sub ListViewMoveToTop(ByVal lv As ListView)
    Dim bWasUnSel As Boolean
    Dim tmpLvItem As ListItem
    Dim newLvItem As ListItem
    Dim tmpSubItem As ListSubItem
    Dim iIcon As Integer
    Dim iSmallIcon As Integer
    Dim i As Integer
    bWasUnSel = False
    For i = 1 To lv.ListItems.Count
        Set tmpLvItem = lv.ListItems(i)
        iIcon = lv.ListItems(i).Icon
        iSmallIcon = lv.ListItems(i).SmallIcon
        If tmpLvItem.Selected Then
            If bWasUnSel Then
                Set newLvItem = lv.ListItems.Add(1, , tmpLvItem.text, iIcon, iSmallIcon)
                newLvItem.Tag = tmpLvItem.Tag
                newLvItem.Checked = tmpLvItem.Checked
                newLvItem.key = tmpLvItem.key
                For Each tmpSubItem In tmpLvItem.ListSubItems
                    newLvItem.SubItems(tmpSubItem.Index) = tmpSubItem.text
                Next
                lv.ListItems.Remove (tmpLvItem.Index)
                newLvItem.Selected = True
                Set newLvItem = Nothing
            End If
        Else
            bWasUnSel = True
        End If
        Set tmpLvItem = Nothing
    Next
End Sub



Public Sub ListViewMovetoBottom(ByVal lv As ListView)
    Dim bWasUnSel As Boolean
    Dim tmpLvItem As ListItem
    Dim newLvItem As ListItem
    Dim tmpSubItem As ListSubItem
    Dim iIcon As Integer
    Dim iSmallIcon As Integer
    Dim i As Integer
    bWasUnSel = False
    For i = lv.ListItems.Count To 1 Step -1
        Set tmpLvItem = lv.ListItems(i)
        iIcon = lv.ListItems(i).Icon
        iSmallIcon = lv.ListItems(i).SmallIcon
        If tmpLvItem.Selected Then
            If bWasUnSel Then
                Set newLvItem = lv.ListItems.Add(lv.ListItems.Count + 1, , tmpLvItem.text, iIcon, iSmallIcon)
                newLvItem.Tag = tmpLvItem.Tag
                newLvItem.Checked = tmpLvItem.Checked
                newLvItem.key = tmpLvItem.key
                For Each tmpSubItem In tmpLvItem.ListSubItems
                    newLvItem.SubItems(tmpSubItem.Index) = tmpSubItem.text
                Next
                lv.ListItems.Remove (tmpLvItem.Index)
                newLvItem.Selected = True
                Set newLvItem = Nothing
            End If
        Else
            bWasUnSel = True
        End If
        Set tmpLvItem = Nothing
    Next
End Sub

Public Sub ListViewMoveSelUp(ByVal lv As ListView)
    Dim bWasUnSel As Boolean
    Dim tmpLvItem As ListItem
    Dim newLvItem As ListItem
    Dim tmpSubItem As ListSubItem
    Dim iIcon As Integer
    Dim iSmallIcon As Integer
    Dim i As Integer
    bWasUnSel = False
    For i = 1 To lv.ListItems.Count
        Set tmpLvItem = lv.ListItems(i)
        iIcon = lv.ListItems(i).Icon
        iSmallIcon = lv.ListItems(i).SmallIcon
        If tmpLvItem.Selected Then
            If bWasUnSel Then
                Set newLvItem = lv.ListItems.Add(tmpLvItem.Index - 1, , tmpLvItem.text, iIcon, iSmallIcon)
                newLvItem.Tag = tmpLvItem.Tag
                newLvItem.Checked = tmpLvItem.Checked
                newLvItem.key = tmpLvItem.key
                For Each tmpSubItem In tmpLvItem.ListSubItems
                    newLvItem.SubItems(tmpSubItem.Index) = tmpSubItem.text
                Next
                lv.ListItems.Remove (tmpLvItem.Index)
                newLvItem.Selected = True
                Set newLvItem = Nothing
            End If
        Else
            bWasUnSel = True
        End If
        Set tmpLvItem = Nothing
    Next
End Sub

Public Sub ListViewMoveSelDown(ByVal lv As ListView)
    Dim bWasUnSel As Boolean
    Dim tmpLvItem As ListItem
    Dim newLvItem As ListItem
    Dim tmpSubItem As ListSubItem
    Dim iIcon As Integer
    Dim iSmallIcon As Integer
    Dim i As Integer
    bWasUnSel = False
    For i = lv.ListItems.Count To 1 Step -1
        Set tmpLvItem = lv.ListItems(i)
        iIcon = lv.ListItems(i).Icon
        iSmallIcon = lv.ListItems(i).SmallIcon
        If tmpLvItem.Selected Then
            If bWasUnSel Then
                Set newLvItem = lv.ListItems.Add(tmpLvItem.Index + 2, , tmpLvItem.text, iIcon, iSmallIcon)
                newLvItem.Tag = tmpLvItem.Tag
                newLvItem.Checked = tmpLvItem.Checked
                newLvItem.key = tmpLvItem.key
                For Each tmpSubItem In tmpLvItem.ListSubItems
                    newLvItem.SubItems(tmpSubItem.Index) = tmpSubItem.text
                Next
                lv.ListItems.Remove (tmpLvItem.Index)
                newLvItem.Selected = True
                Set newLvItem = Nothing
            End If
        Else
            bWasUnSel = True
        End If
        Set tmpLvItem = Nothing
    Next
End Sub

Public Sub Get_Cursor_Pos()


' Place the cursor positions in variable Hold
GetCursorPos HoldCursorPos

'' Display the cursor position coordinates
'MsgBox "X Position is : " & Hold.X_Pos & Chr(10) & _
'   "Y Position is : " & Hold.Y_Pos
End Sub

'-----------------------------------------------------------------------------
Public Function GetListviewVisibleCount(objListView As ListView) As Long
'-----------------------------------------------------------------------------
  
   GetListviewVisibleCount = SendMessage(objListView.hWnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)
   
End Function

Public Function AddBlank(s As String, LenToMake As Integer) As String

  Dim j As Integer
    
  j = LenToMake - Len(s$)
  If j > 0 Then
    AddBlank$ = s$ + Space$(j)
  Else
    AddBlank$ = s$
  End If

End Function

' translate a CTYPE value to text
Public Function GetCTypeString(ByVal ctype As Long, ByVal plugin As Long) As String
    If (plugin) Then ' using a plugin
        Dim pinfo As BASS_PLUGININFO, a As Long

        pinfo = BASS_PluginGetInfo(plugin)  ' get plugin info

        For a = 0 To pinfo.formatc - 1
            If (BASS_PluginGetInfoFormat(plugin, a).ctype = ctype) Then   ' found a "ctype" match...
                GetCTypeString = VBStrFromAnsiPtr(BASS_PluginGetInfoFormat(plugin, a).name)  ' return it's name
                Exit Function
            End If
        Next a
    End If

    ' check built-in stream formats...
    Select Case (ctype)
        Case (BASS_CTYPE_STREAM_OGG):   GetCTypeString = "Ogg Vorbis"
        Case (BASS_CTYPE_STREAM_MP1): GetCTypeString = "MPEG layer 1"
        Case (BASS_CTYPE_STREAM_MP2): GetCTypeString = "MPEG layer 2"
        Case (BASS_CTYPE_STREAM_MP3): GetCTypeString = "MPEG layer 3"
        Case (BASS_CTYPE_STREAM_AIFF): GetCTypeString = "Audio IFF"
        Case (BASS_CTYPE_STREAM_WAV_PCM): GetCTypeString = "PCM WAVE"
        Case (BASS_CTYPE_STREAM_WAV_FLOAT): GetCTypeString = "Floating-point WAVE"
        Case Else: GetCTypeString = "?"
    End Select

    ' other WAVE codec, could use acmFormatTagDetails to get its name, but...
    If (ctype And BASS_CTYPE_STREAM_WAV) Then GetCTypeString = "WAVE"
End Function


Public Sub CheckVersion()
On Error GoTo err1
Dim sTMP As String
'===============================================================
'               Routine to do autamated setup                  '
'===============================================================
'                                                              '
'===============================================================
'  Get all the default paths, version and names                '
'  from the Version.ini                                        '
'===============================================================
strVersionFile = App.FileDescription
GVersion = ReadIni(UCase(App.EXEName), "Version")
strAppPath = App.Path
strDistrib = ReadIni(UCase(App.EXEName), "Distribute")
strSetupPath = ReadIni(UCase(App.EXEName), "SetupPath")
strVersionPath = ReadIni(UCase(App.EXEName), "VersionPath")

'===============================================================
'                  Check versions are different                '
'===============================================================
'  If not the correct VERSION, start the DISTRIBUTION routine  '
'===============================================================
If GVersion <> App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000") Then
   MsgBox "Version " & GVersion & " will now be implemented !!!       ", vbExclamation, "New Version"
   '===============================================================
   '                       Start the copy program                 '
   '===============================================================
   '  This routine will copy the distribute program from the      '
   '  server. After this filecopy, it will SHELL (START) the      '
   '  distribute program to do the filecopy of the latest version '
   '  of the EXE (also from the server) to my App.path.           '
   '  It will end this program.                                   '
   '===============================================================
''''   Dim Ret
''''   Dim sString
''''   Ret = GetString(HKEY_LOCAL_MACHINE, HKeyVersion, "Version")
''''   If Ret = "" Then
''''        GoTo LoadDistibute
''''   Else
''''        If Ret <> sTMP Then
''''            GoTo LoadDistibute
''''        Else
''''          Exit Sub
''''        End If
''''    End If
''''Exit Sub
''''LoadDistibute:
   FileCopy strVersionPath & strDistrib, strAppPath & "\" & strDistrib
   '===============================================================
   '        EXCECUTE the copy-program on my LOCAL machine         '
   '  This process is a-synchronic, and will continue with the    '
   '  code AFTER the Shell-function. That means we can terminate  '
   '  this program now.                                           '
   '===============================================================
   X = Shell(strAppPath & "\" & strDistrib, vbNormalFocus)
   '===============================================================
   '                    Ends the current program                  '
   '===============================================================
   End
End If
Exit Sub
err1:
MsgBox "Error : " & Err.Number & Chr(13) & Chr(13) & "Description : " & Err.Description & " ...                " & Chr(13) & "Source : " & Err.Source & Chr(13) & Chr(13) & "                                    ", vbCritical, "Program ERROR"
End
End Sub

Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    RegOpenKey hKey, strPath, Ret
    GetString = RegQueryStringValue(Ret, strValue)
    RegCloseKey Ret
End Function

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function

Private Sub SaveStringLong(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret As Long
    Dim stringbuffer As String
    
    RegCreateKey hKey, strPath, Ret
    stringbuffer = strData & vbNullChar
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal stringbuffer, Len(stringbuffer)
    RegCloseKey Ret
    
End Sub

Public Function ReadIni(sApp As String, sKey As String, Optional sFile As Variant) As String
'=====================================================================
'  This routine reads the INI file caled VERSION.INI                 '
'  the first command (GetPrivateProfileInt) gets the key             '
'  to loop thru (in this example only)                               '
'  the second command (GetPrivateProfileString) gets the             '
'  actual antries from the INI after the '='                         '
'=====================================================================
'SYNTAX and MEANING:                                                 '
'=====================================================================
'GetPrivateProfileString(                                            '
'                        AppName,      :  Entry identifier in INI    '
'                        Key,          :  Key to look for in INI     '
'                        "",           :  Default ??                 '
'                        Return Value  :  Result found in INI        '
'                        100,          :  Lenght of string to read   '
'                        FileName      :  Path and Name of INI       '
'                        )                                           '
'=====================================================================
Dim sRetStr As String * 255
Dim iRetLen As Integer
Dim iPos1 As Long

On Error Resume Next
If Not IsMissing(sFile) Then
   strVersionFile = sFile
Else
   strVersionFile = App.Path & "\Info.ini"
End If

'Read the appropriate ini entry
iRetLen = GetPrivateProfileString(sApp, sKey, "", sRetStr, 255, strVersionFile)
If iRetLen <> 0 Then
   iPos1 = InStr(1, sRetStr, Chr(0))         'Kry posisie van einde van text
   ReadIni = Mid(sRetStr, 1, iPos1 - 1)      'Set Function = to return value
End If

End Function

Public Sub ColorListView(lsView As ListView, lColor As Long, iRow As Long, bBold As Boolean, DefltColor As Long, bReset As Boolean, ResetFind As Boolean, Optional iCol As Integer)
Dim i1 As Long
Dim iR As Long
Dim il As Long

'Reset All rows to the default colors
If bReset Then
  For iR = 1 To lsView.ListItems.Count
    If ResetFind Then 'Skip the currently selected (Green) entry, but clear all the rest
      If lsView.ListItems(iR).ForeColor = vbPlaylistSelColor Then GoTo ReadNetIR
    End If
     lsView.ListItems(iR).ForeColor = DefltColor
     lsView.ListItems(iR).Bold = False
   '  For iL = 1 To lsView.ListItems(iR).ListSubItems.Count
        lsView.ListItems(iR).ListSubItems(6).ForeColor = DefltColor
        lsView.ListItems(iR).ListSubItems(6).Bold = False
   '  Next iL
ReadNetIR:
  Next iR
End If

lsView.Refresh

'If ResetFind Then GoTo Eind


If iCol > 0 Then
   lsView.ListItems(iRow).ListSubItems(iCol).ForeColor = lColor
   lsView.ListItems(iRow).ListSubItems(iCol).Bold = bBold
Else
   lsView.ListItems(iRow).ForeColor = lColor
  ' lsView.ListItems(iRow).Ghosted = True
   lsView.ListItems(iRow).Bold = bBold
  ' For i1 = 1 To lsView.ListItems(iRow).ListSubItems.Count
      lsView.ListItems(iRow).ListSubItems(6).ForeColor = lColor
      lsView.ListItems(iRow).ListSubItems(6).Bold = bBold
  ' Next i1
End If
lsView.Refresh

Eind:

End Sub






