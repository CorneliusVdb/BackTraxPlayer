Attribute VB_Name = "mp3mod"
Option Explicit
Public Type Mp3IDTag
      Songname As String * 30
      artist As String * 30
      album As String * 30
      year As String * 4
      comment As String * 30
      genre As String * 1
End Type
Public Enum TEstado
    PAUSED
    PLAYING
    STOPED
End Enum
Private Type TMp3
    chan As Long
    Tag As Mp3IDTag
    Estado As TEstado
    OutR(256) As Single
    NList As Integer
End Type
Global Mp3(1 To 3) As TMp3

Public outdev(2) As Long   ' output devices
Public chan(3) As Long     ' the streams


Public Function GetTag(filename As String, ITag As Mp3IDTag) As Boolean
Dim f As Integer
Dim Tagg As String * 3
Dim temp As String
Dim i As Integer
Dim Cmd As String

f = FreeFile
On Error Resume Next
Open filename For Binary As #f
Get #f, FileLen(filename) - 127, Tagg
If Tagg = "TAG" Then
    GetTag = True
    Get #f, , ITag.Songname
    Get #f, , ITag.artist
    Get #f, , ITag.album
    Get #f, , ITag.year
    Get #f, , ITag.comment
    Get #f, , ITag.genre
    ITag.Songname = Replace(ITag.Songname, Chr(0), " ", , , vbBinaryCompare)
    ITag.artist = Replace(ITag.artist, Chr(0), " ", , , vbBinaryCompare)
    ITag.album = Replace(ITag.album, Chr(0), " ", , , vbBinaryCompare)
    ITag.year = Replace(ITag.year, Chr(0), " ", , , vbBinaryCompare)
    ITag.comment = Replace(ITag.comment, Chr(0), " ", , , vbBinaryCompare)
    ITag.genre = Replace(ITag.genre, Chr(0), " ", , , vbBinaryCompare)
Else
    GetTag = False
    temp = filename
    If InStr(temp, "\") Then
        Do
        temp = Right(temp, Len(temp) - 1)
        Loop Until InStr(temp, "\") = 0
        
        'If UCase(Right(temp, 4)) = ".MP3" Then
        temp = Left(temp, Len(temp) - 4)
        
        If InStr(temp, "-") Then
            ITag.artist = Left(temp, InStr(temp, "-") - 1)
            ITag.Songname = Right(temp, Len(temp) - InStr(temp, "-"))
        Else
            ITag.Songname = temp
        End If
    End If
End If

If Trim(ITag.Songname) = "" Then
    temp = filename
    If InStr(temp, "\") Then
        Do
        temp = Right(temp, Len(temp) - 1)
        Loop Until InStr(temp, "\") = 0
        
        'If UCase(Right(temp, 4)) = ".MP3" Then
        temp = Left(temp, Len(temp) - 4)
        
        If InStr(temp, "-") Then
            ITag.artist = Left(temp, InStr(temp, "-") - 1)
            ITag.Songname = Right(temp, Len(temp) - InStr(temp, "-"))
        Else
            ITag.Songname = temp
        End If
    End If



End If
Close #f
End Function

Public Sub BassInit(DeviceId As Long)
  Dim InitBass As Boolean
  Dim iDevice As Integer
  
  ' change and set the current path, to prevent from VB not finding BASS.DLL
  ChDrive App.Path
  ChDir App.Path
  
  iDevice = CInt(DeviceId)
  
  
  If FileExist(App.Path & "\bass.dll") = False Then
    MsgBox "BASS.DLL not exists", vbCritical, "BASS.DLL"
    End
  End If
  
    ' check the correct BASS was loaded
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
        End
    End If

  
  'Make sure we unload the bass devices...
  BassUnload
  
  ' initialize default output device
'''  If (BASS_Init(-1, 44100, 0, Me.hWnd, 0) = 0) Then
'''      Call Error_("Can't initialize device")
'''      End
'''  End If
'''
'''
  InitBass = BASS_Init(iDevice, 44100, 0, frmPlayer.hWnd, 0)
  'InitBass = BASS_Init(-1, 44100, 0, frmPlayer.hWnd)
  ' Start digital output
  If InitBass Then
    BASS_Start
  Else
    MsgBox "Error: Couldn't Initialize Digital Output", vbCritical, "Digital output"
    End
  End If
  

End Sub

Function FileExist(filename) As Boolean
  On Local Error Resume Next
  FileExist = (Dir$(filename) <> "")
End Function

Public Sub BassUnload()
  BASS_Stop
  BASS_StreamFree Mp3(1).chan
  BASS_StreamFree Mp3(2).chan
  BASS_StreamFree Mp3(3).chan
  BASS_Free
  
  Call BASS_PluginFree(0)
        
End Sub

Public Function CargarMp3(nCanal As Integer, filename As String) As Long

  BASS_StreamFree Mp3(nCanal).chan
  Mp3(nCanal).chan = BASS_StreamCreateFile(BASSFALSE, filename, 0, 0, 0)

  
  CargarMp3 = Mp3(nCanal).chan

End Function

Public Sub PlayChan(nCanal As Integer)




Call BASS_StreamFree(Mp3(nCanal).chan)
Call BASS_SetDevice(outdev(0)) ' set the device to create stream on
    
    
    
'Dim r As Long
' r = BASS_StreamPlay(Mp3(nCanal).chan, 0, BASS_SAMPLE_LOOP)
' If r = BASSTRUE Then
'    Mp3(nCanal).Estado = PLAYING
' Else
'    Mp3(nCanal).Estado = STOPED
'End If
End Sub

Public Sub VolumenChan(nCanal As Integer, vol As Integer)
Call BASS_ChannelSetAttributes(Mp3(nCanal).chan, 0, vol, -101)
End Sub

Public Function GetPosMin(nCanal As Integer) As Long
Dim pos As Long
    pos = BASS_ChannelBytes2Seconds(Mp3(nCanal).chan, modBass.BASS_ChannelGetPosition(Mp3(nCanal).chan))
    GetPosMin = pos \ 60
End Function

Public Function GetPosSec(nCanal As Integer) As Long
Dim pos As Long, min As Long
    pos = BASS_ChannelBytes2Seconds(Mp3(nCanal).chan, modBass.BASS_ChannelGetPosition(Mp3(nCanal).chan))
    If pos > -1 Then
        min = pos \ 60
        GetPosSec = pos - min * 60
    Else
        GetPosSec = 0
    End If
End Function

Public Sub SetPos(nCanal As Integer, pos As Integer)
Dim NewPos As Long
NewPos = pos * GetDuration(nCanal) / 100
NewPos = modBass.BASS_ChannelSeconds2Bytes(Mp3(nCanal).chan, NewPos)
modBass.BASS_ChannelSetPosition Mp3(nCanal).chan, NewPos
    
End Sub

Public Function GetPos(nCanal As Integer) As Long
    GetPos = BASS_ChannelBytes2Seconds(Mp3(nCanal).chan, modBass.BASS_ChannelGetPosition(Mp3(nCanal).chan))
End Function
Public Function GetDuration(nCanal As Integer) As Long
    GetDuration = BASS_ChannelBytes2Seconds(Mp3(nCanal).chan, modBass.BASS_StreamGetLength(Mp3(nCanal).chan))
End Function

Public Sub Pause(nCanal As Integer)
BASS_ChannelPause Mp3(nCanal).chan
Mp3(nCanal).Estado = PAUSED
End Sub
Public Sub StopPlaying(nCanal As Integer)
BASS_ChannelPause Mp3(nCanal).chan
Mp3(nCanal).Estado = STOPED
End Sub
Public Sub Play(nCanal As Integer)
BASS_ChannelResume Mp3(nCanal).chan
Mp3(nCanal).Estado = PLAYING
End Sub
Public Sub GetSpectrum(nCanal As Integer)
'Call BASS_ChannelGetData(Mp3(nCanal).chan, Mp3(nCanal).OutR(0), BASS_DATA_FFT512)
Call BASS_ChannelGetData(Mp3(nCanal).chan, Mp3(nCanal).OutR(0), BASS_DATA_FFT1024)
End Sub

