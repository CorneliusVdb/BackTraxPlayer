Attribute VB_Name = "modCustLoop"
'/////////////////////////////////////////////////////////////////////////////////
' modCustLoop.bas - Copyright (c) 2004-2007 (: JOBnik! :) [Arthur Aminov, ISRAEL]
'                                                         [http://www.jobnik.org]
'                                                         [  jobnik@jobnik.org  ]
' Other source: frmCustLoop.frm
'
' BASS custom looping example
' Originally translated from - custloop.c - Example of Ian Luck
'/////////////////////////////////////////////////////////////////////////////////

Option Explicit

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0&    ' color table in RGBs

Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(255) As RGBQUAD
End Type

Public Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Public Const TRANSPARENT = 1
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2

Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Const WIDTH_ As Long = 540  '600    ' display width
Public Const HEIGHT_ As Long = 51  '201   ' height (odd number for centre line)
Public bpp As Long          ' stream bytes per pixel
Public loop_(2) As Long     ' loop start & end
Public lsync As Long        ' looping sync
Public killscan As Boolean

Public wavebuf() As Byte    ' wave buffer
Public chanFreq As Long         ' stream/music handle
Public tmpChan As Long
Dim WavBuff As Long
Dim WavBuffDiv As Long
Public ReturnPos As Integer

Public bh As BITMAPINFO     ' bitmap header

' display error messages
Public Sub Error_(ByVal es As String)
    Call MsgBox(es & vbCrLf & vbCrLf & "error code: " & BASS_ErrorGetCode, vbExclamation, "Error")
End Sub

Sub LoopSyncProc(ByVal handle As Long, ByVal channel As Long, ByVal Data As Long, ByVal user As Long)
    If (BASS_ChannelSetPosition(channel, loop_(0), BASS_POS_BYTE) = 0) Then  ' try seeking to loop start
        Call BASS_ChannelSetPosition(channel, 0, BASS_POS_BYTE) ' failed, go to start of file instead
    End If
End Sub

Sub SetLoopStart(ByVal pos As Long)
   Dim text As String
    loop_(0) = pos
    
   ' a = bassTime.GetPlayingPos(chanFreq)
'    ReturnPos = loop_(0) / 60 / 60 / 60 / 10
    
    ReturnPos = BASS_ChannelBytes2Seconds(chanFreq, loop_(0) / 2)
    
    text = ReturnPos \ 60 & ":" & Format(ReturnPos Mod 60, "00")
    
'''    If (BASS_ChannelSetPosition(chanFreq, loop_(0), BASS_POS_BYTE) = 0) Then  ' try seeking to loop start
'''        Call BASS_ChannelSetPosition(chanFreq, 0, BASS_POS_BYTE) ' failed, go to start of file instead
'''    End If
    
End Sub

Sub SetLoopEnd(ByVal pos As Long)
    loop_(1) = pos
    Call BASS_ChannelRemoveSync(chanFreq, lsync) ' remove old sync
    lsync = BASS_ChannelSetSync(chanFreq, BASS_SYNC_POS Or BASS_SYNC_MIXTIME, loop_(1), AddressOf LoopSyncProc, 0) ' set new sync
End Sub

' scan the peaks
Sub ScanPeaks(ByVal decoder As Long)

    WavBuff = WIDTH_ * HEIGHT_
    WavBuffDiv = 65536  '49152   '32768
    ReDim wavebuf(-WavBuff To WavBuff) As Byte    ' set 'n clear the buffer (600 x 201 = 120600)
    'ReDim wavebuf(-120600 To 120600) As Byte    ' set 'n clear the buffer (600 x 201 = 120600)
    Dim cpos As Long, peak(2) As Long

    Do While (Not killscan)
        Dim level As Long, pos As Long
        level = BASS_ChannelGetLevel(decoder)  ' scan peaks
        pos = BASS_ChannelGetPosition(decoder, BASS_POS_BYTE) / bpp
        If (peak(0) < LoWord(level)) Then peak(0) = LoWord(level) ' set left peak
        If (peak(1) < HiWord(level)) Then peak(1) = HiWord(level) ' set right peak
        If (BASS_ChannelIsActive(decoder) = 0) Then
            pos = -1 ' reached the end
        Else
            pos = BASS_ChannelGetPosition(decoder, BASS_POS_BYTE) / bpp
        End If
        If (pos > cpos) Then
            Dim a As Long
            For a = 0 To (peak(0) * (HEIGHT_ / 2) / WavBuffDiv) - 1
            'For a = 0 To (peak(0) * (HEIGHT_ / 2) / 32768) - 1
                ' draw left peak WavBuff
                wavebuf(IIf((HEIGHT_ / 2 - 1 - a) * WIDTH_ + cpos > WavBuff, WavBuff, (HEIGHT_ / 2 - 1 - a) * WIDTH_ + cpos)) = 1 + a
                'wavebuf(IIf((HEIGHT_ / 2 - 1 - a) * WIDTH_ + cpos > 120600, 120600, (HEIGHT_ / 2 - 1 - a) * WIDTH_ + cpos)) = 1 + a
            Next a
            For a = 0 To (peak(1) * (HEIGHT_ / 2) / WavBuffDiv) - 1
            'For a = 0 To (peak(1) * (HEIGHT_ / 2) / 32768) - 1
                ' draw right peak
                wavebuf(IIf((HEIGHT_ / 2 + 1 + a) * WIDTH_ + cpos > WavBuff, WavBuff, (HEIGHT_ / 2 + 1 + a) * WIDTH_ + cpos)) = 1 + a
                'wavebuf(IIf((HEIGHT_ / 2 + 1 + a) * WIDTH_ + cpos > 120600, 120600, (HEIGHT_ / 2 + 1 + a) * WIDTH_ + cpos)) = 1 + a
            Next a
            If (pos >= WIDTH_) Then Exit Do ' gone off end of display
            cpos = pos
            peak(0) = 0
            peak(1) = 0
        End If
        DoEvents
    Loop
    Call BASS_StreamFree(decoder) ' free the decoder
    
End Sub

' scan the peaks to Find Silence
Public Function ScanSilence(sFile As String) As Double

   Dim level As Long, pos As Long
   Dim Ti As Double, Loc As Double
   Dim decoder As Long
   Dim killscan As Boolean
   
   
     'BASS_SAMPLE_LOOP Or BASS_SAMPLE_FX)

   Dim Bytes As Long
  ' Bytes = BASS_ChannelGetLength(decoder, BASS_POS_BYTE)
   Dim time As Long
  ' time = BASS_ChannelBytes2Seconds(decoder, Bytes)
   
   
   
   
 '  decoder = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, BASS_STREAM_DECODE)
   'If (decoder = 0) Then decoder = BASS_MusicLoad(BASSFALSE, StrPtr(sFile), 0, 0, BASS_MUSIC_DECODE, 1)
   decoder = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, 0)
   
   Ti = 0
   Do While (Not killscan)
      Loc = BASS_ChannelSeconds2Bytes(decoder, Ti)  'Get the location for each iteration
      Ti = Ti + 0.1  '0.01
      BASS_ChannelSetPosition decoder, Loc, BASS_POS_BYTE   'Set position to determine the level at that point...
      level = BASS_ChannelGetLevel(decoder)                 'scan peaks
    '  Debug.Print "Time : " & Ti & "   Level : " & Level
      If LoWord(level) > 100 Then   'silence_Level
         killscan = True
         Exit Do
      End If
      If HiWord(level) > 100 Then  'silence_Level
         killscan = True
         Exit Do
      End If
      
'      If Ti > 110 Then 'Only check first 30 seconds
'         killscan = True
'         Exit Do
'      End If
    Loop
    
    Call BASS_StreamFree(decoder) 'free the decoder
    
    If Ti > 1 Then Ti = Ti - 0.5  '0.25 'Move time sligthly backward so we have at least a half second silence in order for all the buffers to function in time
    If Ti < 0.5 Then Ti = 0      'Ignore if smaller than half second
    'If Ti < 0.25 Then Ti = 0      'Ignore if smaller than quater second
    If level = 0 Or Ti > 30 Then Ti = 0      'Ignore if the level = 0, or time > 30 seconds
    'Set the function = to value
    ScanSilence = Ti
    
End Function


'''''' select a file to play, and start scanning it
'''''Public Function PlayFile(sFile As String) As Boolean
'''''    On Local Error Resume Next    ' if Cancel pressed...
'''''
'''''
'''''        chanFreq = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, 0)
'''''        If (chanFreq = 0) Then chanFreq = BASS_MusicLoad(BASSFALSE, StrPtr(sFile), 0, 0, BASS_MUSIC_RAMPS Or BASS_MUSIC_POSRESET Or BASS_MUSIC_PRESCAN, 1)
'''''
'''''          frmDucking.Show   ' show form
'''''
'''''        With bh.bmiHeader
'''''            .biSize = Len(bh.bmiHeader)
'''''            .biWidth = WIDTH_
'''''            .biHeight = -HEIGHT_
'''''            .biPlanes = 1
'''''            .biBitCount = 8
'''''            .biClrUsed = HEIGHT_ / 2 + 1
'''''            .biClrImportant = HEIGHT_ / 2 + 1
'''''        End With
'''''
'''''        ' setup palette
'''''        Dim a As Byte
'''''
'''''        For a = 1 To HEIGHT_ / 2
'''''            bh.bmiColors(a).rgbRed = (255 * a) / (HEIGHT_ / 2)
'''''            bh.bmiColors(a).rgbGreen = 255 - bh.bmiColors(a).rgbRed
'''''        Next a
'''''
'''''        bpp = BASS_ChannelGetLength(chanFreq, BASS_POS_BYTE) / WIDTH_ ' bytes per pixel
'''''
'''''        If (bpp < BASS_ChannelSeconds2Bytes(chanFreq, 0.02)) Then ' minimum 20ms per pixel (BASS_ChannelGetLevel scans 20ms)
'''''            bpp = BASS_ChannelSeconds2Bytes(chanFreq, 0.02)
'''''        End If
''''''''        lsync = BASS_ChannelSetSync(chan, BASS_SYNC_END Or BASS_SYNC_MIXTIME, 0, AddressOf LoopSyncProc, 0) ' set sync to loop at end
''''''''        Call BASS_ChannelPlay(chan, BASSFALSE) ' start playing
'''''          frmDucking.tmrCustLoop.Enabled = True ' timer's interval is 100ms (10Hz)
'''''
'''''        Dim chan2 As Long
'''''        chan2 = BASS_StreamCreateFile(BASSFALSE, StrPtr(sFile), 0, 0, BASS_STREAM_DECODE)
'''''        If (chan2 = 0) Then chan2 = BASS_MusicLoad(BASSFALSE, StrPtr(sFile), 0, 0, BASS_MUSIC_DECODE, 1)
'''''        Call ScanPeaks(chan2)    ' start scanning peaks
'''''    'End With
'''''    PlayFile = True
'''''
'''''End Function

''''Sub DrawTimeLine(ByVal DC As Long, ByVal pos As Long, ByVal col As Long, ByVal Y As Long)
''''    Dim wpos As Long
''''    wpos = pos / bpp
''''    Dim time As Long
''''    time = BASS_ChannelBytes2Seconds(chanFreq, wpos)
''''    Dim text As String
''''    text = time \ 60 & ":" & Format(time Mod 60, "00")
''''      frmDucking.sspProgress.CurrentX = wpos
''''      frmDucking.sspProgress.Line (wpos, 0)-(wpos, HEIGHT_ - 1), col
'''''''    Call SetTextColor(DC, col)
'''''''    Call SetBkMode(DC, TRANSPARENT)
'''''''    Call SetTextAlign(DC, IIf(wpos >= WIDTH_ / 2, TA_RIGHT, TA_LEFT))
'''''''    Call TextOut(DC, wpos, Y, text, Len(text))
''''End Sub
