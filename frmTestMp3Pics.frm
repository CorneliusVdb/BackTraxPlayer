VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTestMp3Pics 
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   18435
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   18435
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8775
      Top             =   5205
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3825
      Left            =   4200
      ScaleHeight     =   3765
      ScaleWidth      =   3030
      TabIndex        =   2
      Top             =   285
      Width           =   3090
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   255
      TabIndex        =   1
      Top             =   390
      Width           =   3570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   11865
      TabIndex        =   0
      Top             =   5910
      Width           =   1350
   End
End
Attribute VB_Name = "frmTestMp3Pics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' used to create stdPicture from byte array (VB6 image types of bmp, gif, jpg, wmf, emf, ico only)
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function OleLoadPicture Lib "OLEPRO32.DLL" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' used to create stdPicture from byte array via GDI+ (VB6 image types + PNG,TIFF)
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus.dll" (ByVal pbitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus.dll" (ByVal Stream As Long, Image As Long) As Long
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Enum MyImageTypeEnum
    imgBMP = 0
    imgJPG = 1
    imgGIF = 2
    imgWMF = 3
    imgPNG = 4
    imgTIF = 5
    imgUNK = 6
End Enum
Private m_MP3 As cTaggedImages

Private Sub Form_Load()
    Set m_MP3 = New cTaggedImages
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "MP3 Image To View"
        .Filter = "Supported Files|*.mp3;*.wma|MP3 Files|*.mp3|WMA Files|*.wma|All Files|*.*"
        .Flags = cdlOFNFileMustExist Or cdlOFNExplorer
    End With
    Command1.Move 90, 75, 2535, 450
    Command1.Caption = "Load MP3 File"
    List1.Move 60, 540, 4640, 1815
    List1.Enabled = False
    Picture1.Move 4850, 75, 4140, 2310
    'Picture1.AutoSize = False
    'Picture1.BackColor = vbWhite
    Me.Width = Me.Width + 3000
End Sub

Private Sub Command1_Click()

Dim fNum As Integer
 Dim sTagIdent As String * 3
 Dim sTitle As String * 30
 Dim sArtist As String * 30
 Dim sAlbum As String * 30
 Dim sYear As String * 4
 Dim sComment As String * 30

    
    Dim i As Long, sFormat As String
    On Error GoTo ExitRoutine
    CommonDialog1.ShowOpen
    
' fNum = FreeFile
''Replace 'c:\MySong.mp3' with the name of the MP3 file that you want to get its info.
' Open CommonDialog1.FileName For Binary As fNum
' Seek #fNum, LOF(fNum) - 127
' Get #fNum, , sTagIdent
' If sTagIdent = "TAG" Then
' Get #fNum, , sTitle
' Get #fNum, , sArtist
' Get #fNum, , sAlbum
' Get #fNum, , sYear
' Get #fNum, , sComment
' End If
' Close #fNum
' MsgBox sTitle & "," & sArtist & "," & sAlbum & "," & sYear & "," & sComment
 
 
' Exit Sub
 
 
 
    
    List1.Clear
    If m_MP3.LoadTagged_Images(CommonDialog1.FileName) = 0 Then
        List1.AddItem "No images found"
        List1.Enabled = False
    Else
        For i = 1 To m_MP3.ImageCount
            sFormat = m_MP3.ImageType(i)
            If InStr(1, sFormat, "JPG", vbTextCompare) Then
                List1.AddItem "JPG image": List1.ItemData(List1.NewIndex) = imgJPG
            ElseIf InStr(1, sFormat, "JPEG", vbTextCompare) Then
                List1.AddItem "JPG image": List1.ItemData(List1.NewIndex) = imgJPG
            ElseIf InStr(1, sFormat, "PNG", vbTextCompare) Then
                List1.AddItem "PNG image": List1.ItemData(List1.NewIndex) = imgPNG
            ElseIf InStr(1, sFormat, "BMP", vbTextCompare) Then
                List1.AddItem "BMP image": List1.ItemData(List1.NewIndex) = imgBMP
            ElseIf InStr(1, sFormat, "BITMAP", vbTextCompare) Then
                List1.AddItem "BMP image": List1.ItemData(List1.NewIndex) = imgBMP
            ElseIf InStr(1, sFormat, "GIF", vbTextCompare) Then
                List1.AddItem "GIF image": List1.ItemData(List1.NewIndex) = imgGIF
            ElseIf InStr(1, sFormat, "METAFILE", vbTextCompare) Then
                List1.AddItem "WMF image": List1.ItemData(List1.NewIndex) = imgWMF
            ElseIf InStr(1, sFormat, "WMF", vbTextCompare) Then
                List1.AddItem "WMF image": List1.ItemData(List1.NewIndex) = imgWMF
            ElseIf InStr(1, sFormat, "TIF", vbTextCompare) Then
                List1.AddItem "TIF image": List1.ItemData(List1.NewIndex) = imgTIF
            Else
                List1.AddItem "Unknown format": List1.ItemData(List1.NewIndex) = imgUNK
            End If
            
            Select Case m_MP3.ImageCategory(i)
                Case 0: sFormat = "Other"
                Case 1: sFormat = "File icon"
                Case 2: sFormat = "Other file icon"
                Case 3: sFormat = "Cover (front)"
                Case 4: sFormat = "Cover (back)"
                Case 5: sFormat = "Leaflet Page"
                Case 6: sFormat = "Media" ' (e.g. label side of CD)
                Case 7: sFormat = "Lead artist/lead performer/soloist"
                Case 8: sFormat = "Artist/performer"
                Case 9: sFormat = "Conductor"
                Case 10: sFormat = "Band/Orchestra"
                Case 11: sFormat = "Composer"
                Case 12: sFormat = "Lyricist/text writer"
                Case 13: sFormat = "Recording Location"
                Case 14: sFormat = "During recording"
                Case 15: sFormat = "During performance"
                Case 16: sFormat = "Movie/video screen capture"
                Case 17: sFormat = "A bright coloured fish"
                Case 18: sFormat = "Illustration"
                Case 19: sFormat = "Band/artist logotype"
                Case 20: sFormat = "Publisher/Studio logotype"
                Case Else: sFormat = "Undetermined"
            End Select
            List1.List(List1.NewIndex) = List1.List(List1.NewIndex) & " Cat: " & sFormat
            
        Next
        List1.Enabled = True
        List1.ListIndex = 0
    End If
ExitRoutine:
End Sub

Private Sub List1_Click()
    
    Dim bBytes() As Byte, tmpPic As StdPicture
    If m_MP3.ExtractImageData(List1.ListIndex + 1, bBytes()) = False Then Exit Sub
        
    If List1.ItemData(List1.ListIndex) < imgPNG Then
        ' use OLE API to create image
        Set tmpPic = ArrayToPicture(VarPtr(bBytes(0)), UBound(bBytes) + 1)
        ' should above fail, we'll default to GDI+ below
    End If
    If tmpPic Is Nothing Then
        ' use GDI+ API to create image
         Set tmpPic = ArrayToGDIplusStdPicture(VarPtr(bBytes(0)), UBound(bBytes) + 1)
    End If
    Set Picture1.Picture = tmpPic
    
End Sub

Private Function ArrayToPicture(arrayVarPtr As Long, lSize As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    ' Note: The array was already validated as not empty before this was called
    
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    On Error GoTo ExitRoutine
    Set IIStream = IStreamFromArray(arrayVarPtr, lSize)
    
    If Not IIStream Is Nothing Then
        aGUID(0) = &H7BF80980    ' GUID for stdPicture
        aGUID(1) = &H101ABF32
        aGUID(2) = &HAA00BB8B
        aGUID(3) = &HAB0C3000
        Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)
    End If
    
ExitRoutine:
End Function

Private Function HandleToStdPicture(ByVal hImage As Long, ByVal imgType As PictureTypeConstants) As IPicture

    ' function creates a stdPicture object from an image handle (bitmap or icon)
    
    'Private Type PictDesc
    '    Size As Long
    '    Type As Long
    '    hHandle As Long
    '    lParam As Long       for bitmaps only: Palette handle
    '                         for WMF only: extentX (integer) & extentY (integer)
    '                         for EMF/ICON: not used
    'End Type
    
    Dim lpPictDesc(0 To 3) As Long, aGUID(0 To 3) As Long
    
    lpPictDesc(0) = 16&
    lpPictDesc(1) = imgType
    lpPictDesc(2) = hImage
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc(0), aGUID(0), True, HandleToStdPicture)
    
End Function

Private Function ArrayToGDIplusStdPicture(ArrayPtr As Long, Length As Long) As IPicture

    Dim gToken As Long, gSUI As GdiplusStartupInput
    Dim hImage As Long, hBitmap As Long
    Dim IStream As IUnknown
    
    gSUI.GdiplusVersion = 1
    If GdiplusStartup(gToken, gSUI) = 0 Then
        Set IStream = IStreamFromArray(ArrayPtr, Length)
        If Not IStream Is Nothing Then
            If GdipLoadImageFromStream(ObjPtr(IStream), hImage) = 0 Then
                ' create a standard BMP from GDI+ image. Set fill color to BGR vs. RGB
                GdipCreateHBITMAPFromBitmap hImage, hBitmap, _
                     (Picture1.BackColor And &HFF) * &H10000 Or (Picture1.BackColor And &HFF00&) Or _
                     (Picture1.BackColor And &HFF0000) \ &H10000 Or &HFF000000
                GdipDisposeImage hImage
            End If
            Set IStream = Nothing
        End If
        GdiplusShutdown gToken
        If hBitmap Then Set ArrayToGDIplusStdPicture = HandleToStdPicture(hBitmap, vbPicTypeBitmap)
    End If

End Function

Private Function IStreamFromArray(ArrayPtr As Long, Length As Long) As stdole.IUnknown
    
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
    
    On Error GoTo HandleError
    Dim o_hMem As Long
    Dim o_lpMem  As Long
     
    If ArrayPtr = 0& Then
        CreateStreamOnHGlobal 0&, 1&, IStreamFromArray
    ElseIf Length <> 0& Then
        o_hMem = GlobalAlloc(&H2&, Length)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, ByVal ArrayPtr, Length
                Call GlobalUnlock(o_hMem)
                Call CreateStreamOnHGlobal(o_hMem, 1&, IStreamFromArray)
            End If
        End If
    End If
    
HandleError:
End Function



