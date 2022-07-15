VERSION 5.00
Begin VB.UserControl ProgressBar 
   CanGetFocus     =   0   'False
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2700
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   180
   ToolboxBitmap   =   "ProgressBar.ctx":0000
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type SIZE
        cx As Long
        cy As Long
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
Private Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long

Private Const MM_ANISOTROPIC = 8

Private Const PS_DASH = 1                    '  -------
Private Const PS_DASHDOT = 3                 '  _._._._
Private Const PS_DASHDOTDOT = 4              '  _.._.._
Private Const PS_DOT = 2                     '  .......
Private Const PS_SOLID = 0
Private Const PS_NULL = 5
Private Const PS_INSIDEFRAME = 6

Private Const NULL_BRUSH = 5
Private Const NULL_PEN = 8
Private Const DKGRAY_BRUSH = 3
Private Const GRAY_BRUSH = 2
Private Const HOLLOW_BRUSH = NULL_BRUSH
Private Const LTGRAY_BRUSH = 1
Private Const WHITE_BRUSH = 0
Private Const BLACK_BRUSH = 4
Private Const WHITE_PEN = 6
Private Const BLACK_PEN = 7

Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_TOPLEFT Or BF_BOTTOMRIGHT)

Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public Enum StyleConstants
    scNone
    scBump
    scEtched
    scRaised
    scSingle
    scSunken
End Enum

Public Enum OrientationConstants
    ocB2T
    ocL2R
    ocR2L
    ocT2B
End Enum

'Default Property Values:
Const m_def_Steps = 100
Const m_def_FromColor = vbWhite
Const m_def_ToColor = vbBlue
Const m_def_Value = 0
Const m_def_BorderStyle = scSunken
Const m_def_BorderColor = vbBlack
Const m_def_BackColor = vbButtonFace
Const m_def_Orientation = ocL2R

'Property Variables:
Dim m_Steps As Long
Dim m_FromColor As OLE_COLOR
Dim m_ToColor As OLE_COLOR
Dim m_Value As Long
Dim m_BorderStyle As StyleConstants
Dim m_BorderColor As OLE_COLOR
Dim m_Orientation As OrientationConstants

Dim m_lHDC As Long
Dim m_lHBMP As Long


Private Function Min(ByVal Value As Variant, ByVal MinVal As Variant) As Variant
    On Error GoTo ErrHandler
    Min = IIf(Value < MinVal, MinVal, Value)
    Exit Function
ErrHandler:
    Min = Value
End Function
Private Function Max(ByVal Value As Variant, ByVal MaxVal As Variant) As Variant
    On Error GoTo ErrHandler
    Max = IIf(Value > MaxVal, MaxVal, Value)
    Exit Function
ErrHandler:
    Max = Value
End Function

Friend Function ConvertColor(ByVal Value As OLE_COLOR) As OLE_COLOR
    On Error Resume Next
    If Value < 0 Then OleTranslateColor Value, 0, Value
    ConvertColor = Value
End Function

Public Property Get Orientation() As OrientationConstants
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal Value As OrientationConstants)
    m_Orientation = Value
    PropertyChanged "Orientation"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbButtonFace
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to draw the progress bar."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal Value As OLE_COLOR)
    UserControl.BackColor = ConvertColor(Value)
    PropertyChanged "BackColor"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Steps() As Long
Attribute Steps.VB_ProcData.VB_Invoke_Property = ";Misc"
    Steps = m_Steps
End Property

Public Property Let Steps(ByVal Value As Long)
    m_Steps = Min(Value, 0&)
    PropertyChanged "Steps"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Step(Optional ByVal Value As Long = 1) As Boolean
    Me.Value = Me.Value + Value
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbWhite
Public Property Get FromColor() As OLE_COLOR
Attribute FromColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FromColor = m_FromColor
End Property

Public Property Let FromColor(ByVal Value As OLE_COLOR)
    m_FromColor = ConvertColor(Value)
    PropertyChanged "FromColor"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbBlue
Public Property Get ToColor() As OLE_COLOR
Attribute ToColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ToColor = m_ToColor
End Property

Public Property Let ToColor(ByVal Value As OLE_COLOR)
    m_ToColor = ConvertColor(Value)
    PropertyChanged "ToColor"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    Value = m_Value
End Property

Public Property Let Value(ByVal Value As Long)
    m_Value = Min(Max(Value, m_Steps), 0)
    PropertyChanged "Value"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BorderStyle() As StyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the style of border to be drawn on the progress bar control."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As StyleConstants)
    m_BorderStyle = Value
    PropertyChanged "BorderStyle"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbBlack
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color to use when drawing the singe line border around the progress bar. (Only applies when BorderStyle = scSingle)"
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = ConvertColor(Value)
    PropertyChanged "BorderColor"
    Refresh
End Property

Private Sub UserControl_Hide()
    On Error Resume Next
    If m_lHBMP <> 0 Then DeleteObject m_lHBMP: m_lHBMP = 0
    If m_lHDC <> 0 Then DeleteDC m_lHDC: m_lHDC = 0
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    On Error Resume Next
    m_Steps = m_def_Steps
    m_FromColor = ConvertColor(m_def_FromColor)
    m_ToColor = ConvertColor(m_def_ToColor)
    m_Value = m_def_Value
    m_BorderStyle = m_def_BorderStyle
    m_BorderColor = ConvertColor(m_def_BorderColor)
    m_Orientation = m_def_Orientation
    UserControl.BackColor = ConvertColor(m_def_BackColor)
    Refresh
End Sub

Private Sub UserControl_Paint()
    'Refresh
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        m_Steps = .ReadProperty("Steps", m_def_Steps)
        m_FromColor = ConvertColor(.ReadProperty("FromColor", m_def_FromColor))
        m_ToColor = ConvertColor(.ReadProperty("ToColor", m_def_ToColor))
        m_Value = .ReadProperty("Value", m_def_Value)
        m_BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        m_BorderColor = ConvertColor(.ReadProperty("BorderColor", m_def_BorderColor))
        m_Orientation = .ReadProperty("Orientation", m_def_Orientation)
        UserControl.BackColor = ConvertColor(.ReadProperty("BackColor", m_def_BackColor))
    End With
    Refresh
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If m_lHBMP <> 0 Then DeleteObject m_lHBMP: m_lHBMP = 0
    If m_lHDC <> 0 Then DeleteDC m_lHDC: m_lHDC = 0
    Refresh
End Sub

Private Sub UserControl_Show()
    ' Refresh
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    If m_lHBMP <> 0 Then DeleteObject m_lHBMP: m_lHBMP = 0
    If m_lHDC <> 0 Then DeleteDC m_lHDC: m_lHDC = 0
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "Steps", m_Steps, m_def_Steps
        .WriteProperty "FromColor", m_FromColor, m_def_FromColor
        .WriteProperty "ToColor", m_ToColor, m_def_ToColor
        .WriteProperty "Value", m_Value, m_def_Value
        .WriteProperty "BorderStyle", m_BorderStyle, m_def_BorderStyle
        .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
        .WriteProperty "Orientation", m_Orientation, m_def_Orientation
        .WriteProperty "BackColor", UserControl.BackColor, m_def_BackColor
    End With
End Sub

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550

    On Error GoTo ErrHandler
    
    Dim lBrush As Long
    Dim lPen As Long
    Dim r As RECT
    Dim sz As SIZE
    Dim mm As Long
    Dim i As Long
    Dim pt As POINTAPI
    Dim dStep As Double
    Dim dDelta(1 To 3) As Double
    Dim dColor(1 To 3) As Long
    
    Static amDoing As Boolean
    If amDoing Then Exit Sub
    amDoing = True
    
    With UserControl
    
        .AutoRedraw = True
    
        ' This section is used to create the initial bitmap
        ' and offscreen DC that's used. Only happens when
        ' the size changes or when the object is first created
        If m_lHDC = 0 Then
    
            ' always do drawings in pixels
            .ScaleMode = vbPixels
    
            RealizePalette .hdc
    
            m_lHDC = CreateCompatibleDC(.hdc)
            m_lHBMP = CreateCompatibleBitmap(.hdc, .ScaleWidth, .ScaleHeight)
    
            SelectObject m_lHDC, m_lHBMP
    
        End If
        
        ' Basically just to clear the drawing by painting the whole
        ' area, the background color.
        lPen = CreatePen(PS_SOLID, 1, BackColor)
        lBrush = CreateSolidBrush(BackColor)
        SelectObject m_lHDC, lPen
        SelectObject m_lHDC, lBrush
        Rectangle m_lHDC, 0, 0, .ScaleWidth, .ScaleHeight
        SelectObject m_lHDC, GetStockObject(NULL_PEN)
        SelectObject m_lHDC, GetStockObject(NULL_BRUSH)
        DeleteObject lPen
        DeleteObject lBrush
    
        ' This section is used to break down the FromColor
        ' into its constituent RGB values and to determine
        ' the amount of change to each R-G-B value, so that
        ' the color transition will be a smooth one.
        dColor(1) = FromColor And &HFF0000  ' Blue
        dColor(2) = FromColor And &HFF00&  ' Green
        dColor(3) = FromColor And &HFF& ' Red
        If dColor(1) > 0 Then dColor(1) = dColor(1) / &H10000
        If dColor(2) > 0 Then dColor(2) = dColor(2) / &H100&
        dDelta(1) = ToColor And &HFF0000
        dDelta(2) = ToColor And &HFF00&
        dDelta(3) = ToColor And &HFF&
        If dDelta(1) > 0 Then dDelta(1) = dDelta(1) / &H10000
        If dDelta(2) > 0 Then dDelta(2) = dDelta(2) / &H100&
        
        dDelta(1) = dDelta(1) - dColor(1)
        If dDelta(1) <> 0 Then dDelta(1) = dDelta(1) / 255&
        dDelta(2) = dDelta(2) - dColor(2)
        If dDelta(2) <> 0 Then dDelta(2) = dDelta(2) / 255&
        dDelta(3) = dDelta(3) - dColor(3)
        If dDelta(3) <> 0 Then dDelta(3) = dDelta(3) / 255&
        
        ' Set the map mode such that the height or width (based
        ' on orientation) is seen as 255 "units". In each unit or height
        ' or width, you'll draw one shade of the colors such that
        ' the color smoothly transitions from the "FromColor"
        ' to the "ToColor"
        mm = SetMapMode(m_lHDC, MM_ANISOTROPIC)
        If m_Orientation = ocB2T Or m_Orientation = ocT2B Then
            SetWindowExtEx m_lHDC, .ScaleWidth, 255, sz
            SetViewportExtEx m_lHDC, .ScaleWidth, .ScaleHeight, sz
        Else
            SetWindowExtEx m_lHDC, 255, .ScaleHeight, sz
            SetViewportExtEx m_lHDC, .ScaleWidth, .ScaleHeight, sz
        End If
        
        ' Determine just how many "units" each step should cover
        dStep = Min(Steps, 1) / 255&
        
        For i = 0 To 255
    
            If i * dStep > Value Or Value = 0 Then Exit For
    
            ' have to use a wide pen to avoid ugliness
            lPen = CreatePen(PS_SOLID, 3, RGB((dColor(3) + (dDelta(3) * i)) And &HFF&, _
                                              (dColor(2) + (dDelta(2) * i)) And &HFF&, _
                                              (dColor(1) + (dDelta(1) * i)) And &HFF&))
            SelectObject m_lHDC, lPen

            Select Case m_Orientation
            Case ocB2T
                MoveToEx m_lHDC, 0, Abs(256 - i), pt
                LineTo m_lHDC, .ScaleWidth, Abs(256 - i)
            Case ocT2B
                MoveToEx m_lHDC, 0, i - 1, pt
                LineTo m_lHDC, .ScaleWidth, i - 1
            Case ocR2L
                MoveToEx m_lHDC, Abs(256 - i), 0, pt
                LineTo m_lHDC, Abs(256 - i), .ScaleHeight
            Case Else ' ocL2R
                MoveToEx m_lHDC, i - 1, 0, pt
                LineTo m_lHDC, i - 1, .ScaleHeight
            End Select
    
            SelectObject m_lHDC, GetStockObject(NULL_PEN)
            DeleteObject lPen
            
        Next
        
        ' Return the map mode to original setting
        SetMapMode m_lHDC, mm
    
        ' put the border coords in a RECT structure
        SetRect r, 0&, 0&, .ScaleWidth, .ScaleHeight
    
        ' Draw the border (if any)
        Select Case (BorderStyle)
        Case scBump
            DrawEdge m_lHDC, r, EDGE_BUMP, BF_RECT
        Case scEtched
            DrawEdge m_lHDC, r, EDGE_ETCHED, BF_RECT
        Case scRaised
            DrawEdge m_lHDC, r, EDGE_RAISED, BF_RECT
        Case scSingle
            lBrush = CreateSolidBrush(BorderColor)
            SelectObject m_lHDC, lBrush
            FrameRect m_lHDC, r, lBrush
            SelectObject m_lHDC, GetStockObject(NULL_BRUSH)
            DeleteObject lBrush
        Case scSunken
            DrawEdge m_lHDC, r, EDGE_SUNKEN, BF_RECT
        Case Else ' StyleConstants.scNone
        End Select
    
        ' Make sure it refreshes the whole area
        InvalidateRect .hWnd, r, 0&

        ' Place the new drawing back into the UserControl
        BitBlt .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, _
            m_lHDC, 0&, 0&, SRCCOPY
        
        .AutoRedraw = False
        
    End With
    
    amDoing = False
    
Exit Sub
ErrHandler:

    amDoing = False

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub
