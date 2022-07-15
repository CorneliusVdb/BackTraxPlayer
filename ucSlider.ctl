VERSION 5.00
Begin VB.UserControl ucSlider 
   BackStyle       =   0  'Transparent
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   LockControls    =   -1  'True
   ScaleHeight     =   255
   ScaleWidth      =   3135
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   45
      ScaleHeight     =   225
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   15
      Width           =   3075
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   49
         Left            =   2940
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   48
         Left            =   2880
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   47
         Left            =   2820
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00003FFF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   46
         Left            =   2760
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00003FFF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   45
         Left            =   2700
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000087FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   44
         Left            =   2640
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000087FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   43
         Left            =   2580
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000087FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   42
         Left            =   2520
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000CAFF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   41
         Left            =   2460
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000CAFF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   40
         Left            =   2400
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000CAFF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   39
         Left            =   2340
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000E7FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   38
         Left            =   2280
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000E7FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   37
         Left            =   2220
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000E7FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   36
         Left            =   2160
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFF5&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   35
         Left            =   2100
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFF5&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   34
         Left            =   2040
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFF5&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   33
         Left            =   1980
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H006EEFE8&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   32
         Left            =   1920
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H006EEFE8&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   31
         Left            =   1860
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H006EEFE8&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   30
         Left            =   1800
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H006BEBD8&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   29
         Left            =   1740
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H006BEBD8&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   28
         Left            =   1680
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H006BEBD8&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   27
         Left            =   1620
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0067E8C9&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   26
         Left            =   1560
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0067E8C9&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   25
         Left            =   1500
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Image imgSlider 
         Height          =   270
         Left            =   4800
         Picture         =   "ucSlider.ctx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0033B234&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   0
         Left            =   0
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0033B234&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   1
         Left            =   60
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0033B234&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   2
         Left            =   120
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0033B234&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   3
         Left            =   180
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0033B234&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   4
         Left            =   240
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0033B234&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   5
         Left            =   300
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0033B234&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   6
         Left            =   360
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   7
         Left            =   420
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   8
         Left            =   480
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H003DD171&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   9
         Left            =   540
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H003DD171&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   10
         Left            =   600
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H003DD171&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   11
         Left            =   660
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0058D886&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   12
         Left            =   720
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0058D886&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   13
         Left            =   780
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0058D886&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   14
         Left            =   840
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0057D786&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   15
         Left            =   900
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0057D786&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   16
         Left            =   960
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0057D786&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   17
         Left            =   1020
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H005CDC9B&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   18
         Left            =   1080
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H005CDC9B&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   19
         Left            =   1140
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H005CDC9B&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   20
         Left            =   1200
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0063E3BA&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   21
         Left            =   1260
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0063E3BA&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   22
         Left            =   1320
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0063E3BA&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   23
         Left            =   1380
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0067E8C9&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   24
         Left            =   1440
         Top             =   45
         Visible         =   0   'False
         Width           =   75
      End
   End
End
Attribute VB_Name = "ucSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' some code borrowed and modified from Richard Allsebrooks submission EBSlider
' the rest is mine
Public Enum Bar
 Rainbow = 0
 Red = 1
 Blue = 2
 Green = 3
 Yellow = 4
 Purple = 5
 Turq = 6
 Gray = 7
End Enum

Const m_def_DropDownCtrl = True
Const m_def_Caption = ""
Const m_def_DisableDropDown = False
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 0
Const m_def_TickColor = vbBlack
Const m_def_BackColor = &HC0C0C0
Const m_def_CapBkgdColor = vbWhite
Const m_def_CapFontColor = vbBlack
Const m_def_BarColor = 1
Const m_def_ValueHide = True

Dim m_ValueHide As Boolean
Dim m_DropDownCtrl As Boolean
Dim m_Caption As String
Dim m_DisableDropDown As Boolean
Dim m_Min As Long
Dim m_Max As Long
Dim m_Value As Long
Dim m_TickColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_CapBkgdColor As OLE_COLOR
Dim m_CapFontColor As OLE_COLOR
Dim m_BarColor As Integer

Private SliderWidth      As Long

Private Sub Image1_Click()
'   DropDownCtrl = False
'   Image1.Visible = False
'   Image2.Visible = True
End Sub

Private Sub Image2_Click()
'   DropDownCtrl = True
'   Image1.Visible = True
'   Image2.Visible = False
End Sub

Private Sub lblCaption_Click()
'If DisableDropDown = True Then Exit Sub
'DropDownCtrl = Not DropDownCtrl
End Sub

Private Sub UserControl_Initialize()
Dim i As Integer
   m_Min = m_def_Min
   m_Max = m_def_Max
   m_Value = m_def_Value
   m_TickColor = m_def_TickColor
   m_BackColor = m_def_BackColor
   m_CapBkgdColor = m_def_CapBkgdColor
   m_CapFontColor = m_def_CapFontColor
   m_DisableDropDown = False   'm_def_DisableDropDown
   m_BarColor = m_def_BarColor
   m_ValueHide = m_def_ValueHide
   SliderWidth = 130
  
   For i = 0 To 49
    Shape1(i).Height = 120
    Shape1(i).top = Shape1(0).top
   Next i
   imgSlider.top = 0
      
End Sub

Private Sub UserControl_InitProperties()
   Caption = Extender.name
   Min = 0
   Max = 100
   value = 0
'   DropDownCtrl = True
'   DisableDropDown = False

   SliderWidth = 130
   SliderPos
   
End Sub

'Private Sub UserControl_Resize()
'   If DropDownCtrl = False Then
'      UserControl.Height = 240
'   Else
'      UserControl.Height = 825
'   End If
'   UserControl.Width = 2250
'End Sub

Private Sub imgSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim sldPos  As Long
   Dim sldScale  As Single
   Dim shp As Integer
   
   If Button = 1 Then
      With imgSlider
         sldPos = ((.Left + X - SliderWidth / 2) \ 15) * 15
         If sldPos < 0 Then
            sldPos = 0
         ElseIf sldPos > picForm.Width - SliderWidth Then
            sldPos = picForm.Width - SliderWidth - 15
         End If
         .Left = sldPos
         sldScale = ((picForm.Width - 20) - SliderWidth) / (Max - Min)
         value = (sldPos / sldScale) + Min
         For shp = 0 To 49
            If .Left > Shape1(shp).Left Then
               Shape1(shp).Visible = True
            Else
               Shape1(shp).Visible = False
               If value = Min Then Shape1(0).Visible = False
            End If
         Next shp
         If value > Max Then value = Max
         If value = Max Then
            Shape1(49).Visible = True
            value = Max
         End If
      End With
       'lblValue.Caption = Value
   End If
End Sub

Private Sub imgSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'lblValue.Visible = False  'True
End Sub

Private Sub imgSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If ValueHide = True Then
'      lblValue.Visible = False
'   Else
'      lblValue.Visible = False  'True
'   End If
End Sub

Private Sub picForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If ValueHide = True Then
'      lblValue.Visible = False
'   Else
'      lblValue.Visible = False 'True
'   End If
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim sldPos As Long
   Dim sldScale As Single
   Dim shp As Integer
 
   With imgSlider
      If Y > 400 Then
         sldPos = ((X - SliderWidth / 2) \ 15) * 15
         If sldPos < 0 Then
            sldPos = 0
         ElseIf sldPos > picForm.Width - SliderWidth Then
            sldPos = picForm.Width - SliderWidth - 15
         End If
         .Left = sldPos
         sldScale = ((picForm.Width - 20) - .Width) / (Max - Min)
         value = (sldPos / sldScale) + Min
         For shp = 0 To 49
          If value = Max Then Shape1(49).Visible = True
            If .Left > Shape1(shp).Left Then
               Shape1(shp).Visible = True
            Else
               Shape1(shp).Visible = False
               If value = Min Then Shape1(0).Visible = False
            End If
         Next shp
         If value > Max Then value = Max
         If value = Max Then
            Shape1(49).Visible = True
            value = Max
         End If
'         lblValue.Caption = Value
'         lblValue.Visible = False 'True
      End If
   End With
End Sub

Public Property Get BarColor() As Bar
    BarColor = m_BarColor
End Property

Public Property Let BarColor(NewBarColor As Bar)
Dim X As Integer
Dim Y As Integer

   m_BarColor = NewBarColor
   Select Case BarColor
      Case 0:    'Rainbow
         Shape1(0).FillColor = &H33B234              '&HC000&
         Shape1(1).FillColor = &H33B234          '&HC000&
         Shape1(2).FillColor = &H33B234          '&HC000&
         Shape1(3).FillColor = &H33B234          '&HC000&
         Shape1(4).FillColor = &H33B234          '&HC000&
         Shape1(5).FillColor = &H33B234          '&HFF00&
         Shape1(6).FillColor = &HC000&             '&HFF00&
         Shape1(7).FillColor = &HC000&          '&HFF00&
         Shape1(8).FillColor = &HC000&          '&HFF00&
         Shape1(9).FillColor = &H3DD171         '&HFF00&
         Shape1(10).FillColor = &H3DD171        '&HC0FF00
         Shape1(11).FillColor = &H3DD171         '&HC0FF00
         Shape1(12).FillColor = &H58D886          '&HC0FF00
         Shape1(13).FillColor = &H58D886
         Shape1(14).FillColor = &H58D886
         Shape1(15).FillColor = &H57D786
         Shape1(16).FillColor = &H57D786
         Shape1(17).FillColor = &H57D786
         Shape1(18).FillColor = &H5CDC9B
         Shape1(19).FillColor = &H5CDC9B
         Shape1(20).FillColor = &H5CDC9B
         Shape1(21).FillColor = &H63E3BA
         Shape1(22).FillColor = &H63E3BA
         Shape1(23).FillColor = &H63E3BA
         Shape1(24).FillColor = &H67E8C9
         Shape1(25).FillColor = &H67E8C9                 '&HC000&
         Shape1(26).FillColor = &H67E8C9          '&HC000&
         Shape1(27).FillColor = &H6BEBD8          '&HC000&
         Shape1(28).FillColor = &H6BEBD8          '&HC000&
         Shape1(29).FillColor = &H6BEBD8          '&HC000&
         Shape1(30).FillColor = &H6EEFE8          '&HFF00&
         Shape1(31).FillColor = &H6EEFE8         '&HFF00&
         Shape1(32).FillColor = &H6EEFE8         '&HFF00&
         Shape1(33).FillColor = &HFFF5&          '&HFF00&
         Shape1(34).FillColor = &HFFF5&          '&HFF00&
         Shape1(35).FillColor = &HFFF5&         '&HC0FF00
         Shape1(36).FillColor = &HE7FF&          '&HC0FF00
         Shape1(37).FillColor = &HE7FF&           '&HC0FF00
         Shape1(38).FillColor = &HE7FF&
         Shape1(39).FillColor = &HCAFF&
         Shape1(40).FillColor = &HCAFF&
         Shape1(41).FillColor = &HCAFF&
         Shape1(42).FillColor = &H87FF&
         Shape1(43).FillColor = &H87FF&
         Shape1(44).FillColor = &H87FF&
         Shape1(45).FillColor = &H3FFF&
         Shape1(46).FillColor = &H3FFF&
         Shape1(47).FillColor = &HFF&
         Shape1(48).FillColor = &HFF
         Shape1(49).FillColor = &HFF
         
         
       Case 1:    'Red
           For X = 0 To 49
           Y = Y + 5
               Shape1(X).FillColor = RGB(170 + Y, 50, 50)
               'Shape1(X).FillColor = RGB(255 - Y, 50, 50)
           Next X
       Case 2:   'Blue
'           For X = 0 To 49
'            Y = Y + 5
'               Shape1(X).FillColor = RGB(40, 90, 150 + Y)
'               'Shape1(X).FillColor = RGB(50, 50, 255 - Y)
'           Next X
           For X = 0 To 49
            If X < 42 Then
              Shape1(X).FillColor = &HEB7D58
            Else
              Shape1(X).FillColor = vbRed
            End If
           Next X
       Case 3:   'Green
'           For X = 0 To 49
'            Y = Y + 5
'               Shape1(X).FillColor = RGB(50, 150 + Y, 50)
'           Next X
           For X = 0 To 49
            If X < 42 Then
              Shape1(X).FillColor = &HC000&
            Else
              Shape1(X).FillColor = vbRed
            End If
           Next X
       Case 4:  'Yellow
           For X = 0 To 49
            Y = Y + 5
               Shape1(X).FillColor = RGB(160 + Y, 160 + Y, 50)
               'Shape1(X).FillColor = RGB(255 - Y, 255 - Y, 50)
           Next X
       Case 5:               'Purple
           For X = 0 To 49
                Y = Y + 5
               Shape1(X).FillColor = RGB(170 + Y, 50, 170 + Y)
               'Shape1(X).FillColor = RGB(255 - Y, 50, 255 - Y)
           Next X
       Case 6:    'Turq
'           For X = 0 To 49
'                Y = Y + 5
'               Shape1(X).FillColor = RGB(50, 170 + Y, 170 + Y)
'               'Shape1(X).FillColor = RGB(50, 255 - Y, 255 - Y)
'           Next X
           For X = 0 To 49
            If X < 42 Then
              Shape1(X).FillColor = &HC0C000
            Else
              Shape1(X).FillColor = vbRed
            End If
           Next X

       Case 7:    'Gray
           For X = 0 To 49
                Y = Y + 5
               Shape1(X).FillColor = RGB(140 + Y, 140 + Y, 140 + Y)
               'Shape1(X).FillColor = RGB(220 - Y, 220 - Y, 220 - Y)
           Next X
       End Select
   PropertyChanged "BarColor"
End Property

'''Public Property Get DropDownCtrl() As Boolean
'''   DropDownCtrl = m_DropDownCtrl
'''End Property
'''
'''Public Property Let DropDownCtrl(NewDropDownCtrl As Boolean)
'''   m_DropDownCtrl = NewDropDownCtrl
''''   If DropDownCtrl = True Then
''''      If DisableDropDown = True Then
''''         Image1.Visible = False
''''      Else
''''         Image1.Visible = True
''''      End If
''''      Image2.Visible = False
''''   Else
''''      Image1.Visible = False
''''      Image2.Visible = True
''''   End If
''''   PropertyChanged "DropDownCtrl"
''''   UserControl_Resize
'''End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
   m_Caption = NewCaption
   'lblCaption.Caption = Caption
   PropertyChanged "Caption"
End Property

'Public Property Get DisableDropDown() As Boolean
'   DisableDropDown = m_DisableDropDown
'End Property
'
'Public Property Let DisableDropDown(NewDisableDropDown As Boolean)
''   m_DisableDropDown = NewDisableDropDown
''   If m_DisableDropDown = True Then
''      Image1.Visible = False
''      Image2.Visible = False
''      DropDownCtrl = True
''   Else
''      If DropDownCtrl = True Then
''         Image1.Visible = True
''         Image2.Visible = False
''      Else
''         Image1.Visible = False
''         Image2.Visible = True
''      End If
''   End If
''   PropertyChanged "DisableDropDown"
'End Property

Public Property Get Font() As Font
     Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
     Set UserControl.Font = NewFont
     PropertyChanged "Font"
     ' show changes while in IDE
'     With lblValue
'        .FontSize = UserControl.FontSize
'        .FontBold = UserControl.FontBold
'        .Font = UserControl.Font
'      End With
'      With lblCaption
'        .FontSize = UserControl.FontSize
'        .FontBold = UserControl.FontBold
'        .Font = UserControl.Font
'      End With
End Property

Public Property Get Min() As Long
   Min = m_Min
End Property

Public Property Let Min(NewMin As Long)
'   If NewMin = Max Then
'      MsgBox "Sorry but Min and Max cannot be equal.  " & Caption, vbOKOnly, "Please correct error in " & Caption
'      Exit Property
'   End If
'   If NewMin > Max Then
'      MsgBox "Sorry but Min is greater then Max.  " & Caption, vbOKOnly, "Please correct error in " & Caption
'      Exit Property
'   End If
   m_Min = NewMin
   PropertyChanged "Min"
   SliderPos
End Property

Public Property Get Max() As Long
   Max = m_Max
End Property

Public Property Let Max(NewMax As Long)
'   If NewMax = Min Then
'      MsgBox "Sorry but Max and Min cannot be equal.  " & Caption, vbOKOnly, "Please correct error in " & Caption
'      Exit Property
'   End If
'   If NewMax < Min Then
'      MsgBox "Sorry but Max is less than Min.  " & Caption, vbOKOnly, "Please correct error in " & Caption
'      Exit Property
'   End If
   m_Max = NewMax
   PropertyChanged "Max"
   SliderPos
End Property

Public Property Get value() As Long
   value = m_Value
End Property

Public Property Let value(NewValue As Long)
On Error Resume Next
   m_Value = NewValue
   If m_Value > Max Then m_Value = Max
   If m_Value < Min Then m_Value = Min
   PropertyChanged "Value"
   SliderPos
   
End Property

'Public Property Get ValueHide() As Boolean
'   Let ValueHide = m_ValueHide
'End Property
'
'Public Property Let ValueHide(ByVal NewValueHide As Boolean)
'   Let m_ValueHide = NewValueHide
''   If m_ValueHide = True Then
''      lblValue.Visible = False
''   Else
''      lblValue.Visible = False ' True
''   End If
'   PropertyChanged "ValueHide"
'End Property

Public Property Get SliderHide() As Boolean
   Let SliderHide = m_ValueHide
End Property

Public Property Let SliderHide(ByVal NewValueHide As Boolean)
Dim i As Integer
   Let m_ValueHide = NewValueHide
   If m_ValueHide = True Then
      imgSlider.Visible = False
      picForm.Height = 135
      For i = 0 To 49
       Shape1(i).Height = 120
       Shape1(i).top = 5
      Next i
   Else
      imgSlider.Visible = True
      picForm.Height = 255
      For i = 0 To 49
       Shape1(i).Height = 120
       Shape1(i).top = 75
      Next i
   
   End If
   PropertyChanged "ValueHide"
End Property


Public Property Get TickColor() As OLE_COLOR
   TickColor = m_TickColor
End Property

Public Property Let TickColor(NewTickColor As OLE_COLOR)
   Dim X As Integer
   
'   m_TickColor = NewTickColor
'   lblValue.ForeColor = m_TickColor
'   Line1.BorderColor = m_TickColor
'   Line2.BorderColor = m_TickColor
'   Line3.BorderColor = m_TickColor
'   Line4.BorderColor = m_TickColor
'   Line5.BorderColor = m_TickColor
'   Line6.BorderColor = m_TickColor
   'if you want bar border color to change then uncomment following lines
  ' For x = 0 To 49
  '     Shape1(x).BorderColor = m_TickColor
   'Next x
   PropertyChanged "TickColor"
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(NewBackColor As OLE_COLOR)
   m_BackColor = NewBackColor
   picForm.BackColor = m_BackColor
   PropertyChanged "BackColor"
End Property

Public Property Get CapBkgdColor() As OLE_COLOR
   CapBkgdColor = m_CapBkgdColor
End Property

Public Property Let CapBkgdColor(NewCapBkgdColor As OLE_COLOR)
   m_CapBkgdColor = NewCapBkgdColor
  ' lblCaption.BackColor = m_CapBkgdColor
   PropertyChanged "CapBkgdColor"
End Property

Public Property Get CapFontColor() As OLE_COLOR
   CapFontColor = m_CapFontColor
End Property

Public Property Let CapFontColor(NewCapFontColor As OLE_COLOR)
   m_CapFontColor = NewCapFontColor
   'lblCaption.ForeColor = m_CapFontColor
   PropertyChanged "CapFontColor"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   
  ' DropDownCtrl = PropBag.ReadProperty("DropDownCtrl", m_def_DropDownCtrl)
  ' Caption = PropBag.ReadProperty("Caption", m_def_Caption)
  ' DisableDropDown = PropBag.ReadProperty("DisableDropDown", m_def_DisableDropDown)
   Min = PropBag.ReadProperty("Min", m_def_Min)
   Max = PropBag.ReadProperty("Max", m_def_Max)
   value = PropBag.ReadProperty("Value", m_def_Value)
   Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
   TickColor = PropBag.ReadProperty("TickColor", m_def_TickColor)
   BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
   CapBkgdColor = PropBag.ReadProperty("CapBkgdColor", m_def_CapBkgdColor)
   CapFontColor = PropBag.ReadProperty("CapFontColor", m_def_CapFontColor)
   BarColor = PropBag.ReadProperty("BarColor", m_def_BarColor)
   SliderHide = PropBag.ReadProperty("ValueHide", m_def_ValueHide)
'     With lblValue
'        .FontSize = UserControl.FontSize
'        .FontBold = UserControl.FontBold
'        .Font = UserControl.Font
'      End With
'      With lblCaption
'        .FontSize = UserControl.FontSize
'        .FontBold = UserControl.FontBold
'        .Font = UserControl.Font
'      End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      Call .WriteProperty("DropDownCtrl", m_DropDownCtrl, m_def_DropDownCtrl)
      Call .WriteProperty("Caption", m_Caption, m_def_Caption)
      Call .WriteProperty("DisableDropDown", m_DisableDropDown, m_def_DisableDropDown)
      Call .WriteProperty("Min", m_Min, m_def_Min)
      Call .WriteProperty("Max", m_Max, m_def_Max)
      Call .WriteProperty("Value", m_Value, m_def_Value)
      Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
      Call .WriteProperty("TickColor", m_TickColor, m_def_TickColor)
      Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
      Call .WriteProperty("CapBkgdColor", m_CapBkgdColor, m_def_CapBkgdColor)
      Call .WriteProperty("CapFontColor", m_CapFontColor, m_def_CapFontColor)
      Call .WriteProperty("BarColor", m_BarColor, m_def_BarColor)
      Call .WriteProperty("ValueHide", m_ValueHide, m_def_ValueHide)
   End With
End Sub

Private Function SliderPos()
Dim sldScale  As Single
Dim shp As Integer

On Error Resume Next
   
With imgSlider
  If Max - Min <> 0 Then
    sldScale = (picForm.Width - SliderWidth) / (Max - Min)
    .Left = (value - Min) * sldScale - 3
  End If
  For shp = 0 To 49
    If .Left + 15 > Shape1(shp).Left Then
      Shape1(shp).Visible = True
    Else
      Shape1(shp).Visible = False
      If value = Min Then Shape1(0).Visible = False
    End If
  Next shp
End With
   ' lblValue.Caption = Value
End Function
