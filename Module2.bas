Attribute VB_Name = "Module2"
Option Explicit
 
Public defWindowProc As Long
Public hSliderHwnd As Long
Private hSliderBGBrush As Long
 
Private Const WM_USER = &H400&
Private Const TBM_GETTOOLTIPS = (WM_USER + 30)
Private Const TTM_ACTIVATE = (WM_USER + 1)
 
Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_GETMINMAXINFO As Long = &H24
Private Const WM_TIMECHANGE = &H1E
Private Const WM_DESTROY = &H2
 
Private Const WM_CTLCOLORSTATIC = &H138
 
Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
 
Private Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, _
   ByVal uMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long
 
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
 
Private Declare Function CreateSolidBrush Lib "gdi32" _
  (ByVal crColor As Long) As Long
 
Private Declare Function DeleteObject Lib "gdi32" _
   (ByVal hObject As Long) As Long
   
'StatusBar
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR
 
'ProgressBar
Public Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Public Const PBM_SETBARCOLOR = (WM_USER + 9)
 
Private Declare Function OleTranslateColor Lib "olepro32" _
   (ByVal clr As OLE_COLOR, ByVal hpal As Long, _
   pcolorref As Long) As Long
   
Private Declare Function GetSysColor Lib "user32" _
      (ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal _
        ByteLen As Long)
 
 
Private Type OLECOLOR
   RedOrSys As Byte
   Green As Byte
   Blue As Byte
   Type As Byte
End Type
 
Function WinColor(VBColor As Long) As Long
   Dim SysClr As OLECOLOR
 
   CopyMemory SysClr, VBColor, Len(SysClr)
 
   If SysClr.Type = &H80 Then 'Es ist eine Systemfarbe
      WinColor = GetSysColor(SysClr.RedOrSys)
   Else 'Es ist keine Systemfarbe
      WinColor = VBColor
   End If
End Function

 
Public Function lR(ByVal Color As Long) As Byte
   CopyMemory lR, WinColor(Color), 1
End Function

 
Public Function lG(ByVal Color As Long) As Byte
   CopyMemory lG, ByVal VarPtr(WinColor(Color)) + 1, 1
End Function

 
Public Function lB(ByVal Color As Long) As Byte
   CopyMemory lB, ByVal VarPtr(WinColor(Color)) + 2, 1
End Function

 
Public Property Let SBBackColor(ByRef StatusBar As StatusBar, ByVal New_Value As OLE_COLOR)
 
  OleTranslateColor New_Value, 0, New_Value
 
  SendMessage StatusBar.hwnd, SB_SETBKCOLOR, 0, ByVal New_Value
End Property
 
Public Property Let PBBackColor(ByRef ProgBar As ProgressBar, ByVal nBackColor As OLE_COLOR)
 
  OleTranslateColor nBackColor, 0, nBackColor
  SendMessage ProgBar.hwnd, PBM_SETBKCOLOR, 0, ByVal nBackColor
End Property
 
Public Property Let PBBarColor(ByRef ProgBar As ProgressBar, ByVal nBarColor As OLE_COLOR)
 
  ' neue Vordergrundfarbe
  OleTranslateColor nBarColor, 0, nBarColor
  SendMessage ProgBar.hwnd, PBM_SETBARCOLOR, 0&, nBarColor
End Property
 
Public Sub CreateSliderBrush(clrref As Long, bReset As Boolean)
 
   If (hSliderBGBrush <> 0) Or (bReset = True) Then
      Call DeleteSliderBrush
   End If
 
   If hSliderBGBrush = 0 Then
      hSliderBGBrush = CreateSolidBrush(clrref)
   End If
 
End Sub
 
Public Sub DeleteSliderBrush()
 
   If (hSliderBGBrush <> 0) Then
      DeleteObject hSliderBGBrush
      hSliderBGBrush = 0
   End If
 
End Sub

 Public Sub SubClass(hwnd As Long)
 
   On Error Resume Next
   defWindowProc = SetWindowLong(hwnd, _
                                 GWL_WNDPROC, _
                                 AddressOf WindowProc)
 
End Sub
 
Public Sub UnSubClass(hwnd As Long)
 
   If defWindowProc Then
      SetWindowLong hwnd, GWL_WNDPROC, defWindowProc
      defWindowProc = 0
   End If
 
End Sub

 
Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
 
   Select Case hwnd
 
      Case frmPlayer.hwnd
 
         Select Case uMsg
 
            Case WM_CTLCOLORSTATIC
 
              If (lParam = hSliderHwnd) And (hSliderBGBrush <> 0) Then
 
                  WindowProc = hSliderBGBrush
                  Exit Function
 
               Else
 
                  WindowProc = CallWindowProc(defWindowProc, _
                                              hwnd, _
                                              uMsg, _
                                              wParam, _
                                              lParam)
 
                  Exit Function
 
               End If
 
 
            Case WM_DESTROY
 
               If (hSliderBGBrush <> 0) Then
                  Call DeleteSliderBrush
                  hSliderBGBrush = 0
               End If
 
               Call UnSubClass(hwnd)
 
            Case Else
 
               WindowProc = CallWindowProc(defWindowProc, _
                                            hwnd, _
                                            uMsg, _
                                            wParam, _
                                            lParam)
               Exit Function
 
          End Select
 
 
      Case Else
 
      WindowProc = CallWindowProc(defWindowProc, _
                                  hwnd, _
                                  uMsg, _
                                  wParam, _
                                  lParam)
   End Select
 
End Function


