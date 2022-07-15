Attribute VB_Name = "SLD60MDL"
'**********    VBSlider v1.1 for Visual Basic 3.0-6.0    **********
'*******                                                     ******
'****         Property of WolfeByte Solutions 1995-2002        ****
'**                                                              **
'**      This program is protected by and subject to all         **
'**    Federal copyright laws governing the duplication and      **
'**  distribution of authored software.  With the purchase and   **
'** use of this program you agree to release WolfeByte Solutions **
'**  of all liability and/or damages as related to the use of    **
'**    this program and also acknowledge that no claims or       **
'**      warranties regarding its usage have been offered.       **
'****                                                          ****
'*******                                                     ******
'**********      Release Version 1.1  April 28, 1996     **********

Option Explicit
Option Base 1

'All subs are in this module.  The main calls to these subs are in the
'form1_Load event and the picture1_MouseDown, MouseMove and MouseUp events.
    
'To set the sldCrntVal at run time use the following code (with Slider(x) in front of each)
'*** For Horizontal Sliders
'    sldCrntVal = your_value
'    sldCrntPos = (sldCrntVal - sldMinVal) / ((sldMaxVal - sldMinVal) / sldNumDiv)
'    DrawButton Slider(x), picture1(x)
    
'*** For Vertical Sliders
'    sldCrntVal = your_value
'    sldCrntPos = (sldMaxVal - sldCrntVal) / ((sldMaxVal - sldMinVal) / sldNumDiv)
'    DrawButton Slider(x), picture1(x)

'Code is also included (optional) in the Picture1_KeyDown event to move the
'button on the slider when the arrow keys are pressed.

'Make a record structure for font types
Type ftType
  ftName As String
  ftSize As Integer
  ftBold As Integer
  ftItalic As Integer
  ftColor  As Long
End Type

'Make a record structure for a Slider - one time only
Type sldProps
    sldOrient As Integer
    sld3D As Integer
    sldBevel As Integer
    sldInBrdr  As Integer
    sldColor  As Long
    btn3D As Integer
    btnBevel As Integer
    btnInBrdr As Integer
    btnColor  As Long
    btnHght As Integer
    btnWdth As Integer
    btnMrkSnap As Integer
    sldNumDiv  As Integer
    sldMaxVal As Single
    sldMinVal As Single
    sldFont As ftType
    sldLftTxt As String
    sldRgtTxt As String
    sldGtrLgth As Single
    sldCrntVal  As Single
    sldCrntPos As Single
    sldCrntMove As Integer
End Type

'Declare array for number of Sliders you want
Global Slider(3) As sldProps

Global p1%   'Will be set to the twipsperpixel * 1
Global p2%   'Will be set to the twipsperpixel * 2
Global p3%   'Will be set to the twipsperpixel * 3

Sub AdjustButton(inSld As sldProps, inPct As PictureBox, inX As Single, inY As Single)

'Called from the MouseMove and MouseUp events to redraw the button position
     
     Dim gtrPct!, mrkWdth!, mrkHght!, mrkVal!

'Redraw the button if mouse is over the Slider area based on orientation
     Select Case inSld.sldOrient
        Case 1                  'Horizontal
            'Calculate the gutter distance from edge and the mark width and the value of each mark
            gtrPct = inPct.ScaleWidth - (inPct.ScaleWidth * inSld.sldGtrLgth)
            mrkWdth = (inPct.ScaleWidth - (gtrPct * 2)) / inSld.sldNumDiv
            mrkVal = (inSld.sldMaxVal - inSld.sldMinVal) / inSld.sldNumDiv

            If inY > (inPct.ScaleHeight / 2) - (inSld.btnHght / 2) - inSld.btnBevel And inY < (inPct.ScaleHeight / 2) + (inSld.btnHght / 2) Then
               'If x is left of gutter then place at min OR if x is to right then place at max
               If inX < gtrPct + p1 Then
                  inX = gtrPct
               ElseIf inX > inPct.ScaleWidth - gtrPct Then
                  inX = inPct.ScaleWidth - gtrPct
               End If
               
               If inSld.btnMrkSnap Then       'If snap to click true then round the x to the closest mark
                  inSld.sldCrntPos = CInt((inX - gtrPct) / mrkWdth)
               Else
                  inSld.sldCrntPos = (inX - gtrPct) / mrkWdth
               End If
                
               inSld.sldCrntVal = (inSld.sldCrntPos * mrkVal) + inSld.sldMinVal
               DrawButton inSld, inPct
            End If
        Case 2                  'Vertical
            'Calculate the gutter distance from edge and the mark width and the value of each mark
            gtrPct = inPct.ScaleHeight - (inPct.ScaleHeight * inSld.sldGtrLgth)
            mrkHght = (inPct.ScaleHeight - (gtrPct * 2)) / inSld.sldNumDiv
            mrkVal = (inSld.sldMaxVal - inSld.sldMinVal) / inSld.sldNumDiv

            If inX > (inPct.ScaleWidth / 2) - (inSld.btnWdth / 2) - inSld.btnBevel And inX < (inPct.ScaleWidth / 2) + (inSld.btnWdth / 2) Then
               'If y is above gutter then place at max OR if y is below then place at min
               If inY < gtrPct + p1 Then
                  inY = gtrPct
               ElseIf inY > inPct.ScaleHeight - gtrPct Then
                  inY = inPct.ScaleHeight - gtrPct
               End If
               
               If inSld.btnMrkSnap Then       'If snap to click true then round the x to the closest mark
                  inSld.sldCrntPos = CInt((inY - gtrPct) / mrkHght)
               Else
                  inSld.sldCrntPos = (inY - gtrPct) / mrkHght
               End If
                
               inSld.sldCrntVal = inSld.sldMaxVal - (inSld.sldCrntPos * mrkVal)
               DrawButton inSld, inPct
            End If
     
     End Select

End Sub

Sub ClickSlider(inSld As sldProps, inPct As PictureBox, inX As Single, inY As Single)

'Called from the mouse down to see if the mouse is pointing to the Slider
'area - if so then start tracking the mouse move.
     
'Start mouse tracking for mouse move depending on slider orientation.
     Select Case inSld.sldOrient
        Case 1                  'Horizontal
            'X value can be to the edge of the slider area - easier to 'hit' the end on a move.
            If inY > (inPct.ScaleHeight / 2) - (inSld.btnHght / 2) - inSld.btnBevel And inY < (inPct.ScaleHeight / 2) + (inSld.btnHght / 2) And inX > inSld.sldBevel And inX < inPct.ScaleWidth - inSld.sldBevel Then
                inSld.sldCrntMove = True
            End If
        Case 2                  'Vertical
            'Y value can be to the edge of the slider area - easier to 'hit' the end on a move.
            If inX > (inPct.ScaleWidth / 2) - (inSld.btnWdth / 2) - inSld.btnBevel And inX < (inPct.ScaleWidth / 2) + (inSld.btnWdth / 2) And inY > inSld.sldBevel And inY < inPct.ScaleHeight - inSld.sldBevel Then
                inSld.sldCrntMove = True
            End If
     End Select

End Sub

Sub DefineSlider1(inPct As PictureBox)

'This sub needs to be called once from the form load event so that the
'Slider(1) is assigned its properties.  This could be in one sub as an
'array but I separate them so they are easier to locate and read - create
'one for each Slider you have - DefineSlider2..3...
'You also have to increase the number of elements in the Slider array which
'is dimensioned in the declarations section - Global Slider(x) as sldProps.
'Prepare necessary settings for picturebox
    inPct.AutoRedraw = True                 'Allow creation of pct in memory
    inPct.ScaleMode = 1                     'Draw in twips mode
    p1 = Screen.TwipsPerPixelX              'This is how wide one drawing line will be - used for bevels and some line drawing
    p2 = p1 * 2                             'Width of two lines
    p3 = p1 * 3                             'Width of three lines (add more if needed)

'Properties for Slider1
    Slider(1).sldOrient = 1            '1-Horizontal, 2-Vertical
    Slider(1).sld3D = 0                '0-none, 1-raised, 2-sunken
    Slider(1).sldBevel = 0             'Must be 0 if 3D=0
    Slider(1).sldInBrdr = False        'Border inside of bevel - can be used even if 3D=0 and Bevel=0
    Slider(1).sldColor = CLng(&H2F2F2F)      'QBColor(7)    'Slider background color
    Slider(1).btn3D = 1
    Slider(1).btnBevel = p1
    Slider(1).btnInBrdr = True
    Slider(1).btnColor = vbCyan 'QBColor(9)    'Button color
    Slider(1).btnHght = 15 * p1        'Button height
    Slider(1).btnWdth = 15 * p1        'Button width
    Slider(1).btnMrkSnap = False        'Button stop on marks only (True) or smooth (False)
    Slider(1).sldNumDiv = 2           'Number of division segments
    Slider(1).sldMaxVal = 100           'Maximum slide value
    Slider(1).sldMinVal = 0           'Minimum slide value
    Slider(1).sldFont.ftName = "MS Sans Serif"
    Slider(1).sldFont.ftSize = 6  '8.25
    Slider(1).sldFont.ftBold = False
    Slider(1).sldFont.ftItalic = False
    Slider(1).sldFont.ftColor = CLng(&HFF00&)      'vbWhite  'QBColor(0)
    Slider(1).sldLftTxt = "0"        'Left or bottom mark text
    Slider(1).sldRgtTxt = "100"       'Right or top mark text
    Slider(1).sldGtrLgth = 0.85        'Length of center bar as a percent of total width or height
    Slider(1).sldCrntVal = 0          'Current (starting) pointer value
    Slider(1).sldCrntPos = 1           'Current (starting) pointer position
    Slider(1).sldCrntMove = False      'Initial mouse tracking off


'''''This sub needs to be called once from the form load event so that the
'''''Slider(1) is assigned its properties.  This could be in one sub as an
'''''array but I separate them so they are easier to locate and read - create
'''''one for each Slider you have - DefineSlider2..3...
'''''You also have to increase the number of elements in the Slider array which
'''''is dimensioned in the declarations section - Global Slider(x) as sldProps.
''''
'''''Prepare necessary settings for picturebox
''''    inPct.AutoRedraw = True                 'Allow creation of pct in memory
''''    inPct.ScaleMode = 1                     'Draw in twips mode
''''    p1 = Screen.TwipsPerPixelX              'This is how wide one drawing line will be - used for bevels and some line drawing
''''    p2 = p1 * 2                             'Width of two lines
''''    p3 = p1 * 3                             'Width of three lines (add more if needed)
''''
'''''Properties for Slider1
''''    Slider(1).sldOrient = 1            '1-Horizontal, 2-Vertical
''''    Slider(1).sld3D = 0                '0-none, 1-raised, 2-sunken
''''    Slider(1).sldBevel = 0             'Must be 0 if 3D=0
''''    Slider(1).sldInBrdr = False        'Border inside of bevel - can be used even if 3D=0 and Bevel=0
''''    Slider(1).sldColor = QBColor(7)    'Slider background color
''''    Slider(1).btn3D = 1
''''    Slider(1).btnBevel = p2
''''    Slider(1).btnInBrdr = True
''''    Slider(1).btnColor = QBColor(7)    'Button color
''''    Slider(1).btnHght = 20 * p1        'Button height
''''    Slider(1).btnWdth = 10 * p1        'Button width
''''    Slider(1).btnMrkSnap = True        'Button stop on marks only (True) or smooth (False)
''''    Slider(1).sldNumDiv = 10           'Number of division segments
''''    Slider(1).sldMaxVal = 93           'Maximum slide value
''''    Slider(1).sldMinVal = 18           'Minimum slide value
''''    Slider(1).sldFont.ftName = "MS Sans Serif"
''''    Slider(1).sldFont.ftSize = 8.25
''''    Slider(1).sldFont.ftBold = False
''''    Slider(1).sldFont.ftItalic = False
''''    Slider(1).sldFont.ftColor = QBColor(0)
''''    Slider(1).sldLftTxt = "Low"        'Left or bottom mark text
''''    Slider(1).sldRgtTxt = "High"       'Right or top mark text
''''    Slider(1).sldGtrLgth = 0.85        'Length of center bar as a percent of total width or height
''''    Slider(1).sldCrntVal = 33          'Current (starting) pointer value
''''    Slider(1).sldCrntPos = 2           'Current (starting) pointer position
''''    Slider(1).sldCrntMove = False      'Initial mouse tracking off

End Sub

Sub DefineSlider2(inPct As PictureBox)

'*** See DefineSlider1 for descriptions

    inPct.AutoRedraw = True
    inPct.ScaleMode = 1
    p1 = Screen.TwipsPerPixelX
    p2 = p1 * 2
    p3 = p1 * 3

'Properties for Slider1
    Slider(2).sldOrient = 2
    Slider(2).sld3D = 1
    Slider(2).sldBevel = p2
    Slider(2).sldInBrdr = False
    Slider(2).sldColor = QBColor(3)
    Slider(2).btn3D = 1
    Slider(2).btnBevel = p1
    Slider(2).btnInBrdr = False
    Slider(2).btnColor = QBColor(4)
    Slider(2).btnHght = 11 * p1
    Slider(2).btnWdth = 17 * p1
    Slider(2).btnMrkSnap = False
    Slider(2).sldNumDiv = 4
    Slider(2).sldMaxVal = 40
    Slider(2).sldMinVal = 0
    Slider(2).sldFont.ftName = "MS Sans Serif"
    Slider(2).sldFont.ftSize = 8.25
    Slider(2).sldFont.ftBold = False
    Slider(2).sldFont.ftItalic = False
    Slider(2).sldFont.ftColor = QBColor(14)
    Slider(2).sldLftTxt = "Min"
    Slider(2).sldRgtTxt = "Max"
    Slider(2).sldGtrLgth = 0.85
    Slider(2).sldCrntVal = 20
    Slider(2).sldCrntPos = 2
    Slider(2).sldCrntMove = False

End Sub

Sub DefineSlider3(inPct As PictureBox)

'*** See DefineSlider1 for descriptions

    inPct.AutoRedraw = True
    inPct.ScaleMode = 1
    p1 = Screen.TwipsPerPixelX
    p2 = p1 * 2
    p3 = p1 * 3

'Properties for Slider1
    Slider(3).sldOrient = 1
    Slider(3).sld3D = 2
    Slider(3).sldBevel = p1
    Slider(3).sldInBrdr = True
    Slider(3).sldColor = QBColor(7)
    Slider(3).btn3D = 2
    Slider(3).btnBevel = p1
    Slider(3).btnInBrdr = True
    Slider(3).btnColor = QBColor(9)
    Slider(3).btnHght = 20 * p1
    Slider(3).btnWdth = 6 * p1
    Slider(3).btnMrkSnap = False
    Slider(3).sldNumDiv = 2
    Slider(3).sldMaxVal = 50
    Slider(3).sldMinVal = -50
    Slider(3).sldFont.ftName = "MS Sans Serif"
    Slider(3).sldFont.ftSize = 8.25
    Slider(3).sldFont.ftBold = False
    Slider(3).sldFont.ftItalic = True
    Slider(3).sldFont.ftColor = QBColor(9)
    Slider(3).sldLftTxt = "Neg."
    Slider(3).sldRgtTxt = "Pos."
    Slider(3).sldGtrLgth = 0.8
    Slider(3).sldCrntVal = 0
    Slider(3).sldCrntPos = 1
    Slider(3).sldCrntMove = False

End Sub

Sub DrawButton(inSld As sldProps, inPct As PictureBox)

'Called once at load from the DrawSlider.  Called from the AdjustButton sub
'each time the mouse moves.  This will redraw the the button on the Slider.

     Dim gtrPct!, mrkWdth!, mrkHght!
     
'Draw button and gutter line base on slider orientation
     Select Case inSld.sldOrient
        Case 1                  'Horizontal
          'Calculate the gutter distance from edge
           gtrPct = inPct.ScaleWidth - (inPct.ScaleWidth * inSld.sldGtrLgth)
           mrkWdth = (inPct.ScaleWidth - (gtrPct * 2)) / inSld.sldNumDiv
           
           'First redraw the current button (entire button slide area)
           DrawSldBoxes inPct, gtrPct - (inSld.btnWdth / 2) - p1, (inPct.ScaleHeight / 2) - Int(inSld.btnHght / 2) - p1, inPct.ScaleWidth - gtrPct + (inSld.btnWdth / 2) + p2, (inPct.ScaleHeight / 2) + Int(inSld.btnHght / 2) + p2, inSld.sldColor, 0, 0, False
           
           'Draw gutter line
           inPct.Line (gtrPct, Int(inPct.ScaleHeight / 2) - p1)-(inPct.ScaleWidth - gtrPct + p1, Int(inPct.ScaleHeight / 2) - p1), QBColor(0)
           inPct.Line (gtrPct, Int(inPct.ScaleHeight / 2))-(inPct.ScaleWidth - gtrPct + p1, Int(inPct.ScaleHeight / 2)), QBColor(15)
           
           'Draw the button on the slider based on current value
           DrawSldBoxes inPct, gtrPct + (inSld.sldCrntPos * mrkWdth) - (inSld.btnWdth / 2), (inPct.ScaleHeight / 2) - (inSld.btnHght / 2) - p1, gtrPct + (inSld.sldCrntPos * mrkWdth) + (inSld.btnWdth / 2), (inPct.ScaleHeight / 2) + (inSld.btnHght / 2) - p1, inSld.btnColor, inSld.btn3D, inSld.btnBevel, inSld.btnInBrdr
        
        Case 2                  'Vertical
          'Calculate the gutter distance from edge
           gtrPct = inPct.ScaleHeight - (inPct.ScaleHeight * inSld.sldGtrLgth)
           mrkHght = (inPct.ScaleHeight - (gtrPct * 2)) / inSld.sldNumDiv
           
           'First redraw the current button (entire button slide area)
           DrawSldBoxes inPct, (inPct.ScaleWidth / 2) - (inSld.btnWdth / 2) - p3, gtrPct - (inSld.btnHght / 2) - p3, (inPct.ScaleWidth / 2) + (inSld.btnWdth / 2) + p2, inPct.ScaleHeight - gtrPct + (inSld.btnHght / 2) + p2, inSld.sldColor, 0, 0, False
           
           'Draw gutter line
           inPct.Line (Int(inPct.ScaleWidth / 2) - p1, gtrPct)-(Int(inPct.ScaleWidth / 2) - p1, inPct.ScaleHeight - gtrPct + p1), QBColor(0)
           inPct.Line (Int(inPct.ScaleWidth / 2), gtrPct)-(Int(inPct.ScaleWidth / 2), inPct.ScaleHeight - gtrPct + p1), QBColor(15)
           
           'Draw the button on the slider based on current value
           DrawSldBoxes inPct, (inPct.ScaleWidth / 2) - (inSld.btnWdth / 2) - p1, gtrPct + (inSld.sldCrntPos * mrkHght) - (inSld.btnHght / 2), (inPct.ScaleWidth / 2) + (inSld.btnWdth / 2), gtrPct + (inSld.sldCrntPos * mrkHght) + (inSld.btnHght / 2), inSld.btnColor, inSld.btn3D, inSld.btnBevel, inSld.btnInBrdr
           
     End Select

End Sub

Sub DrawSldBoxes(inPct As PictureBox, lbLeft%, lbTop%, lbRight%, lbBottom%, lbBackColor&, lb3D%, lbBevel%, lbInBrdr%)


'Called from within this module to make box areas and/or beveling

  Dim X%, Y%
    
'Prepare picturebox settings needed for drawing
   inPct.FillColor = lbBackColor     'Box will be black and filled with this
   inPct.FillStyle = 0               'Fill in as solid
   inPct.DrawStyle = 0               'Drawing outline will be solid
   inPct.DrawWidth = 1               'Units will be 1 line wide
  
'Print box on picturebox and then a border inside if needed
   If lbInBrdr Then
      inPct.Line (lbLeft + lbBevel, lbTop + lbBevel)-(lbRight - lbBevel, lbBottom - lbBevel), 0, B
   Else
      inPct.Line (lbLeft, lbTop)-(lbRight, lbBottom), lbBackColor, B
   End If

'Draw beveling around the box - Bevel line lengths are incremented via loop
'to make 45 degree corners and are drawn inward (inside the picturebox)
   Select Case lb3D
     Case 1                 'raised beveling
       For X = 0 To (lbBevel - p1) / p1
         Y = X * p1
         inPct.Line (lbLeft + Y, lbBottom - Y)-(lbRight - Y + p1, lbBottom - Y), RGB(92, 92, 92)
         inPct.Line (lbLeft + Y, lbTop + Y)-(lbRight - Y + p1, lbTop + Y), RGB(255, 255, 255)
         inPct.Line (lbRight - Y, lbTop + Y)-(lbRight - Y, lbBottom - Y), RGB(92, 92, 92)
         inPct.Line (lbLeft + Y, lbTop + Y)-(lbLeft + Y, lbBottom - Y), RGB(255, 255, 255)
       Next
     Case 2                 'sunken beveling
       For X = 0 To (lbBevel - p1) / p1
         Y = X * p1
         inPct.Line (lbLeft + Y, lbBottom - Y)-(lbRight - Y + p1, lbBottom - Y), RGB(255, 255, 255)
         inPct.Line (lbLeft + Y, lbTop + Y)-(lbRight - Y + p1, lbTop + Y), RGB(92, 92, 92)
         inPct.Line (lbRight - Y, lbTop + Y)-(lbRight - Y, lbBottom - Y), RGB(255, 255, 255)
         inPct.Line (lbLeft + Y, lbTop + Y)-(lbLeft + Y, lbBottom - Y), RGB(92, 92, 92)
       Next
     End Select

End Sub

Sub DrawSldText(inPct As PictureBox, lbLeft%, lbTop%, lbRight%, lbBottom%, lbText$, lbHorzAlign$, lbVertAlign!)

'Called from within this module to draw text on the picturebox and align
'accordingly within area passed in.  The vertical alignment of the text is
'a percent of distance from the top of the area - pass in one of
'these -> 0 = top, .5 = middle, 1 = bottom.
  
  Select Case LCase(lbHorzAlign)
    Case "left"
      inPct.CurrentX = lbLeft
      inPct.CurrentY = lbTop + (((lbBottom - lbTop) * lbVertAlign) - (inPct.TextHeight(lbText) * lbVertAlign))
    Case "right"
      inPct.CurrentX = lbRight - inPct.TextWidth(lbText)
      inPct.CurrentY = lbTop + (((lbBottom - lbTop) * lbVertAlign) - (inPct.TextHeight(lbText) * lbVertAlign))
    Case "center"
      inPct.CurrentX = lbLeft + (((lbRight - lbLeft) / 2) - (inPct.TextWidth(lbText) / 2))
      inPct.CurrentY = lbTop + (((lbBottom - lbTop) * lbVertAlign) - (inPct.TextHeight(lbText) * lbVertAlign))
  End Select
    
  inPct.Print lbText

End Sub

Sub DrawSlider(inSld As sldProps, inPct As PictureBox)

'Called from the form1_Load event to draw the Slider - drawn one time only.
    
    Dim X%, mrkWdth!, mrkHght!, gtrPct!

    'Draw outside area of the slider
    DrawSldBoxes inPct, 0, 0, inPct.ScaleWidth - p1, inPct.ScaleHeight - p1, inSld.sldColor, inSld.sld3D, inSld.sldBevel, inSld.sldInBrdr
  
    'Draw text labels and markers
    SetSldFonts inSld.sldFont, inPct
    Select Case inSld.sldOrient
       Case 1                    'Horizontal
          'Calculate the gutter distance from edge
          gtrPct = inPct.ScaleWidth - (inPct.ScaleWidth * inSld.sldGtrLgth)
          'Draw left and right text labels
          DrawSldText inPct, inSld.sldBevel + p1, inPct.ScaleHeight - inSld.sldBevel - inPct.TextHeight(inSld.sldLftTxt), gtrPct - (inSld.btnWdth / 2) - p2, inPct.ScaleHeight - inSld.sldBevel - p2, inSld.sldLftTxt, "right", 0.5
          DrawSldText inPct, inPct.ScaleWidth - gtrPct + (inSld.btnWdth / 2) + p3, inPct.ScaleHeight - inSld.sldBevel - inPct.TextHeight(inSld.sldLftTxt), inPct.ScaleWidth - inSld.sldBevel - p1, inPct.ScaleHeight - inSld.sldBevel - p2, inSld.sldRgtTxt, "left", 0.5
          
          'Draw tick lines - number of lines drawn will be 1 more than number of segments wanted (sldNumDiv)
          mrkWdth = (inPct.ScaleWidth - (gtrPct * 2)) / inSld.sldNumDiv
          inPct.Line (gtrPct, inPct.ScaleHeight - inSld.sldBevel - inPct.TextHeight(inSld.sldLftTxt))-(gtrPct, inPct.ScaleHeight - inSld.sldBevel - p2), inSld.sldFont.ftColor
          For X = 1 To inSld.sldNumDiv
             inPct.Line (gtrPct + (X * mrkWdth), inPct.ScaleHeight - inSld.sldBevel - inPct.TextHeight(inSld.sldLftTxt))-(gtrPct + (X * mrkWdth), inPct.ScaleHeight - inSld.sldBevel - p2), inSld.sldFont.ftColor
          Next X

       Case 2                    'Vertical
          'Calculate the gutter distance from edge
          gtrPct = inPct.ScaleHeight - (inPct.ScaleHeight * inSld.sldGtrLgth)
          'Draw bottom and top text labels
          DrawSldText inPct, inSld.sldBevel + p1, inPct.ScaleHeight - gtrPct + p2 + (inSld.btnHght / 2), inPct.ScaleWidth - inSld.sldBevel - p2, inPct.ScaleHeight - gtrPct + p2 + (inSld.btnHght / 2) + inPct.TextHeight(inSld.sldLftTxt) + p2, inSld.sldLftTxt, "right", 0.5
          DrawSldText inPct, inSld.sldBevel + p1, gtrPct - (inSld.btnHght / 2) - inPct.TextHeight(inSld.sldRgtTxt) - p2, inPct.ScaleWidth - inSld.sldBevel - p2, gtrPct - p2 - (inSld.btnHght / 2), inSld.sldRgtTxt, "right", 0.5
          
          'Draw tick lines - number of lines drawn will be 1 more than number of segments wanted (sldNumDiv)
          mrkHght = (inPct.ScaleHeight - (gtrPct * 2)) / inSld.sldNumDiv
          inPct.Line (inPct.ScaleWidth - inSld.sldBevel - (12 * p1), gtrPct)-(inPct.ScaleWidth - inSld.sldBevel - p2, gtrPct), inSld.sldFont.ftColor
          For X = 1 To inSld.sldNumDiv
             inPct.Line (inPct.ScaleWidth - inSld.sldBevel - (12 * p1), gtrPct + (X * mrkHght))-(inPct.ScaleWidth - inSld.sldBevel - p2, gtrPct + (X * mrkHght)), inSld.sldFont.ftColor
          Next X
    
    End Select
    
    'Draw button on the slider
    DrawButton inSld, inPct

End Sub

Sub SetSldFonts(inFont As ftType, inPct As PictureBox)
    
'Called from within this sub to set the fonts of the picturebox before
'printing any text.  The font name will not be checked - if you get an
'error here - check font spelling first.
    
    inPct.FontName = inFont.ftName
    inPct.FontSize = inFont.ftSize
    inPct.FontBold = inFont.ftBold
    inPct.FontItalic = inFont.ftItalic
    inPct.ForeColor = inFont.ftColor

End Sub

