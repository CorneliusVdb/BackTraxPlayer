VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowPalettes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "   "
   ClientHeight    =   11070
   ClientLeft      =   0
   ClientTop       =   18150
   ClientWidth     =   20445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmShowPalettes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmShowPalettes.frx":0442
   ScaleHeight     =   11070
   ScaleWidth      =   20445
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   8235
      Left            =   15330
      TabIndex        =   9
      Top             =   2400
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   14526
      _Version        =   131074
      BackColor       =   0
      Caption         =   "SSPanel2"
      BevelOuter      =   0
      Begin Threed.SSCommand cmdUp 
         Height          =   945
         Left            =   0
         TabIndex        =   11
         Top             =   90
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1667
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   0
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmShowPalettes.frx":3987
         AutoSize        =   1
         Alignment       =   8
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdDown 
         Height          =   945
         Left            =   30
         TabIndex        =   10
         Top             =   6555
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1667
         _Version        =   131074
         ForeColor       =   16777215
         BackColor       =   0
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Candara"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmShowPalettes.frx":41A9
         AutoSize        =   1
         Alignment       =   8
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
   End
   Begin Threed.SSPanel pnlHeading 
      Height          =   405
      Left            =   30
      TabIndex        =   3
      Top             =   1350
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   714
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   7104768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "                                                                               Available  Play  Lists"
      BorderWidth     =   1
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin MSComctlLib.ListView lvPalettes 
      Height          =   7380
      Left            =   5610
      TabIndex        =   1
      Top             =   2520
      Width           =   10000
      _ExtentX        =   17648
      _ExtentY        =   13018
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   4210752
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   19455
      Top             =   2310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowPalettes.frx":49C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowPalettes.frx":BBFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowPalettes.frx":14EDD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   885
      Left            =   18150
      TabIndex        =   5
      Top             =   180
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1561
      _Version        =   131074
      BackColor       =   3092271
      BevelOuter      =   0
      Begin Threed.SSCommand cmdExit 
         Height          =   810
         Left            =   975
         TabIndex        =   7
         Top             =   45
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1429
         _Version        =   131074
         ForeColor       =   15194953
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   810
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1429
         _Version        =   131074
         ForeColor       =   15194953
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OK"
         AutoSize        =   1
         ButtonStyle     =   3
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   9585
      Left            =   15
      TabIndex        =   0
      Top             =   1440
      Width           =   20490
      _ExtentX        =   36142
      _ExtentY        =   16907
      _Version        =   131074
      BackColor       =   0
      Caption         =   "SSPanel1"
      BevelOuter      =   0
      Begin VB.TextBox txtPalette 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6855
         MaxLength       =   50
         TabIndex        =   2
         Top             =   435
         Width           =   8205
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H006C6900&
         Caption         =   "Play List Name :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   4695
         TabIndex        =   4
         Top             =   450
         Width           =   2145
      End
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   990
      TabIndex        =   8
      Top             =   720
      Width           =   2205
   End
End
Attribute VB_Name = "frmShowPalettes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPreviousSelection As Integer
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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Sub LvScrollDown()
SetItemFocusA lvPalettes, 16
End Sub

Sub LvScrollUp()
SetItemFocusA lvPalettes, 1
End Sub

Function SetItemFocusA(ByRef ctlListview As MSComctlLib.ListView, ByVal iIndex As Long, Optional ByVal iVisibleIndex = 3) As Boolean
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
        lvItemsPerPage = SendMessage(.hWnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&) + 1
        
        ' Do we even need to scroll? Not if the selected track is already in view
        If (lvCurrentTopIndex >= iIndex) Or (iIndex > lvCurrentTopIndex + lvItemsPerPage) Then
        
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
        End If
    End With

    SetItemFocusA = True
    Exit Function
    
Hell:
MsgBox "ERROR Loading Pallets", vbExclamation, "ERROR"

End Function


Private Sub cmdDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LvScrollDown
End Sub

Private Sub cmdExit_Click()
PaletteName = ""
bSavePalette = False
Unload Me

End Sub
'''
'''Private Sub cmdKey_Click(Index As Integer)
'''Select Case Index
'''   Case 0   'ENTER
'''      'sKeyboardText = txtPalette.Text
'''      'Unload Me
'''   Case 41  'Space
'''      txtPalette.text = txtPalette.text & " "
'''   Case 42  'BackSpace
'''      If Len(txtPalette.text) > 0 Then txtPalette.text = Mid(txtPalette.text, 1, Len(txtPalette.text) - 1)
'''   Case 39, 40 'Caps lock
'''      For I = 11 To 36
'''         If lblStatus.Caption = "CAPS OFF" Then
'''            cmdKey(I).Caption = UCase(cmdKey(I).Caption)
'''         Else
'''            cmdKey(I).Caption = LCase(cmdKey(I).Caption)
'''         End If
'''      Next I
'''      If lblStatus.Caption = "CAPS OFF" Then
'''         lblStatus.Caption = "CAPS ON"
'''      Else
'''         lblStatus.Caption = "CAPS OFF"
'''      End If
'''
'''   Case 1 To 38, 41
'''      txtPalette.text = txtPalette.text & cmdKey(Index).Caption
'''
'''End Select
'''
'''End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.BackColor = &HE7DB49
cmdExit.ForeColor = vbBlack
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.ForeColor = &HE7DB49
cmdExit.BackColor = vbBlack
End Sub

Private Sub cmdOK_Click()

OkAndExit

End Sub

Sub OkAndExit()
Dim iExist As Integer
Dim bExists As Boolean
Dim sText As String


If Trim(txtPalette.text) = "" Then Exit Sub

bOverwritePalet = False

If bSavePalette Then
   bExists = False
   
   'first test to see if this name exists, and if so, ask if it is ok...
   For i = 1 To lvPalettes.ListItems.Count
     If UCase(Trim(lvPalettes.ListItems(i))) = UCase(Trim(txtPalette.text)) Then
        bExists = True
        Exit For
     End If
   Next i
   If bExists Then
      iExist = MsgBox("This Palette entry already exists." & Chr(13) & Chr(13) & "Do you want to Overwrite ???", vbYesNo + vbQuestion, "Entry Exists")
      If iExist = vbNo Then
         txtPalette.SetFocus
         Exit Sub
      Else
        bOverwritePalet = True
      End If
   End If
   
   palletArr(0) = ""
   If Not bExists Then
       palletArr(0) = Trim(txtPalette.text)
   End If
  
   SavePalete Trim(txtPalette.text), iPageno
   bSavePalette = False
  
   If Trim(LCase(txtPalette.Tag)) = "tmp001" Then
      'Remove the Tmp001,dat file, sice we just re-created it
      Dim FSO As New FileSystemObject
      If FSO.FileExists(App.Path & "\Palets\tmp001.dat") Then
         Kill App.Path & "\Palets\tmp001.dat"
      End If
      Set FSO = Nothing
   End If
End If

PaletteName = Trim(txtPalette.text)
SaveSetting regMainKey, regSubKey, "Palette Name", PaletteName
'SaveSetting regMainKey, regSubKey, "Palette Name", PaletteName

Unload Me

End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOk.BackColor = &HE7DB49
cmdOk.ForeColor = vbBlack
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOk.ForeColor = &HE7DB49
cmdOk.BackColor = vbBlack
End Sub

Private Sub cmdUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LvScrollUp
End Sub

Private Sub Form_Activate()

EnableCloseButton Me.hWnd, False

If bSavePalette Then
  Me.Caption = "Save Palette"
  'lvPalettes.Enabled = False
  txtPalette.SetFocus
Else
  Me.Caption = "Load Palette"
End If

lvPalettes.ListItems(1).Selected = True

If lvPalettes.ListItems.Count > 0 Then
   Call SetLVSubImages(lvPalettes, iPreviousSelection, 2, 0, True)
   Call SetLVSubImages(lvPalettes, lvPalettes.SelectedItem.Index, 2, 1, True)
   iPreviousSelection = lvPalettes.SelectedItem.Index
   txtPalette.text = lvPalettes.SelectedItem
End If
txtPalette.Tag = LCase(Trim(txtPalette.text))

sOverwritePalet = Trim(txtPalette.text)

txtPalette.SelStart = 0
txtPalette.SelLength = Len(txtPalette.text)
txtPalette.SetFocus

End Sub

Private Sub Form_Load()

lvPalettes.View = lvwReport
lvPalettes.ColumnHeaders.Add , , "Available Palettes ", 9000
lvPalettes.ColumnHeaders.Add , , "Added Time ", 0
lvPalettes.ColumnHeaders.Add , , "Selected ", 500

lvPalettes.Width = 10000
lvPalettes.BorderStyle = ccNone

'Me.Picture = LoadPicture(App.Path & "\tmpBanner")

lblVersion.Caption = "Version   :    " & App.Major & "." & App.Minor & "." & App.Revision
lblVersion.Top = 750
   lblVersion.Left = 255
   lblVersion.Width = 2940
   lblVersion.FontSize = 7




'''HelpContextID = hlpPlayLists

cmdUp.Visible = False
cmdDown.Visible = False

LoadPallets
'If bSavePalette Then
'  sspKeyboard.Visible = True
'  lvPalettes.Height = 4485  '4635  '4395   '8070
'Else
'  sspKeyboard.Visible = False
'  lvPalettes.Height = 4635 '7200 ' 8040  '4440
'End If

'sspKeyboard.BackColor = vbBlack

lvPalettes.BackColor = vbBlack

'Me.Width = 20385
'Me.Height = 11010 - 60  '11070
'Me.Top = 0
'Me.Left = 0

SSPanel1.Left = -15
SSPanel1.Top = 1440
SSPanel1.Height = Me.Height
SSPanel1.Width = Me.Width + 300

End Sub

Sub LoadPallets()
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

FD = FreeFile
'FileToOpen = App.Path & "\Palettes.Dat"
i = -1

'Loop through files in palets dir, and load filenames into each. Also read first line to determine order for sort...
Dim newFolder As Folder
Dim NewFile As file
Dim newFileName As String

ReDim sArrH(100, 2)

Set newFolder = FSO.GetFolder(App.Path & "\Palets")
For Each NewFile In newFolder.Files
  newFileName = NewFile
  Open NewFile For Input As FD
  Input #FD, sHeading
  Close FD
  sStr = Left(NewFile.name, Len(NewFile.name) - 4)
  iInner = iInner + 1
  sArrH(iInner, 0) = "  " & sStr
  If InStr(1, sHeading, ":") = 3 Then
    sArrH(iInner, 1) = Mid(sHeading, 4)
  Else
    sArrH(iInner, 1) = Mid(sHeading, 5)
  End If
  
Next NewFile
 
Set newFolder = Nothing


'If Not FSO.FileExists(FileToOpen) Then Exit Sub
'
'ReDim sArrH(20, 2)
'
'iInner = -1
'Open FileToOpen For Input As FD
'Do Until (EOF(FD) = True)
'  Input #FD, sHeading
'  If Left(sHeading, 1) = "[" Then
'    If Left(sHeading, 5) <> "[END " Then
'      iInner = iInner + 1
'      sArrH(iInner, 0) = "  " & Replace(Replace(sHeading, "[", ""), "]", "")
'    End If
'  ElseIf Left(sHeading, 3) = "00:" Then
'      sArrH(iInner, 1) = Mid(sHeading, 4)
'  End If
'Loop
'Close FD

'Sort the array first
Sort2Array sArrH(), 1

'Load sorted array into listview
iPos = 0
For iInner = 0 To UBound(sArrH)
   If Trim(sArrH(iInner, 0)) <> "" Then 'Ignore the blanks
     ' If iInner < 10 Then 'Make sure we only list top 9
         iPos = iPos + 1
         Set mItem = lvPalettes.ListItems.Add(, , sArrH(iInner, 0), 0, 3)
         mItem.SubItems(1) = sArrH(iInner, 1)
         Call SetLVSubImages(lvPalettes, iPos, 2, 0, True)
     ' End If
   End If
Next iInner

cmdUp.Visible = iPos > 15
cmdDown.Visible = iPos > 15


'Sort according to alphabet
lvPalettes.SortKey = 1
lvPalettes.SortOrder = 1
' Set Sorted to True to sort the list.
lvPalettes.Sorted = True

'loop throuh the list and add the items to the global array
For iPos = 1 To lvPalettes.ListItems.Count
   palletArr(iPos) = Trim(lvPalettes.ListItems(iPos))
Next iPos

'lvPalettes

End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
   If ApplyStandardTheme Then
'      Me.Width = 20445  '20500  '17925
'      Me.Height = 11010 - 60  '11070
      
      Me.Width = frmPlayer.Width - 120
      Me.Height = frmPlayer.Height - 120
      Me.Top = frmPlayer.Top + 60
      Me.Left = frmPlayer.Left + 60

   Else
   '   Me.Width = 20445  '20500  '20395  '20370   '18030   '17925
   '   Me.Height = 11010 - 60  '11070
      Me.Width = frmPlayer.Width - 180
      Me.Height = frmPlayer.Height - 180
      Me.Top = frmPlayer.Top + 90
      Me.Left = frmPlayer.Left + 90
   End If
End If
End Sub

'''Sub SavePalete(pHeading As String)
'''Dim iInner As Integer
'''Dim FD
'''Dim FileToOpen As String
'''Dim sHeading As String
'''Dim sStr As String
'''Dim sArr() As String
'''Dim sKeepHeading As String
'''Dim bHeadingFound As Boolean
'''Dim bEnd As Boolean
'''Dim bEmptyFile As Boolean
'''Dim sNow As String
'''Dim sTemp As String
'''Dim iP As Integer
'''Dim iValid As Integer
'''Dim bValid As Boolean
'''Dim iMax As Integer
'''
'''
'''sNow = Format(Now, "YYYYMMDDHHmmSS")
'''FD = FreeFile
'''FileToOpen = App.Path & "\Palettes.Dat"
'''iMax = 9
'''iInner = 0
'''bEmptyFile = False
'''ReDim Preserve sArr(1)
'''
''''Create empty file if nothing exists
'''If Not FSO.FileExists(FileToOpen) Then
'''  Open FileToOpen For Output As FD
'''  Close FD
'''End If
'''
'''
''''Loop through the Pallet list and find entries in the file for each
''''Just check is palletArr(0) is not empty because the NEW name will be in entry 0
'''If Trim(palletArr(iP)) <> "" Then  'New entry
'''   iMax = 8  'This is set to 8 so we can only read 8 entires of the palaet list
'''   ReDim Preserve sArr(32)
'''   sArr(0) = "[" & UCase(pHeading) & "]"
'''   sArr(1) = "00:" & sNow
'''   iInner = 1
'''   For i = 1 To 30 'ALWAYS store 30 entries, even if the selection on the player was less
'''      iInner = iInner + 1
'''      If i <= iMaxBut Then 'Maxbutton will be set according to the selecion on the player...
'''         If frmPlayer.sspSongTitle(i).LinkItem <> "" Then
'''         'If frmPlayer.sspSongTitle(i).TagVariant <> "" Then
'''            sArr(iInner) = Format(i, "00") & ":" & AddBlank(frmPlayer.sspSongTitle(i).LinkItem, 2) & "|" & frmPlayer.sspSongTitle(i).Tag & "|" & frmPlayer.lblVol(i).Caption & "|" & frmPlayer.cmdSong(i).TagVariant & "|" & frmPlayer.sspProgress(i).Tag
'''            'sArr(iInner) = Format(i, "00") & ":" & AddBlank(frmPlayer.sspSongTitle(i).TagVariant, 2) & "|" & frmPlayer.sspSongTitle(i).Tag & "|" & frmPlayer.lblVol(i).Caption & "|" & frmPlayer.cmdSong(i).TagVariant & "|" & frmPlayer.sspProgress(i).Tag
'''            'sArr(iInner) = Format(i, "00") & ":" & AddBlank(frmPlayer.sspSongTitle(i).TagVariant, 2) & "|" & frmPlayer.sspSongTitle(i).Tag & "|" & frmPlayer.cpvVolume(i).value & "|" & frmPlayer.cmdSong(i).TagVariant & "|" & frmPlayer.sspProgress(i).Tag
'''          Else
'''            sArr(iInner) = Format(i, "00") & ":"
'''          End If
'''       Else
'''         sArr(iInner) = Format(i, "00") & ":"
'''      End If
'''   Next i
'''   iInner = iInner + 1
'''   sArr(iInner) = "[END " & UCase(pHeading) & "]"
'''End If
'''
''''Read all the entries in the file
'''i = 0
'''For iP = 1 To iMax   '8 or 9. If we store a NEW name, only 8 of the current list is stored.
'''   'Opens the file each time we loop through this palet array
'''   Open FileToOpen For Input As FD
'''
'''   Do Until (EOF(FD) = True)
'''      Input #FD, sTemp
'''      If sTemp = "" Then GoTo ReadNext  'Skip when at end and perhaps a space in last line...
'''      If "[" & UCase(palletArr(iP)) & "]" = UCase(sTemp) Then
'''         bValid = True
'''      End If
'''      'A valid heading was found, thus add everything off this group...
'''      If bValid Then
'''         iInner = iInner + 1
'''         ReDim Preserve sArr(iInner)
'''
'''         If UCase(pHeading) = UCase(palletArr(iP)) Then  'We are going to overwrite the values in the file for this case
'''            If "[" & UCase(palletArr(iP)) & "]" = UCase(sTemp) Then  'We are going to overwrite the values in the file for this case
'''               sArr(iInner) = sTemp
'''            ElseIf Left(sTemp, 2) = "00" Then
'''               sArr(iInner) = "00:" & sNow               'Check for first entry only, and write the CURRENT DateTime here
'''            ElseIf sTemp = "[END " & Replace(Replace(UCase(palletArr(iP)), "[", ""), "]", "") & "]" Then 'Last entry, write to array and exit do after this
'''               sArr(iInner) = sTemp
'''               bValid = False
'''               i = 0
'''               'Close FD 'Close the file, since we are going to open it again with next entry in the list array...
'''               Exit Do
'''            Else
'''               i = i + 1
'''               If i <= iMaxBut Then 'Maxbutton will be set according to the selecion on the player...
'''                  If frmPlayer.sspSongTitle(i).LinkItem <> "" Then
'''                  'If frmPlayer.sspSongTitle(i).TagVariant <> "" Then
'''                     sArr(iInner) = Format(i, "00") & ":" & AddBlank(frmPlayer.sspSongTitle(i).LinkItem, 2) & "|" & frmPlayer.sspSongTitle(i).Tag & "|" & frmPlayer.lblVol(i).Caption & "|" & frmPlayer.cmdSong(i).TagVariant & "|" & frmPlayer.sspProgress(i).Tag
'''                     'sArr(iInner) = Format(i, "00") & ":" & AddBlank(frmPlayer.sspSongTitle(i).TagVariant, 2) & "|" & frmPlayer.sspSongTitle(i).Tag & "|" & frmPlayer.lblVol(i).Caption & "|" & frmPlayer.cmdSong(i).TagVariant & "|" & frmPlayer.sspProgress(i).Tag
'''                     'sArr(iInner) = Format(i, "00") & ":" & AddBlank(frmPlayer.sspSongTitle(i).TagVariant, 2) & "|" & frmPlayer.sspSongTitle(i).Tag & "|" & frmPlayer.cpvVolume(i).value & "|" & frmPlayer.cmdSong(i).TagVariant & "|" & frmPlayer.sspProgress(i).Tag
'''                   Else
'''                     sArr(iInner) = Format(i, "00") & ":"
'''                   End If
'''                Else
'''                  sArr(iInner) = Format(i, "00") & ":"
'''               End If
'''            End If
'''         Else 'Just write whatever is in the file...
'''            sArr(iInner) = sTemp
'''            If sTemp = "[END " & Replace(Replace(UCase(palletArr(iP)), "[", ""), "]", "") & "]" Then 'Last entry, write to array and exit do after this
'''               bValid = False
'''               i = 0
'''               'Close FD 'Close the file, since we are going to open it again with next entry in the list array...
'''               Exit Do
'''            End If
'''         End If
'''      End If
'''ReadNext:
'''   Loop
'''   Close FD 'Make sure the file is closed here
'''Next iP
'''Close FD
'''
'''
'''
'''
''''Write new file from array
'''Open FileToOpen For Output As FD
'''For i = 0 To UBound(sArr)
'''  If Trim(sArr(i)) <> "" Then Print #FD, sArr(i)
'''Next i
'''Close FD
'''
'''
'''End Sub


Private Sub lvPalettes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' When a ColumnHeader object is clicked, the ListView control is                    '
' sorted by the subitems of that column.                                            '
' Set the SortKey to the Index of the ColumnHeader - 1                              '
'====================================================================================
lvPalettes.SortKey = ColumnHeader.Index - 1
lvPalettes.SortOrder = lvPalettes.SortOrder Xor 1
' Set Sorted to True to sort the list.
lvPalettes.Sorted = True

End Sub

Private Sub lvPalettes_DblClick()
If lvPalettes.ListItems.Count > 0 Then
  txtPalette.text = lvPalettes.SelectedItem
  Call OkAndExit
End If

End Sub

Private Sub lvPalettes_ItemClick(ByVal Item As MSComctlLib.ListItem)

   If lvPalettes.ListItems.Count > 0 Then
      Call SetLVSubImages(lvPalettes, iPreviousSelection, 2, 0, True)
      Call SetLVSubImages(lvPalettes, lvPalettes.SelectedItem.Index, 2, 1, True)
      iPreviousSelection = lvPalettes.SelectedItem.Index
      
      txtPalette.text = lvPalettes.SelectedItem
     
   End If

End Sub

Private Sub txtPalette_GotFocus()
'If bSavePalette Then
'   txtPalette.SelStart = 0
'   txtPalette.SelLength = Len(txtPalette.text)
'End If
End Sub

Private Sub txtPalette_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call OkAndExit
  End If
End Sub
