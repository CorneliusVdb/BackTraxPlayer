VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmSerial 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serial Number"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel6 
      Height          =   315
      Left            =   2505
      TabIndex        =   2
      Top             =   2970
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   131074
      BackColor       =   16761024
      BorderWidth     =   1
      BevelOuter      =   0
      Begin VB.TextBox txtSerial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   60
         MaxLength       =   11
         TabIndex        =   3
         Top             =   60
         Width           =   1410
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ESC to exit"
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   4635
      TabIndex        =   6
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "For the low, low price of only R99.99, you can get the FULL version of BacktraxPlayer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   735
      Left            =   420
      TabIndex        =   5
      Top             =   225
      Width           =   5460
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSerial.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E7DB49&
      Height          =   1425
      Left            =   420
      TabIndex        =   4
      Top             =   1170
      Width           =   5460
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   810
      Left            =   4635
      TabIndex        =   1
      Top             =   2715
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1429
      _Version        =   131074
      ForeColor       =   15194953
      BackColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Candara"
         Size            =   14.25
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your serial number here :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   420
      TabIndex        =   0
      Top             =   2910
      Width           =   1875
   End
End
Attribute VB_Name = "frmSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bValidationDone As Boolean

Private Sub cmdExit_Click()

If Not ValidateSerial(txtSerial.text) Then
   MsgBox "Invalid Serial number !!" & Chr(13) & Chr(13) & "Please try again.", vbCritical, "Validation Failed"
   txtSerial.text = ""
   txtSerial.SetFocus
   Exit Sub
Else
   SaveSetting regMainKey, regSubKey, "SerialNumber", Trim(txtSerial.text)
   SaveSetting regMainKey, regSubKey, "DemoUsed", -9
   bValidationDone = True
   DemoFlag = False
   Unload Me
End If

End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.BackColor = &HE7DB49
cmdExit.ForeColor = vbBlack
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.BackColor = &HE7DB49
cmdExit.ForeColor = vbBlack
End Sub

Private Sub Form_Activate()
txtSerial.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'ESCAPE
   Unload Me
End If
End Sub

Private Sub Form_Load()

Screen.MousePointer = vbDefault
DoEvents
bValidationDone = False
bSkipValidation = False
txtSerial.text = GetSetting(regMainKey, regSubKey, "SerialNumber")
If txtSerial.text = "DEMO" Then txtSerial.text = ""
Label3.Caption = "For the low, low price of only R250.00, you can get the FULL version of BacktraxPlayer."
Label2.Caption = "To obtain a Serial Number, please contact Lilac Productions" & Chr(13) & Chr(13) & "Tel. : +27 73 343 9960" & Chr(13) & "Mail : registrations@lilacpro.co.za" & Chr(13) & "Webpage : www.lilacpro.co.za"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Not bValidationDone Then bSkipValidation = True

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And Shift = 7 Then
   Dim sSerial As String
   sSerial = GenerateNewSerial
   'MsgBox "Serial number : " & sSerial, vbInformation, "Serial Number"
   txtSerial.text = sSerial
   'Debug.Print "Serial number : " & sSerial
   Exit Sub
End If

End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
