VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmKeyboard 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox rtfData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   660
      TabIndex        =   44
      Top             =   300
      Width           =   7110
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   38
      Left            =   6735
      TabIndex        =   43
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "!"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   0
      Left            =   7095
      TabIndex        =   42
      Top             =   2325
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "GO"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   1
      Left            =   495
      TabIndex        =   41
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   37
      Left            =   6075
      TabIndex        =   40
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CAPS OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6720
      TabIndex        =   39
      Top             =   4440
      Width           =   1425
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   41
      Left            =   1455
      TabIndex        =   38
      Top             =   3690
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "S P A C E"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   36
      Left            =   5430
      TabIndex        =   37
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "m"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   35
      Left            =   4770
      TabIndex        =   36
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "n"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   34
      Left            =   4110
      TabIndex        =   35
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "b"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   33
      Left            =   3450
      TabIndex        =   34
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "v"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   32
      Left            =   2790
      TabIndex        =   33
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "c"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   31
      Left            =   2130
      TabIndex        =   32
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "x"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   30
      Left            =   1470
      TabIndex        =   31
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "z"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   39
      Left            =   705
      TabIndex        =   30
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   4210752
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmKeyboard.frx":0000
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   29
      Left            =   6435
      TabIndex        =   29
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "l"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   28
      Left            =   5775
      TabIndex        =   28
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "k"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   27
      Left            =   5115
      TabIndex        =   27
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "j"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   26
      Left            =   4455
      TabIndex        =   26
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "h"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   25
      Left            =   3795
      TabIndex        =   25
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "g"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   24
      Left            =   3135
      TabIndex        =   24
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "f"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   23
      Left            =   2475
      TabIndex        =   23
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "d"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   22
      Left            =   1815
      TabIndex        =   22
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "s"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   21
      Left            =   1155
      TabIndex        =   21
      Top             =   2325
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "a"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   40
      Left            =   7395
      TabIndex        =   20
      Top             =   3000
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   4210752
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmKeyboard.frx":0452
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   20
      Left            =   6765
      TabIndex        =   19
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "p"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   19
      Left            =   6105
      TabIndex        =   18
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "o"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   18
      Left            =   5445
      TabIndex        =   17
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "i"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   17
      Left            =   4785
      TabIndex        =   16
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "u"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   16
      Left            =   4125
      TabIndex        =   15
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "y"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   15
      Left            =   3465
      TabIndex        =   14
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "t"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   14
      Left            =   2805
      TabIndex        =   13
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "r"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   13
      Left            =   2145
      TabIndex        =   12
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "e"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   12
      Left            =   1485
      TabIndex        =   11
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "w"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   11
      Left            =   825
      TabIndex        =   10
      Top             =   1650
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "q"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   42
      Left            =   7095
      TabIndex        =   9
      Top             =   960
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1138
      _Version        =   131074
      BackColor       =   8421504
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmKeyboard.frx":08A4
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   10
      Left            =   6435
      TabIndex        =   8
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   9
      Left            =   5775
      TabIndex        =   7
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "9"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   8
      Left            =   5115
      TabIndex        =   6
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "8"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   7
      Left            =   4455
      TabIndex        =   5
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "7"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   6
      Left            =   3795
      TabIndex        =   4
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "6"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   5
      Left            =   3135
      TabIndex        =   3
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "5"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   4
      Left            =   2475
      TabIndex        =   2
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "4"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   3
      Left            =   1815
      TabIndex        =   1
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "3"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdKey 
      Height          =   645
      Index           =   2
      Left            =   1155
      TabIndex        =   0
      Top             =   975
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1138
      _Version        =   131074
      ForeColor       =   16777215
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2"
      ButtonStyle     =   1
      BevelWidth      =   1
      Outline         =   0   'False
   End
End
Attribute VB_Name = "frmKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub cmdKey_Click(Index As Integer)
'
''Select Case Index
''   Case 0   'ENTER
''     ' sKeyboardText = rtfData.Text
''      Unload Me
''   Case 41  'Space
''      rtfData.Text = rtfData.Text & " "
''   Case 15  'BackSpace
''      If Len(rtfData.Text) > 0 Then rtfData.Text = Mid(rtfData.Text, 1, Len(rtfData.Text) - 1)
''   Case 39, 40 'Caps lock
''      For i = 11 To 36
''         If lblStatus.Caption = "CAPS OFF" Then
''            cmdKey(i).Caption = UCase(cmdKey(i).Caption)
''         Else
''            cmdKey(i).Caption = LCase(cmdKey(i).Caption)
''         End If
''      Next i
''      If lblStatus.Caption = "CAPS OFF" Then
''         lblStatus.Caption = "CAPS ON"
''      Else
''         lblStatus.Caption = "CAPS OFF"
''      End If
''
''   Case 1 To 38, 41
''      rtfData.Text = rtfData.Text & cmdKey(Index).Caption
''
''End Select
'
'End Sub
'
'Private Sub Form_Load()
'On Error Resume Next
'
'lblStatus.Caption = "CAPS OFF"
'
'End Sub
