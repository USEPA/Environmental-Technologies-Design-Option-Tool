VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmDevelopNotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FaVOr Development Notes"
   ClientHeight    =   6345
   ClientLeft      =   975
   ClientTop       =   1365
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnlAbout 
      Height          =   5595
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7335
      _Version        =   65536
      _ExtentX        =   12938
      _ExtentY        =   9869
      _StockProps     =   15
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      Begin VB.Shape boxAbout 
         Height          =   1575
         Left            =   840
         Top             =   3585
         Width           =   4575
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Houghton, Michigan 49931-1295"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   1200
         TabIndex        =   8
         Top             =   4860
         Width           =   3195
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "1400 Townsend Drive"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   1200
         TabIndex        =   7
         Top             =   4560
         Width           =   3075
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Michigan Technological University"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   1200
         TabIndex        =   6
         Top             =   4260
         Width           =   3015
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Dept. of Geology and Geological Engineering"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   1200
         TabIndex        =   5
         Top             =   3960
         Width           =   4155
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Dr. Alex S. Mayer"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Top             =   3660
         Width           =   2955
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "All correspondence or questions regarding the model should be forwarded to:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   3180
         Width           =   6855
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"DevelopNotes.frx":0000
         ForeColor       =   &H80000008&
         Height          =   1275
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   1740
         Width           =   6795
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"DevelopNotes.frx":01BD
         ForeColor       =   &H80000008&
         Height          =   1275
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   300
         Width           =   6735
      End
   End
   Begin Threed.SSPanel pnlAuthors 
      Height          =   5595
      Left            =   7590
      TabIndex        =   9
      Top             =   240
      Width           =   7335
      _Version        =   65536
      _ExtentX        =   12938
      _ExtentY        =   9869
      _StockProps     =   15
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alex S. Mayer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   540
         TabIndex        =   18
         Top             =   1290
         Width           =   3615
      End
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Assistant Professor"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   885
         TabIndex        =   17
         Top             =   1680
         Width           =   3600
      End
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Geology and Geological Eng., MTU"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   885
         TabIndex        =   16
         Top             =   1980
         Width           =   3600
      End
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rich Voigt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   540
         TabIndex        =   15
         Top             =   420
         Width           =   3735
      End
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Vaughn Wildfong"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   600
         TabIndex        =   14
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "M.S. Civil and Environmental Eng., MTU"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   885
         TabIndex        =   13
         Top             =   2820
         Width           =   3600
      End
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Dennis Blanchard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   600
         TabIndex        =   12
         Top             =   3300
         Width           =   2715
      End
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "B.S. Geological Eng., MTU"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   885
         TabIndex        =   11
         Top             =   3720
         Width           =   3600
      End
      Begin VB.Label lblAuthors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "M.S. Civil and Environmental Eng., MTU"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   885
         TabIndex        =   10
         Top             =   840
         Width           =   3600
      End
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdPage 
      Height          =   495
      Left            =   3660
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5760
      Width           =   2325
      _Version        =   65536
      _ExtentX        =   4101
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "{ctl.caption}"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDevelopNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CMD_CAPTION_1 = "About FaVOr"
Const CMD_CAPTION_2 = "The Authors"
Dim PAGEMODE As Integer
Const PAGEMODE_1 = 1
Const PAGEMODE_2 = 2




Const frmDevelopNotes_declarations_end = True


Sub Refresh_This()
  Select Case PAGEMODE
    Case PAGEMODE_1:
      pnlAbout.ZOrder
      DoEvents
      cmdPage.Caption = CMD_CAPTION_1
    Case PAGEMODE_2:
      pnlAuthors.ZOrder
      DoEvents
      cmdPage.Caption = CMD_CAPTION_2
  End Select
End Sub


Private Sub cmdOK_Click()
  Unload Me
  Exit Sub
End Sub
Private Sub cmdPage_Click()
  Select Case PAGEMODE
    Case PAGEMODE_1: PAGEMODE = PAGEMODE_2
    Case PAGEMODE_2: PAGEMODE = PAGEMODE_1
  End Select
  Call Refresh_This
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  pnlAuthors.Move pnlAbout.Left, pnlAbout.Top
  PAGEMODE = PAGEMODE_1
  Call CenterOnForm(Me, frmMain)
  Call Refresh_This
End Sub



