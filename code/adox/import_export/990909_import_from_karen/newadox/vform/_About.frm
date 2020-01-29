VERSION 4.00
Begin VB.Form frmAboutOld 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4455
   ClientLeft      =   6060
   ClientTop       =   5130
   ClientWidth     =   5640
   Height          =   4860
   Left            =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Top             =   4785
   Width           =   5760
   Begin VB.CommandButton Command2 
      Caption         =   "System Info..."
      Height          =   315
      Left            =   4140
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   4140
      TabIndex        =   6
      Top             =   3180
      Width           =   1335
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1035
      Left            =   1200
      TabIndex        =   9
      Top             =   1680
      Width           =   4275
      _Version        =   65536
      _ExtentX        =   7541
      _ExtentY        =   1826
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Label lbl_regserialnum 
         Caption         =   "lbl_regserialnum"
         Height          =   255
         Left            =   1500
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   " Serial Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lbl_regcompany 
         Caption         =   "lbl_regcompany"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   3555
      End
      Begin VB.Label lbl_regname 
         Caption         =   "lbl_regname"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Label lbl_version 
      Caption         =   "lbl_version"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   360
      Width           =   4200
   End
   Begin VB.Label lbl_copyright 
      Caption         =   "lbl_copyright"
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   900
      Width           =   4300
   End
   Begin VB.Label lbl_progname 
      Caption         =   "lbl_progname"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   120
      Width           =   4200
   End
   Begin VB.Image ProgramIcon 
      Height          =   795
      Left            =   180
      Top             =   180
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   " This program is licensed to:"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1440
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Warning: This computer program is protected by"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   3060
      Width           =   3675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " maximum extent possible under law."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   5
      Top             =   3960
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " and criminal penalties, and will be prosecuted to the"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   3780
      Width           =   3795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " program, or any portion of it, may result in severe civil"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   3600
      Width           =   3930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Unauthorized reproduction or distribution of this"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   3420
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " copyright law and international treaties."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   3240
      Width           =   2955
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   60
      X2              =   5460
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmAboutOld"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Unload Me

End Sub


Private Sub Form_Load()
Dim ver As String

ver = " Version " + get_program_version_with_build_info()
'AppVersion
'ver = ver & "   (build "
'ver = ver & Trim$(App.Major) & "."
'ver = ver & Trim$(App.Minor) & "."
'ver = ver & Trim$(App.Revision) & ")"

Me.Caption = "About " + App.Title
lbl_progname.Caption = " " + App.Title
lbl_version.Caption = ver
lbl_copyright.Caption = " Copyright " + Chr(169) + " " + AppCopyright

lbl_regname.Caption = " " + AppRegisteredUser
lbl_regcompany.Caption = " " + AppRegisteredCompany
lbl_regserialnum.Caption = " " + AppRegisteredSerial

ProgramIcon.Picture = frmMain.Icon

'frm.Left = (Screen.Width - frm.Width) / 2
'frm.Top = (Screen.Height - frm.Height) / 2
Call CenterOnForm(Me, frmMain)

End Sub


