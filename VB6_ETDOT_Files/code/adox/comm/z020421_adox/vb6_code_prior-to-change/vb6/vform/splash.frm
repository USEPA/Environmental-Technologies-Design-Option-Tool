VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5745
   ClientLeft      =   2520
   ClientTop       =   2490
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5745
   ScaleWidth      =   9180
   Begin Threed.SSPanel sspanel_disclaimer 
      Height          =   3795
      Left            =   9210
      TabIndex        =   18
      Top             =   4890
      Width           =   8325
      _Version        =   65536
      _ExtentX        =   14684
      _ExtentY        =   6694
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel3 
         Height          =   3075
         Left            =   120
         TabIndex        =   20
         Top             =   570
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   5424
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lblDisclaimer 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "lblDisclaimer"
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
            Height          =   2895
            Left            =   90
            TabIndex        =   21
            Top             =   90
            Width           =   7845
         End
      End
      Begin VB.Label lblDisclaimerTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Disclaimer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   60
         TabIndex        =   19
         Top             =   60
         Width           =   8175
      End
   End
   Begin Threed.SSPanel sspanel_logos 
      Height          =   4035
      Left            =   420
      TabIndex        =   4
      Top             =   990
      Width           =   8325
      _Version        =   65536
      _ExtentX        =   14684
      _ExtentY        =   7117
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picLogos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Index           =   2
         Left            =   390
         Picture         =   "splash.frx":0000
         ScaleHeight     =   705
         ScaleWidth      =   3285
         TabIndex        =   23
         Top             =   3240
         Width           =   3285
      End
      Begin Threed.SSPanel Panel3D3 
         Height          =   825
         Left            =   120
         TabIndex        =   8
         Top             =   750
         Width           =   8055
         _Version        =   65536
         _ExtentX        =   14208
         _ExtentY        =   1455
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lbldesc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Development:"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   450
            Width           =   2055
         End
         Begin VB.Label lbldesc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Model and Software:"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label lblAuthors 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Shumin Hu      Ke Li      David W. Hand      John C. Crittenden"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   2340
            TabIndex        =   10
            Top             =   180
            Width           =   5595
         End
         Begin VB.Label lblAuthors 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Eric J. Oman   David R. Hokanson   Karen A . Mansfeldt"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   2340
            TabIndex        =   9
            Top             =   450
            Width           =   5595
         End
      End
      Begin VB.PictureBox picLogos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   1
         Left            =   1050
         Picture         =   "splash.frx":722A
         ScaleHeight     =   1695
         ScaleWidth      =   2085
         TabIndex        =   5
         Top             =   1560
         Width           =   2085
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1425
         Left            =   4320
         TabIndex        =   12
         Top             =   1950
         Width           =   3765
         _Version        =   65536
         _ExtentX        =   6641
         _ExtentY        =   2514
         _StockProps     =   15
         Caption         =   "SSPanel1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lblVersionInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Version {ver} {STANDARD}"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   120
            Width           =   3645
         End
         Begin VB.Label lblVersionInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Expires on MM/DD/YYYY"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   16
            Top             =   360
            Width           =   3645
         End
         Begin VB.Label lblVersionInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Copyright {years}"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   15
            Top             =   600
            Width           =   3645
         End
         Begin VB.Label lblVersionInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Michigan Technological University"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   60
            TabIndex        =   14
            Top             =   840
            Width           =   3645
         End
         Begin VB.Label lblVersionInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Houghton, Michigan"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   4
            Left            =   60
            TabIndex        =   13
            Top             =   1080
            Width           =   3645
         End
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "CenCITT"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   7
         Top             =   180
         Width           =   5685
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Center for Clean Industrial and Treatment Technologies"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Top             =   420
         Width           =   6495
      End
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "E&xit"
      Height          =   525
      Left            =   7440
      TabIndex        =   2
      Top             =   5070
      Width           =   1305
   End
   Begin VB.CommandButton cmdButton2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "I agree, never show again"
      Height          =   525
      Left            =   1980
      TabIndex        =   1
      Top             =   5070
      Width           =   2535
   End
   Begin VB.CommandButton cmdButton1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Continue"
      Height          =   525
      Left            =   420
      TabIndex        =   0
      Top             =   5070
      Width           =   1455
   End
   Begin Threed.SSPanel sspanel_maintitle 
      Height          =   795
      Left            =   420
      TabIndex        =   3
      Top             =   120
      Width           =   8325
      _Version        =   65536
      _ExtentX        =   14684
      _ExtentY        =   1402
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   705
         Left            =   150
         TabIndex        =   24
         Top             =   60
         Width           =   8085
         Begin VB.Image Image1 
            Height          =   11520
            Left            =   360
            Picture         =   "splash.frx":79A4
            Top             =   -60
            Width           =   15360
         End
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdButton1_Click()
  splash_button_pressed = 1
  Unload Me
End Sub

Private Sub cmdButton2_Click()
  splash_button_pressed = 2
  Unload Me
End Sub

Private Sub cmdExit_Click()
  splash_button_pressed = 3
  Unload Me
End Sub

Private Sub Form_Load()
Dim s As String
  Me.height = 6165
  Me.width = 9300
  Me.left = (Screen.width - Me.width) / 2
  Me.top = (Screen.height - Me.height) / 2
  If (splash_mode = 0) Then
    'SHOW THE CONTINUE/EXIT FRONT WINDOW.
    cmdButton1.visible = True
    cmdButton1.Caption = "&Continue"
    cmdButton2.visible = False
    cmdExit.visible = True
    'VERSION INFO.
    lblVersionInfo(0).Caption = "Version " & get_program_version_with_build_info()
    'EXPIRATION INFO.
    lblVersionInfo(1).Caption = get_expiration_info()
    'COPYRIGHT INFO.
    lblVersionInfo(2).Caption = "Copyright " & AppCopyrightYears
    'ETC.
    sspanel_disclaimer.visible = False
    'cmdButton1.SetFocus
    cmdButton1.TabIndex = 0
  End If
  If (splash_mode = 1) Then
    'SHOW THE DISCLAIMER WINDOW.
    cmdButton1.visible = True
    cmdButton1.Caption = "I Agree"
    cmdButton2.visible = True
    cmdExit.visible = True
    sspanel_logos.visible = False
    sspanel_disclaimer.left = sspanel_logos.left
    sspanel_disclaimer.top = sspanel_logos.top
    sspanel_disclaimer.visible = True
    
    If UCase(get_program_releasetype()) = "BETA" Then
      s = "This is a DEMONSTRATION version of the AdOx/Dyestudy program. "
      s = s & "This demonstration version cannot be changed or saved. "
      s = s & "For the full version of this program, please contact "
      s = s & "Dr. David W. Hand (dwhand@mtu.edu or 906-487-2777). "
      s = s & "Additional information about this program is available at "
      s = s & "our web site (http://www.cpas.mtu.edu/etdot/)."
    Else
      s = "By choosing " & Chr$(34) & "I Agree" & Chr$(34) & " you acknowledge that "
      s = s & "this software is under development and not guaranteed to be free "
      s = s & "of errors.  Furthermore there may be errors in the software that "
      s = s & "lead to erroneous output.  MTU shall not be liable for any loss, "
      s = s & "damage, injury, or casualty of any kind, or by whomever, caused "
      s = s & "to the person or property arising out of or resulting from "
      s = s & "receipt and use of any aspect of the software.  "
      s = s & "References to specific commercial products, processes, or services "
      s = s & "by trademark, manufacturer, or otherwise does not necessarily "
      s = s & "constitute or imply endorsement/recommendation by the authors or "
      s = s & "the respective organizations under which the software "
      s = s & "was developed."
    End If
    lblDisclaimer.Caption = s
   
    'cmdButton1.SetFocus
    cmdButton1.TabIndex = 0
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If (splash_button_pressed = 0) Then
    'If they got here, they must have selected "Close",
    'so perform the exit functionality.
    splash_button_pressed = 3
  End If
End Sub

