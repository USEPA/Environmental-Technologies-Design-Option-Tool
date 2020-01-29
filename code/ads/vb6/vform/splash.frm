VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Information"
   ClientHeight    =   5745
   ClientLeft      =   1245
   ClientTop       =   1575
   ClientWidth     =   9150
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5745
   ScaleWidth      =   9150
   Begin Threed.SSFrame SSFrame1 
      Height          =   2055
      Left            =   9750
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   3705
      _Version        =   65536
      _ExtentX        =   6535
      _ExtentY        =   3625
      _StockProps     =   14
      Caption         =   "Invisible -- Unused"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
         Height          =   1755
         Index           =   0
         Left            =   360
         Picture         =   "splash.frx":0000
         ScaleHeight     =   1755
         ScaleWidth      =   1815
         TabIndex        =   11
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label lbldesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   1320
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lbldesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1320
         TabIndex        =   12
         Top             =   990
         Width           =   2055
      End
   End
   Begin Threed.SSPanel sspanel_disclaimer 
      Height          =   3795
      Left            =   9210
      TabIndex        =   6
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
         TabIndex        =   8
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
            TabIndex        =   9
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
         TabIndex        =   7
         Top             =   60
         Width           =   8175
      End
   End
   Begin Threed.SSPanel sspanel_logos 
      Height          =   3915
      Left            =   90
      TabIndex        =   4
      Top             =   1080
      Width           =   8955
      _Version        =   65536
      _ExtentX        =   15796
      _ExtentY        =   6906
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
      Begin Threed.SSPanel sspNames 
         Height          =   2115
         Left            =   4830
         TabIndex        =   5
         Top             =   1740
         Width           =   4065
         _Version        =   65536
         _ExtentX        =   7170
         _ExtentY        =   3731
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
      End
      Begin Threed.SSPanel SSPanelLogos 
         Height          =   3795
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   4425
         _Version        =   65536
         _ExtentX        =   7805
         _ExtentY        =   6694
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
            Height          =   735
            Index           =   2
            Left            =   150
            Picture         =   "splash.frx":3776
            ScaleHeight     =   735
            ScaleWidth      =   3315
            TabIndex        =   17
            Top             =   2760
            Width           =   3315
         End
         Begin VB.PictureBox picLogos 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1755
            Index           =   1
            Left            =   810
            Picture         =   "splash.frx":A9A0
            ScaleHeight     =   1755
            ScaleWidth      =   1815
            TabIndex        =   16
            Top             =   90
            Width           =   1815
         End
         Begin VB.Label lblCompany 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "{Set in Form_Load}"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   870
            TabIndex        =   15
            Top             =   1980
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1455
         Left            =   4830
         TabIndex        =   19
         Top             =   60
         Width           =   4065
         _Version        =   65536
         _ExtentX        =   7170
         _ExtentY        =   2566
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
            Left            =   30
            TabIndex        =   22
            Top             =   210
            Width           =   4005
         End
         Begin VB.Label lblVersionInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Expires on MM/DD/YYYY"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   30
            TabIndex        =   21
            Top             =   570
            Width           =   4005
         End
         Begin VB.Label lblVersionInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Copyright {years}"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   30
            TabIndex        =   20
            Top             =   930
            Width           =   4005
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "E&xit"
      Height          =   525
      Left            =   7740
      TabIndex        =   2
      Top             =   5070
      Width           =   1305
   End
   Begin VB.CommandButton cmdButton2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "I agree, never show again"
      Height          =   525
      Left            =   9390
      TabIndex        =   1
      Top             =   3780
      Width           =   2535
   End
   Begin VB.CommandButton cmdButton1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Continue"
      Height          =   525
      Left            =   90
      TabIndex        =   0
      Top             =   5070
      Width           =   1455
   End
   Begin Threed.SSPanel sspanel_maintitle 
      Height          =   945
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   8955
      _Version        =   65536
      _ExtentX        =   15796
      _ExtentY        =   1667
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   480
         Picture         =   "splash.frx":B11A
         ScaleHeight     =   735
         ScaleWidth      =   7965
         TabIndex        =   23
         Top             =   90
         Width           =   7965
      End
   End
   Begin VB.Label lblAdditionalNotice 
      Alignment       =   2  'Center
      Caption         =   "{This program is protected by .... see code in Form_Load}"
      Height          =   465
      Left            =   1950
      TabIndex        =   18
      Top             =   5070
      Width           =   5385
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Const frmSplash_decl_end = True


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
Dim ctr_location As Integer
Dim s As String
'Call debug_output("L1")
  Me.Height = 6165
  Me.Width = 9300
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
'Call debug_output("L2")
  '
  ' CENTER THE LOGOS.
  '
  ctr_location = SSPanelLogos.Width / 2 - lblCompany(0).Width / 2
  lblCompany(0).Left = ctr_location
'  ctr_location = SSPanelLogos.Top + SSPanelLogos.Height / 4 - SSPanel1.Height / 2
'  SSPanel1.Top = ctr_location
  ctr_location = SSPanelLogos.Left + 3 * SSPanelLogos.Width / 4 - SSPanel1.Width / 2
  picLogos(1).Visible = True
  ctr_location = SSPanelLogos.Width / 2 - picLogos(1).Width / 2
  picLogos(1).Left = ctr_location
  ctr_location = SSPanelLogos.Width / 2 - picLogos(2).Width / 2
  picLogos(2).Left = ctr_location
  picLogos(1).Top = 50
  picLogos(2).Top = SSPanelLogos.Height - picLogos(2).Height - 50
  ctr_location = picLogos(1).Top + picLogos(1).Height + 50
  lblCompany(0).Top = ctr_location
  '
  ' MISCELLANEOUS SETTINGS.
  '
  sspNames.Caption = _
      "David R. Hokanson" & Chr$(13) & Chr$(13) & _
      "David W. Hand" & Chr$(13) & Chr$(13) & _
      "John C. Crittenden" & Chr$(13) & Chr$(13) & _
      "Tony N. Rogers" & Chr$(13) & Chr$(13) & _
      "Eric J. Oman"
  lblAdditionalNotice.Caption = _
      "This program is protected by U.S. and international" & _
      vbCrLf & _
      "copyright laws as described in Help About."
  lblCompany(0).Caption = _
      "National Center for" & Chr$(13) & _
      "Clean Industrial and Treatment Technologies" & Chr$(13) & _
      Chr$(13) & _
      "Michigan Technological University" & Chr$(13) & _
      "Houghton, Michigan"
  picTitle.Visible = True
  '
  ' LICENSE-RELATED SETTINGS.
  ''''lblVersionInfo(0).Caption = "Version " & get_program_version_with_build_info()
  ''''lblVersionInfo(0).Caption = _
      "Version " & get_program_version_with_build_info_VB4(True) & _
      " (" & get_program_releasetype() & ")"
  lblVersionInfo(0).Caption = _
      get_program_version_with_build_info_VB4(True)
  'lblVersionInfo(0).Caption = _
  '    "Version 1.0"
  '    MsgBox "Fix this !!!!  (contact ejoman@mtu.edu)"
  lblVersionInfo(1).Caption = get_expiration_info(True)
  lblVersionInfo(2).Caption = "Copyright " & AppCopyrightYears
  '
  ' PROGRAM-SPECIFIC SETTINGS.
  '
  If (AppProgramKey = "ADS") Then
    If (Activate_PSDMInRoom = True) Then
      ''''sspanel_maintitle.Caption = "Indoor Air Filtration Model"
      ''''lblCompany(0).Caption = "MTU"
      lblCompany(0).Caption = _
          "Michigan Technological University" & Chr$(13) & _
          "Houghton, Michigan"
      picLogos(1).Visible = False
      sspanel_maintitle.Caption = _
          AppName_For_Display_Long & " (" & AppName_For_Display_Short & ")"
      picTitle.Visible = False
    Else
      ' DO NOTHING.
    End If
  End If
  If (AppProgramKey = "ASAP") Then
    ' DO NOTHING.
  End If
  If (AppProgramKey = "STEPP") Then
    ' DO NOTHING.
  End If
  If (splash_mode = 0) Then
    '
    ' SHOW THE CONTINUE/EXIT FRONT WINDOW.
    '
'Call debug_output("L3")
    cmdButton1.Visible = True
    cmdButton1.Caption = "&Continue"
    cmdButton2.Visible = False
    cmdExit.Visible = True
    '
    ' ETC.
    '
    sspanel_disclaimer.Visible = False
    'cmdButton1.SetFocus
    cmdButton1.TabIndex = 0
'Call debug_output("L4")
  End If
  ''''If (splash_mode = 1) Then
  If (splash_mode = 1) Or (splash_mode = 101) Then
'Call debug_output("L5")
    'SHOW THE DISCLAIMER WINDOW.
    If (splash_mode = 101) Then
      cmdButton1.Visible = False
      cmdButton2.Visible = False
      cmdExit.Visible = True
      cmdExit.Caption = "&Close"
    Else
      cmdButton1.Visible = True
      cmdButton1.Caption = "I Agree"
      cmdButton2.Visible = True
      cmdExit.Visible = True
    End If
    'cmdButton1.Visible = True
    'cmdButton1.Caption = "I Agree"
    'cmdButton2.Visible = True
    'cmdExit.Visible = True
    sspanel_logos.Visible = False
    sspanel_disclaimer.Left = sspanel_logos.Left
    sspanel_disclaimer.Top = sspanel_logos.Top
    sspanel_disclaimer.Visible = True
    s = "By choosing " & Chr$(34) & "I Agree" & Chr$(34) & " you acknowledge that "
    s = s & "this software is under development and not guaranteed to be free "
    s = s & "of errors.  Furthermore there may be errors in the software that "
    s = s & "lead to erroneous output.  MTU shall not be liable for any loss, "
    s = s & "damage, injury, or casualty of whatsoever kind, or by whomsoever "
    s = s & "caused to the person or property of anyone arising out of or "
    s = s & "resulting from receipt and use of any aspect of the software.  "
    s = s & "References to specific commercial products, processes, or services "
    s = s & "by trademark, manufacturer, or otherwise does not necessarily "
    s = s & "constitute or imply endorsement/recommendation by the authors or "
    s = s & "the respective organizations under which the software "
    s = s & "was developed."
    lblDisclaimer.Caption = s
    'cmdButton1.SetFocus
    cmdButton1.TabIndex = 0
'Call debug_output("L6")
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If (splash_button_pressed = 0) Then
    'If they got here, they must have selected "Close",
    'so perform the exit functionality.
    splash_button_pressed = 3
  End If
End Sub

Private Sub lblAuthors_Click(Index As Integer)

End Sub
