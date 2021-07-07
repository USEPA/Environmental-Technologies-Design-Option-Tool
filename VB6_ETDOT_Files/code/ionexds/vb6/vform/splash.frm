VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5745
   ClientLeft      =   2820
   ClientTop       =   2685
   ClientWidth     =   9180
   ControlBox      =   0   'False
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
   Icon            =   "splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5745
   ScaleWidth      =   9180
   Begin Threed.SSPanel sspanel_disclaimer 
      Height          =   3795
      Left            =   9210
      TabIndex        =   5
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
         TabIndex        =   7
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
            TabIndex        =   8
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
         TabIndex        =   6
         Top             =   60
         Width           =   8175
      End
   End
   Begin Threed.SSPanel sspanel_logos 
      Height          =   4005
      Left            =   390
      TabIndex        =   4
      Top             =   960
      Width           =   4305
      _Version        =   65536
      _ExtentX        =   7594
      _ExtentY        =   7064
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
         Height          =   1755
         Index           =   1
         Left            =   1140
         Picture         =   "splash.frx":030A
         ScaleHeight     =   1755
         ScaleWidth      =   1815
         TabIndex        =   16
         Top             =   120
         Width           =   1815
      End
      Begin VB.PictureBox picLogos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   480
         Picture         =   "splash.frx":0A84
         ScaleHeight     =   735
         ScaleWidth      =   3315
         TabIndex        =   15
         Top             =   3210
         Width           =   3315
      End
      Begin VB.Label lblSponsors 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "lblSponsors"
         Height          =   195
         Left            =   1650
         TabIndex        =   17
         Top             =   1950
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "E&xit"
      Height          =   525
      Left            =   7650
      TabIndex        =   2
      Top             =   5070
      Width           =   1305
   End
   Begin VB.CommandButton cmdButton2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "I agree, never show again"
      Height          =   525
      Left            =   9060
      TabIndex        =   1
      Top             =   5310
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdButton1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Continue"
      Height          =   525
      Left            =   390
      TabIndex        =   0
      Top             =   5070
      Width           =   1455
   End
   Begin Threed.SSPanel sspanel_maintitle 
      Height          =   735
      Left            =   390
      TabIndex        =   3
      Top             =   120
      Width           =   8475
      _Version        =   65536
      _ExtentX        =   14949
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Ion Exchange Design Software (IonExDesignS    )"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TM"
         Height          =   195
         Left            =   7230
         TabIndex        =   9
         Top             =   120
         Width           =   285
      End
   End
   Begin Threed.SSPanel panelAuthors 
      Height          =   1845
      Left            =   4920
      TabIndex        =   10
      Top             =   2970
      Width           =   4035
      _Version        =   65536
      _ExtentX        =   7117
      _ExtentY        =   3254
      _StockProps     =   15
      Caption         =   "panelAuthors"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1515
      Left            =   4920
      TabIndex        =   11
      Top             =   1200
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   2672
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
      Begin VB.Label lblVersionInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copyright {years}"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   14
         Top             =   1050
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
         TabIndex        =   13
         Top             =   630
         Width           =   3645
      End
      Begin VB.Label lblVersionInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Version {ver} {STANDARD}"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   210
         Width           =   3645
      End
   End
   Begin VB.Label lblAdditionalNotice 
      Alignment       =   2  'Center
      Caption         =   "{This program is protected by .... see code in Form_Load}"
      Height          =   465
      Left            =   2040
      TabIndex        =   18
      Top             =   5100
      Width           =   5385
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
'Call debug_output("L1")
  Me.Height = 5900
  Me.Width = 9300
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
'Call debug_output("L2")
  If (splash_mode = 0) Then
    'SHOW THE CONTINUE/EXIT FRONT WINDOW.
'Call debug_output("L3")
    cmdButton1.Visible = True
    cmdButton1.Caption = "&Continue"
    cmdButton2.Visible = False
    cmdExit.Visible = True
    'VERSION INFO.
    lblVersionInfo(0).Caption = "Version " & get_program_version_with_build_info()
    'EXPIRATION INFO.
    lblVersionInfo(1).Caption = get_expiration_info()
    'COPYRIGHT INFO.
    lblVersionInfo(2).Caption = "Copyright " & AppCopyrightYears
    'ETC.
    sspanel_disclaimer.Visible = False
    'cmdButton1.SetFocus
    cmdButton1.TabIndex = 0
    panelAuthors.Caption = "David R. Hokanson" & Chr$(13) & Chr$(13) & "David W. Hand" & Chr$(13) & Chr$(13) & "John C. Crittenden"
    lblSponsors.Caption = "National Center for" & Chr$(13) & "Clean Industrial and Treatment Technologies" & Chr$(13) & Chr$(13) & "Michigan Technological University" & Chr$(13) & "Houghton, Michigan"
    lblAdditionalNotice.Caption = _
      "This program is protected by U.S. and international" & _
      vbCrLf & _
      "copyright laws as described in Help About."

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

