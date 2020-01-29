VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmTitle 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5760
   ClientLeft      =   5055
   ClientTop       =   1755
   ClientWidth     =   9180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5760
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdButton2 
      Caption         =   "I agree, never show again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2130
      TabIndex        =   18
      Top             =   5190
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7140
      TabIndex        =   17
      Top             =   5190
      Width           =   1515
   End
   Begin VB.CommandButton cmdButton1 
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   510
      TabIndex        =   0
      Top             =   5190
      Width           =   1515
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1035
      Left            =   540
      TabIndex        =   1
      Top             =   120
      Width           =   8115
      _Version        =   65536
      _ExtentX        =   14314
      _ExtentY        =   1826
      _StockProps     =   15
      Caption         =   "Advanced Oxidation Process Software"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel sspanel_logos 
      Height          =   3855
      Left            =   540
      TabIndex        =   2
      Top             =   1260
      Width           =   8115
      _Version        =   65536
      _ExtentX        =   14314
      _ExtentY        =   6800
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1725
         Left            =   6060
         Picture         =   "title.frx":0000
         ScaleHeight     =   1725
         ScaleWidth      =   1920
         TabIndex        =   12
         Top             =   1950
         Width           =   1920
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         Picture         =   "title.frx":077A
         ScaleHeight     =   1695
         ScaleWidth      =   1725
         TabIndex        =   11
         Top             =   1980
         Width           =   1725
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   1931
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
         Begin VB.Label Label6 
            Caption         =   "Eric J. Oman"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2220
            TabIndex        =   9
            Top             =   600
            Width           =   5535
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Development:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label4 
            Caption         =   "Shumin Hu      David W. Hand      John C. Crittenden"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2220
            TabIndex        =   7
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Model and Software:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1995
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1155
         Left            =   2340
         TabIndex        =   10
         Top             =   2100
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   2037
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
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Houghton, Michigan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   16
            Top             =   750
            Width           =   3255
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Michigan Technological University"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   540
            Width           =   3255
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Copyright 1997-1998"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   14
            Top             =   330
            Width           =   3255
         End
         Begin VB.Label lblVersion 
            Alignment       =   2  'Center
            Caption         =   "Version {ver}"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   13
            Top             =   120
            Width           =   3255
         End
      End
      Begin VB.Shape Shape2 
         Height          =   1785
         Left            =   6030
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Shape Shape1 
         Height          =   1695
         Left            =   120
         Top             =   1980
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Center for Clean Industrial and Treatment Technologies"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   7995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CenCITT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   7995
      End
   End
   Begin Threed.SSPanel sspanel_disclaimer 
      Height          =   3855
      Left            =   5490
      TabIndex        =   19
      Top             =   5880
      Width           =   8115
      _Version        =   65536
      _ExtentX        =   14314
      _ExtentY        =   6800
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
      Begin VB.Label lblDisclaimer 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   450
         TabIndex        =   21
         Top             =   600
         Width           =   7185
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Disclaimer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   60
         TabIndex        =   20
         Top             =   90
         Width           =   7995
      End
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'splash_mode: 0 = Continue/Exit window
'             1 = I Agree/I agree, never show again/Exit window
Dim splash_mode As Integer

'splash_button_pressed:
'1 = Continue or I Agree
'2 = I agree, never show again
'3 = Exit
Dim splash_button_pressed As Integer

Public Function Run() As Boolean
Dim tpath$
Dim tstr$
Dim must_read_disclaimer As Integer

  '''SET UP INI FILE PATH.
  ''tpath$ = GetWindowsDir() & ProgramIniFile$
  
  'SHOW THE CONTINUE/EXIT FRONT WINDOW.
  splash_mode = 0
  splash_button_pressed = 0
  frmTitle.Show 1
  Select Case splash_button_pressed
    Case 1:         'Hit Continue
      'DO NOTHING.
    Case 3:         'Hit Exit
      End
  End Select
    
  'IS THE DISCLAIMER WINDOW STILL ACTIVE?
  must_read_disclaimer = True
  tstr$ = INI_GetSetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer")
  If (tstr$ = "1") Then
    must_read_disclaimer = False
  End If
  
  If (must_read_disclaimer) Then
    'SHOW THE DISCLAIMER WINDOW.
    splash_mode = 1
    splash_button_pressed = 0
    frmTitle.Show 1
    Select Case splash_button_pressed
      Case 1:         'Hit I Agree
        'DO NOTHING.
      Case 2:         'Hit I agree, never show again
        Call ini_putsetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer", "1")
      Case 3:         'Hit Exit
        End
    End Select
  End If

  Run = True
  
End Function


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
  Call CenterOnScreen(Me)
  
  If (splash_mode = 0) Then
    'SHOW THE CONTINUE/EXIT FRONT WINDOW.
    cmdButton1.visible = True
    cmdButton1.Caption = "&Continue"
    cmdButton2.visible = False
    cmdExit.visible = True
    lblVersion.Caption = "Version " & get_program_version_with_build_info()
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
  End If
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  If (splash_button_pressed = 0) Then
    'If they got here, they must have selected "Close",
    'so perform the exit functionality.
    splash_button_pressed = 3
  End If
  
End Sub


