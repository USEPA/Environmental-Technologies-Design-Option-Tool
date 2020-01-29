VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmMainMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aeration System Analysis Program (ASAP)"
   ClientHeight    =   6510
   ClientLeft      =   3645
   ClientTop       =   2070
   ClientWidth     =   9495
   DrawStyle       =   1  'Dash
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9495
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   6030
      Width           =   2325
   End
   Begin Threed.SSPanel pnl_main 
      Height          =   1575
      Index           =   1
      Left            =   2340
      TabIndex        =   7
      Top             =   810
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8700
      _ExtentY        =   2773
      _StockProps     =   15
      Caption         =   "Packed Tower Aeration"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Begin VB.CommandButton cmd_titlemethod 
         Caption         =   "Enter Rating Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   900
         TabIndex        =   1
         Top             =   1020
         Width           =   3855
      End
      Begin VB.CommandButton cmd_titlemethod 
         Caption         =   "Enter Design Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   900
         TabIndex        =   0
         Top             =   510
         Width           =   3855
      End
      Begin VB.Label lbl_number 
         Caption         =   "2."
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
         Index           =   1
         Left            =   420
         TabIndex        =   10
         Top             =   1095
         Width           =   435
      End
      Begin VB.Label lbl_number 
         Caption         =   "1."
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
         Index           =   0
         Left            =   420
         TabIndex        =   9
         Top             =   585
         Width           =   435
      End
   End
   Begin Threed.SSPanel pnl_main 
      Height          =   585
      Index           =   0
      Left            =   2730
      TabIndex        =   8
      Top             =   30
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   "Main Menu"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel pnl_main 
      Height          =   1575
      Index           =   2
      Left            =   2340
      TabIndex        =   11
      Top             =   2520
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8700
      _ExtentY        =   2773
      _StockProps     =   15
      Caption         =   "Bubble Aeration"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Begin VB.CommandButton cmd_titlemethod 
         Caption         =   "Enter Design Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   900
         TabIndex        =   2
         Top             =   510
         Width           =   3855
      End
      Begin VB.CommandButton cmd_titlemethod 
         Caption         =   "Enter Rating Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   900
         TabIndex        =   3
         Top             =   1020
         Width           =   3855
      End
      Begin VB.Label lbl_number 
         Caption         =   "3."
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
         Index           =   2
         Left            =   420
         TabIndex        =   13
         Top             =   585
         Width           =   435
      End
      Begin VB.Label lbl_number 
         Caption         =   "4."
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
         Index           =   3
         Left            =   420
         TabIndex        =   12
         Top             =   1095
         Width           =   435
      End
   End
   Begin Threed.SSPanel pnl_main 
      Height          =   1575
      Index           =   3
      Left            =   2340
      TabIndex        =   14
      Top             =   4230
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8700
      _ExtentY        =   2773
      _StockProps     =   15
      Caption         =   "Surface Aeration"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Begin VB.CommandButton cmd_titlemethod 
         Caption         =   "Enter Rating Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   900
         TabIndex        =   5
         Top             =   1020
         Width           =   3855
      End
      Begin VB.CommandButton cmd_titlemethod 
         Caption         =   "Enter Design Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   900
         TabIndex        =   4
         Top             =   510
         Width           =   3855
      End
      Begin VB.Label lbl_number 
         Caption         =   "6."
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
         Index           =   5
         Left            =   420
         TabIndex        =   16
         Top             =   1095
         Width           =   435
      End
      Begin VB.Label lbl_number 
         Caption         =   "5."
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
         Index           =   4
         Left            =   420
         TabIndex        =   15
         Top             =   585
         Width           =   435
      End
   End
   Begin VB.Menu mnuModels 
      Caption         =   "&Models"
      Begin VB.Menu mnuModelsItem 
         Caption         =   "&1 -- Packed Tower Aeration (Design Mode)"
         Index           =   10
      End
      Begin VB.Menu mnuModelsItem 
         Caption         =   "&2 -- Packed Tower Aeration (Rating Mode)"
         Index           =   20
      End
      Begin VB.Menu mnuModelsItem 
         Caption         =   "&3 -- Bubble Aeration (Design Mode)"
         Index           =   30
      End
      Begin VB.Menu mnuModelsItem 
         Caption         =   "&4 -- Bubble Aeration (Rating Mode)"
         Index           =   40
      End
      Begin VB.Menu mnuModelsItem 
         Caption         =   "&5 -- Surface Aeration (Design Mode)"
         Index           =   50
      End
      Begin VB.Menu mnuModelsItem 
         Caption         =   "&6 -- Surface Aeration (Rating Mode)"
         Index           =   60
      End
      Begin VB.Menu mnuModelsItem 
         Caption         =   "-"
         Index           =   198
      End
      Begin VB.Menu mnuModelsItem 
         Caption         =   "E&xit"
         Index           =   199
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Online Help ..."
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Online Manual ..."
         Index           =   6
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Manual Printing Instructions ..."
         Index           =   7
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Version History ..."
         Index           =   10
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Disclaimer ..."
         Index           =   20
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Technical Assistance Provided By ..."
         Index           =   30
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   190
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About ASAP ..."
         Index           =   200
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Whether Password is needed to use program
'   False --> No Password Needed
'   True --> Password needed
Const Password_Protection_On = False

Dim disclaimer_agreed As Integer
Dim frmMainMenu_Okay_To_Unload As Boolean



Const frmMainMenu_declarations_end = True


Private Sub cmd_passwd_Click_OLD()
'If StrComp(decrypt_string("V\;>Sa["), txt_passwd.Text, 1) Then
'    Unload Me
'    End
'Else
'    Frm_Passwd.Visible = False
'End If
End Sub

Private Sub cmd_title_Click_OLD(Index As Integer)
'  Dim Check_has_seen_disclaimer As String
'  Dim msg$
'
''check_area
'Select Case Index
'Case 0          '// Continue or I Agree button
'    frame_title.Visible = False
'
'    If (disclaimer_agreed = False) Then
'      Check_has_seen_disclaimer = INI_Getsetting("has_seen_disclaimer")
'      If (Check_has_seen_disclaimer = "1") Then
'        disclaimer_agreed = True
'      End If
'    End If
'    If (disclaimer_agreed = False) Then
'      If (Not DemoMode%) And Password_Protection_On Then
'        Frm_Passwd.Top = -400
'        Frm_Passwd.Left = -40
'        Frm_Passwd.Visible = True
'        Frm_Passwd.ZOrder
'        pnl_passwd.Caption = "This Program Needs a Password to Work!" + Chr(13) + "If You Do Not Know the Password." + Chr(13) + "You should not be using the Program."
'        txt_passwd.SetFocus
'      End If
'
'      msg$ = ""
'
'msg$ = "By choosing " & Chr$(34) & "I Agree" & Chr$(34) & " you acknowledge that this software is under development and not guaranteed to be free of errors.  Furthermore there may be errors in the software that lead to erroneous output.  MTU shall not be liable for any loss, damage, injury, or casualty of whatsoever kind, or by whomsoever caused to the person or property of anyone arising out of or resulting from receipt and use of any aspect of the software.  References to specific commercial products, processes, or services by trademark, manufacturer, or otherwise does not necessarily constitute or imply endorsement/recommendation by the authors or the respective organizations under which the software was developed."
'liability.Caption = msg$
'
'      Disclaimer.ZOrder
'      Disclaimer.Visible = True
'      Disclaimer.Move 0, 0
'      cmd_title(0).Caption = "I Agree"
'      cmd_title(0).ZOrder
'      cmd_title(1).ZOrder
'      disclaimer_agreed = True
'
'      cmddisclaimer.Top = cmd_title(0).Top
'      cmddisclaimer.Left = cmd_title(0).Left + cmd_title(0).Width + 200
'      cmddisclaimer.Visible = True
'      cmddisclaimer.ZOrder
'    Else
'      Disclaimer.Visible = False
'      cmd_title(0).Visible = False
'      cmd_title(1).Visible = False
'      cmddisclaimer.Visible = False
'
'      pnl_main(1).Visible = True
'      pnl_main(2).Visible = True
'      pnl_main(3).Visible = True
'    End If
'
'Case 1
'    Unload Me
'    End
'End Select
End Sub

Private Sub cmd_titlemethod_Click(Index As Integer)
    Dim i%

Screen.MousePointer = 11

Select Case Index
    Case 0   ' Packed Tower Design Mode
        CurrMethod% = 0
        CurrMode% = 0
        frmPTADScreen1.Show
        If (StartScreen1DefaultCase() = False) Then
          frmPTADScreen1.Hide
          GoTo exit_Normalize_Mouse_Pointer
        End If
    
    Case 1   ' Packed Tower Rating Mode
        CurrMethod% = 0
        CurrMode% = 1
        ShownScreen1Previously = False
        frmPTADScreen2.Show
        ''''Call StartScreen2DefaultCase
        If (StartScreen2DefaultCase() = False) Then
          frmPTADScreen2.Hide
          GoTo exit_Normalize_Mouse_Pointer
        End If
    
    Case 2   'Bubble Aeration Design Mode
        CurrMethod% = 1
        CurrMode = 0
        frmBubble.Caption = "Bubble Aeration - Design Mode"
        frmBubble!mnuFile(0).Caption = "Switch to &Rating Mode"
        BubbleAerationMode = DESIGN_MODE
        For i% = 1 To 4
            frmBubble!txtTankParameters(i%).Enabled = False
        Next i%
        frmBubble.Show
        ''''Call StartBubbleDefaultCase
        If (StartBubbleDefaultCase() = False) Then
          ''''frmBubble.Hide
          frmBubble.frmBubble_Okay_To_Unload = True
          Unload frmBubble
          GoTo exit_Normalize_Mouse_Pointer
        End If

    Case 3   'Bubble Aeration Rating Mode
        CurrMethod% = 1
        CurrMode% = 1
        frmBubble.Caption = "Bubble Aeration - Rating Mode"
        frmBubble!mnuFile(0).Caption = "Switch to &Design Mode"
        BubbleAerationMode = RATING_MODE
        bub.TankVolume.UserInput = True
        For i% = 1 To 4
            frmBubble!txtTankParameters(i%).Enabled = True
        Next i%
        frmBubble.Show
        ''''Call StartBubbleDefaultCase
        If (StartBubbleDefaultCase() = False) Then
          ''''frmBubble.Hide
          frmBubble.frmBubble_Okay_To_Unload = True
          Unload frmBubble
          GoTo exit_Normalize_Mouse_Pointer
        End If

    Case 4   'Surface Aeration Design Mode
        CurrMethod% = 2
        CurrMode% = 0
        frmSurface.Caption = "Surface Aeration - Design Mode"
        frmSurface!mnuFile(0).Caption = "Switch to &Rating Mode"
        SurfaceAerationMode = DESIGN_MODE
        For i% = 1 To 4
            frmSurface!txtTankParameters(i%).Enabled = False
        Next i%
        frmSurface.Show
        ''''Call StartSurfaceDefaultCase
        If (StartSurfaceDefaultCase() = False) Then
          ''''frmSurface.Hide
          frmSurface.frmSurface_Okay_To_Unload = True
          Unload frmSurface
          GoTo exit_Normalize_Mouse_Pointer
        End If
    
    Case 5   'Surface Aeration Rating Mode
        CurrMethod% = 2
        CurrMode% = 1
        frmSurface.Caption = "Surface Aeration - Rating Mode"
        frmSurface!mnuFile(0).Caption = "Switch to &Design Mode"
        SurfaceAerationMode = RATING_MODE
        sur.TankHydraulicRetentionTime.UserInput = True
        For i% = 1 To 4
            frmSurface!txtTankParameters(i%).Enabled = True
        Next i%
        frmSurface.Show
        ''''Call StartSurfaceDefaultCase
        If (StartSurfaceDefaultCase() = False) Then
          ''''frmSurface.Hide
          frmSurface.frmSurface_Okay_To_Unload = True
          Unload frmSurface
          GoTo exit_Normalize_Mouse_Pointer
        End If
End Select

'frame_title.Visible = False
frmMainMenu.Hide
exit_Normalize_Mouse_Pointer:
Screen.MousePointer = 0

End Sub

Private Sub cmddisclaimer_Click_OLD()
'  Call INI_PutSetting("has_seen_disclaimer", "1")
'  Call cmd_title_Click(0)
End Sub

Private Sub cmdExit_Click()
  frmMainMenu_Okay_To_Unload = True
  Unload Me
End Sub

Private Sub Form_Load()
Dim i%
Dim duh As Integer
Dim version_text As String

  frmMainMenu.Width = SCREEN_WIDTH_STANDARD
  frmMainMenu.Height = SCREEN_HEIGHT_STANDARD
  'Center the form on the screen
  If WindowState = 0 Then
    'don't attempt if screen Minimized or Maximized
    Move (Screen.Width - frmMainMenu.Width) / 2, (Screen.Height - frmMainMenu.Height) / 2
  End If
      
  'Disclaimer.Visible = False
  'cmd_title(0).Visible = False
  'cmd_title(1).Visible = False
  'cmddisclaimer.Visible = False

  For i% = 0 To 3
    pnl_main(i%).Left = frmMainMenu.Width / 2 - pnl_main(i%).Width / 2
  Next i%
  'cmdExit.Move frame_title.Width / 2 - cmdExit.Width / 2
  pnl_main(1).Visible = True
  pnl_main(2).Visible = True
  pnl_main(3).Visible = True
  
  Exit Sub





Screen.MousePointer = 11

'---- Setup helpfiles
If (fileexists(App.Path & "\help\asap.hlp")) Then App.HelpFile = App.Path & "\help\asap.hlp"
    
''''If (Not security_ok()) Then Exit Sub

If (StudentMode%) Then
    If (Not fileexists(decrypt_string("es#52k#>h5>2VO52k"))) Then
        MsgBox "MTU Version Only!!", 16
        End
    End If
End If


pnl_main(1).Visible = False
pnl_main(2).Visible = False
pnl_main(3).Visible = False

NL = Chr$(13) & Chr$(10)
disclaimer_agreed = False

''''ChDrive App.Path
''''ChDir App.Path
Call ChangeDir_Main
SaveAndLoadPath = App.Path


'Frm_Passwd.Visible = False
frmMainMenu.Width = SCREEN_WIDTH_STANDARD
frmMainMenu.Height = SCREEN_HEIGHT_STANDARD

'Center the form on the screen
If WindowState = 0 Then
  'don't attempt if screen Minimized or Maximized
  Move (Screen.Width - frmMainMenu.Width) / 2, (Screen.Height - frmMainMenu.Height) / 2
End If

For i% = 0 To 3
  pnl_main(i%).Left = frmMainMenu.Width / 2 - pnl_main(i%).Width / 2
Next i%

'frame_title.ZOrder
'frame_title.Move -50, -200

'pnl_title(0).Move frame_title.Width / 2 - pnl_title(0).Width / 2
'pnl_title(1).Move frame_title.Width / 2 - pnl_title(1).Width / 2
'pnl_title(2).Move frame_title.Width / 2 - pnl_title(2).Width / 2
'cmdExit.Move frame_title.Width / 2 - cmdExit.Width / 2

'cmd_title(0).Top = pnl_title(2).Top + pnl_title(2).Height - 150 + 100
'cmd_title(0).Left = 240
'cmd_title(0).ZOrder
'cmd_title(1).Top = cmd_title(0).Top
'cmd_title(1).Left = SCREEN_WIDTH_STANDARD - cmd_title(1).Width - 240 - 240
'cmd_title(1).ZOrder

'pnl_title(0).Caption = "A S A P" & Chr$(13) & "Aeration System Analysis Program"
'pnl_title(1).Caption = "Packed Tower Aeration" & Chr$(13) & Chr$(13) & "Bubble Aeration" & Chr$(13) & Chr$(13) & "Surface Aeration"
'pnl_title(2).Caption = "CenCITT" & Chr$(13) & "Center for Clean Industrial and Treatment Technologies"
'pnl_title(3).Caption = "Model && Software Developers:" & Chr$(13) & "David R. Hokanson     David W. Hand     John C. Crittenden"
If (StudentMode%) Then
  version_text = "Version 1.00(S)"
Else
  version_text = "Version 1.00"
End If
'pnl_title(4).Caption = version_text & Chr$(13) & Chr$(13) & "Copyright 1993-1997" & Chr$(13) & "Michigan Technological University" & Chr$(13) & "Houghton, Michigan"
    
' -------------------------------------------------------------------------------------
'Initialize Default Power Variables
scr1.Power.BlowerEfficiency = 35#
scr1.Power.PumpEfficiency = 80#
Scr2.Power.BlowerEfficiency = 35#
Scr2.Power.PumpEfficiency = 80#

bub.Power.BlowerEfficiency = 35#
bub.Power.TankWaterDepth = 4#
bub.Power.NumberOfBlowersinEachTank = 1

'picMtu.Picture = LoadPicture(app.Path & "\mtu_logo.bmp")
'picCencitt.Picture = LoadPicture(app.Path & "\cencitt1.bmp")

Me.Show

ReadMainPackingDB
ReadUserPackingDB

''''Call ini_initializethisprogram("asap")

Screen.MousePointer = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (frmMainMenu_Okay_To_Unload) Then
    Cancel = False
  Else
    Cancel = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ''''On Error Resume Next

    'Forms for Packed Tower Aeration
    Unload frmAirWaterProperties
    'Unload frmListContaminant
    'Unload frmPropContaminant
    ''''Unload HelpTipForm
    Unload frmShowOndaKLaProperties
    Unload frmSelectPacking
    Unload frmShowPackingProperties
    Unload frmPower
    Unload frmPowerScreen2
    Unload frmFlowsLoadingsScreen2
    'Unload frmListcontaminantScreen2
    'Unload frmPropContaminantScreen2
    Unload frmOptimizeContaminant
    Unload frmPTADScreen1
    Unload frmPTADScreen2

    'Forms for Bubble Aeration
    Unload frmBubble
    Unload frmBubbleEffluentConcentrations
    Unload frmOxygenMassTransferCoeff
    'Unload frmListContaminantBubble
    Unload frmBubblePower
    'Unload frmPropContaminantBubble
    Unload frmBubbleAchievingRemovalEfficiency
    Unload frmWaterPropertiesBubble

    'Forms for Surface Aeration
    Unload frmSurface
    Unload frmSurfaceEffluentConcentrations
    'Unload frmListContaminantSurface
    'Unload frmPropContaminantSurface
    Unload frmWaterPropertiesSurface

    'Forms for all three modules
    ''''Unload frmViewContaminantPropertiesPTAD
    Unload frmViewEffluentConcentrationsASAP

    '
    ' CLOSE ANY WINDOWS WE MISSED.
    '
    Call Close_All_Windows

    '
    ' END THE PROGRAM.
    '
    frmMainMenu_Okay_To_Unload = True
    End

End Sub


'Private Function security_ok() As Integer
'Dim duh As Integer
''---- If distributing via CD, check if this is a valid
''.... copy of the ADSIM program:
''.... (Note: this global var is set in DEMOMODE.BAS)
'
'If (Mode_Distribution_on_CD) Then
'  If ((Not StudentMode) Or (Not fileexists("R:\esbeam\data\tack\esbeam.tck"))) Then
'    If (Not fileexists(App.Path & "\5YVJ058Z.CE3")) Then
'      If (fileexists(GetWindowsDir() & "\etchk.exe")) Then
'        ChDrive GetWindowsDir()
'        ChDir GetWindowsDir()
'        duh = Shell(GetWindowsDir() & "\etchk.exe", 1)
'          DoEvents
'        Do While (True)
'          DoEvents
'          If (fileexists(GetWindowsDir() & "\exit.x")) Then
'            Kill GetWindowsDir() & "\exit.x"
'            DoEvents
'            End
'          End If
'          DoEvents
'          If (fileexists(GetWindowsDir() & "\go.x")) Then
'            Kill GetWindowsDir() & "\go.x"
'            DoEvents
'            Exit Do
'          End If
'        Loop
'      Else
'        security_ok = False
'      End If
'    End If
'  End If
'End If
'
' security_ok = True
' ChDrive App.Path
' ChDir App.Path
'End Function

Private Sub txt_passwd_KeyPress_OLD(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       KeyAscii = 0
'       cmd_passwd_Click
'       Exit Sub
'    End If
End Sub


Private Sub mnuHelpItem_Click(Index As Integer)
  Call Launch_ASAP_mnuHelp_Item(Index)
''''Dim fn_This As String
''''  Select Case Index
'''''    Case 5:       'CONTENTS.
'''''      SendKeys "{F1}", True
''''    Case 5:       'ONLINE HELP.
''''      Call Launch_ASAP_HLP_File
''''    Case 7:       'ONLINE MANUAL.
''''      fn_This = MAIN_APP_PATH & "\help\asap.pdf"
''''      If (fileexists(fn_This) = False) Then
''''        Call Show_Message("The file `" & fn_This & "` is missing.")
''''        Exit Sub
''''      End If
''''      Call LaunchFile_General("", fn_This)
''''      'Call LaunchFile_General("", MAIN_APP_PATH & "\help\asap.pdf")
''''    Case 10:      'VIEW VERSION HISTORY.
''''      fn_This = App.Path & "\dbase\readme.txt"
''''      If (fileexists(fn_This) = False) Then
''''        Call Show_Message("The file `" & fn_This & "` is missing.")
''''        Exit Sub
''''      End If
''''      Call Launch_Notepad(fn_This)
''''    Case 20:      'VIEW DISCLAIMER.
''''      'SHOW THE DISCLAIMER WINDOW.
''''      splash_mode = 101
''''      splash_button_pressed = 0
''''      frmSplash.Show 1
''''    Case 30:      'TECHNICAL ASSISTANCE PROVIDED BY.
''''      frmTechAssistance.Show 1
''''    Case 200:     'ABOUT.
''''      frmAbout.Show 1
''''  End Select
End Sub


Private Sub mnuModelsItem_Click(Index As Integer)
Dim Button_To_Push As Integer
  Button_To_Push = -1
  Select Case Index
    Case 10: Button_To_Push = 0
    Case 20: Button_To_Push = 1
    Case 30: Button_To_Push = 2
    Case 40: Button_To_Push = 3
    Case 50: Button_To_Push = 4
    Case 60: Button_To_Push = 5
    Case 199:
      Call cmdExit_Click
      Exit Sub
  End Select
  If (Button_To_Push <> -1) Then
    Call cmd_titlemethod_Click(Button_To_Push)
    Exit Sub
  End If
End Sub





