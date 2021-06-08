VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "{Generic Application -- Me.Caption set as Name_App_Short}"
   ClientHeight    =   1695
   ClientLeft      =   900
   ClientTop       =   3570
   ClientWidth     =   5640
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
   Icon            =   "Main.frx":0000
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   Begin VB.CommandButton cmdPoreDiffusionModel 
      Caption         =   "cmdPoreDiffusionModel"
      Height          =   1065
      Left            =   180
      TabIndex        =   17
      Top             =   120
      Width           =   5265
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2805
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   4948
      _StockProps     =   14
      Caption         =   "Test Frame"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1545
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2250
         TabIndex        =   14
         Text            =   "txtData(3)"
         Top             =   1650
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1545
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2250
         TabIndex        =   11
         Text            =   "txtData(2)"
         Top             =   1230
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   780
         Width           =   1545
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2250
         TabIndex        =   8
         Text            =   "txtData(1)"
         Top             =   810
         Width           =   1995
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2250
         TabIndex        =   6
         Text            =   "txtData(0)"
         Top             =   390
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Flow Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   300
         TabIndex        =   16
         Top             =   1680
         Width           =   1845
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Mass:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   13
         Top             =   1260
         Width           =   1845
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Diameter:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   840
         Width           =   1845
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   420
         Width           =   1845
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   1290
      Width           =   5640
      _Version        =   65536
      _ExtentX        =   9948
      _ExtentY        =   714
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
      Begin Threed.SSPanel sspanel_Dirty 
         Height          =   285
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Dirty"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspanel_Status 
         Height          =   285
         Left            =   2220
         TabIndex        =   2
         Top             =   60
         Width           =   7185
         _Version        =   65536
         _ExtentX        =   12674
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Status"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1035
      Left            =   5880
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   1826
      _StockProps     =   14
      Caption         =   "Used -- Invisible"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open ..."
         Index           =   1
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As ..."
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Select P&rinter"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print ..."
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   190
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&1 Old File #1"
         Index           =   191
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&2 Old File #2"
         Index           =   192
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&3 Old File #3"
         Index           =   193
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&4 Old File #4"
         Index           =   194
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   199
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   200
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Online Help ..."
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Online Manual ..."
         Enabled         =   0   'False
         Index           =   20
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Version History ..."
         Index           =   80
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Disclaimer ..."
         Index           =   85
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Development Notes ..."
         Index           =   90
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   98
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About ..."
         Index           =   99
      End
   End
   Begin VB.Menu mnuMTU 
      Caption         =   "&MTU Internal"
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Make menu invisible"
         Index           =   198
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Read me"
         Index           =   199
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Const frmMain_declarations_end = True


Sub Avoid_Weird_Focus_Problem()
  Call unitsys_control_MostRecent_Force_lostfocus
  'frmMain.SetFocus
  '
  ' NOTE: IT IS VERY IMPORTANT TO SET FOCUS HERE
  ' TO SOME NON-UNITTEXTBOX CONTROL, I.E. DON'T
  ' SET IT TO txtData(0...3), BUT cboUnits(0)
  ' IS OKAY.
  cboUnits(0).SetFocus
  'Text1.SetFocus
End Sub


Sub Populate_frmMain_Units()
  Call unitsys_register(frmMain, lblData(0), txtData(0), cboUnits(0), "length", _
      "m", "m", "", "", 100#, True)
  Call unitsys_register(frmMain, lblData(1), txtData(1), cboUnits(1), "length", _
      "m", "m", "", "", 100#, True)
  Call unitsys_register(frmMain, lblData(2), txtData(2), cboUnits(2), "mass", _
      "kg", "kg", "", "", 100#, True)
  Call unitsys_register(frmMain, lblData(3), txtData(3), cboUnits(3), "flow_volumetric", _
      "m³/s", "m³/s", "", "", 100#, True)
End Sub


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Private Sub cmdPoreDiffusionModel_Click()
Dim ShellVar As Double

''''  ShellVar = Shell("ionex.exe", vbNormalFocus)
  ShellVar = Shell("pdm.exe", vbNormalFocus)

End Sub

Private Sub Form_Load()
Dim is_internal_mtu As Boolean
  '
  ' MISC INITS.
  '
  Call Local_DirtyStatus_Set(Project_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  Me.Caption = Name_App_Short
  Me.Width = 5760
  Me.Height = 2385
  Call CenterOnScreen(Me)
  ''''CommonDialog1.filename = App.Path & "\examples\*.dat"
  CommonDialog1.FileName = _
      App.Path & "\examples\*." & FileExt_App
  '
  ' CHECK FOR FILE THAT INDICATES THIS IS INTERNAL TO MTU:
  '
  is_internal_mtu = False
  If (check_internal_to_mtu()) Then is_internal_mtu = True
  mnuMTU.Visible = is_internal_mtu
  '
  ' POPULATE UNITS INTO SCROLLBOX CONTROLS.
  '
  Call Populate_frmMain_Units
  '
  ' CREATE A NEW FILE IN MEMORY.
  '
  Call file_new
  '
  ' POPULATE LAST-FEW-FILES LIST.
  '
  Call OldFileList_Populate( _
      1, _
      frmMain.mnuFileItem(199), _
      frmMain.mnuFileItem(191), _
      frmMain.mnuFileItem(192), _
      frmMain.mnuFileItem(193), _
      frmMain.mnuFileItem(194))
      
''''added hokanson 6/19/02
   cmdPoreDiffusionModel.FontSize = "14"
   cmdPoreDiffusionModel.Caption = "Pore Diffusion Model"

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (file_query_unload() = False) Then
    Cancel = True
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call frmMain_Close_All_Windows
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub mnuFileItem_Click(index As Integer)
  Select Case index
    Case 0:      'New
      If (file_query_unload()) Then
        Call Avoid_Weird_Focus_Problem
        Call file_new
      End If
    Case 1:      'Open ...
      If (file_query_unload()) Then
        Call Avoid_Weird_Focus_Problem
        Call File_OpenAs("")
      End If
    Case 2:      'Save
      If (Current_Filename = "") Then
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs("")
      Else
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs(Current_Filename)
      End If
    Case 3:      'Save As ...
      Call Avoid_Weird_Focus_Problem
      Call File_SaveAs("")
    Case 6:       'Select Printer ...
      CommonDialog1.ShowPrinter
    'Case 85:      'Print ...
    '  frmPrint.Show 1
    Case 191 To 194:      'Last-few-files list
      If (file_query_unload()) Then
        If (mnuFileItem(index).Visible) Then
          Call File_OpenAs(OldFiles(1, index - 190))
        End If
      End If
    Case 200:     'Exit
      'NOTE: MDIForm_QueryUnload() TAKES CAKE OF THIS.
      'If we do it here, _two_ message boxes will pop up
      'when the user has data which needs saving !
      'If (file_query_unload()) Then
      '  Unload Me
      'End If
      Unload Me
  End Select
End Sub
Private Sub mnuHelpItem_Click(index As Integer)
Dim fn_This As String
  Select Case index
    Case 10:      'ONLINE HELP.
      'NOTE: We currently do NOT have the resources to
      'create an online help file for AdDesignS (1/7/98)
      'therefore no online help is available.
      Call Show_Message("Online help is currently unavailable.  " & _
          "Please refer to the printed manual or the Acrobat-format ADS.PDF file.")
      Exit Sub
      'Call LaunchFile_General("", MAIN_APP_PATH & "\help\ads.hlp")
    Case 20:      'ONLINE MANUAL.
      fn_This = MAIN_APP_PATH & "\help\ads.pdf"
      If (FileExists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call LaunchFile_General("", fn_This)
    Case 80:
      fn_This = App.Path & "\dbase\readme.txt"
      If (FileExists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call Launch_Notepad(fn_This)
    Case 85:    'VIEW DISCLAIMER.
      'SHOW THE DISCLAIMER WINDOW.
      splash_mode = 101
      splash_button_pressed = 0
      frmSplash.Show 1
    Case 99:    'ABOUT.
      frmAbout.Show 1
  End Select
End Sub
Private Sub mnuMTUItem_Click(index As Integer)
  Select Case index
    Case 40:    'KEEP TEMPORARY MODEL FILES.
      mnuMTUItem(40).Checked = Not mnuMTUItem(40).Checked
    Case 198:   'MAKE INVISIBLE.
      mnuMTU.Visible = False
    Case 199:   'READ ME.
      Call Show_Message("This menu should only appear on internal " & _
          "testing machines at MTU.  To remove the `MTU Internal` " & _
          "menu, select `Make menu invisible`.  This will make " & _
          "the menu invisible until the program is closed and reloaded.")
  End Select
End Sub


Private Sub cboUnits_Click(index As Integer)
Dim Ctl As Control
Set Ctl = cboUnits(index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub cboUnits_KeyPress(index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub txtData_GotFocus(index As Integer)
Dim Ctl As Control
Set Ctl = txtData(index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case index
    Case 0
      StatusMessagePanel = "Type in the bed diameter"
    Case 1
      StatusMessagePanel = "Type in the bed length"
    Case 2
      StatusMessagePanel = "Type in the mass of adsorbent in the bed"
    Case 3
      StatusMessagePanel = "Type in the inlet flowrate"
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtData_KeyPress(index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtData_LostFocus(index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtData(index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  If (index = 4) Then
    Val_Low = 1E-20 * 60#
    Val_High = 1E+20 * 60#
  Else
    Val_Low = 1E-20
    Val_High = 1E+20
  End If
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case index
        Case 0:     'BED LENGTH.
          NowProj.length = NewValue
        Case 1:     'BED DIAMETER.
          NowProj.Diameter = NewValue
        Case 2:     'BED MASS.
          NowProj.Mass = NewValue
        Case 3:     'BED FLOW RATE.
          NowProj.FlowRate = NewValue
      End Select
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set( _
            Project_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmMain_Refresh
    End If
  End If
End Sub



