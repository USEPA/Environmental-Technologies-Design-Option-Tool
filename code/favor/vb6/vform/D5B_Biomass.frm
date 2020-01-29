VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmD5B_Biomass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Biomass Calculation"
   ClientHeight    =   4320
   ClientLeft      =   6435
   ClientTop       =   2115
   ClientWidth     =   7710
   ControlBox      =   0   'False
   HelpContextID   =   8000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame3 
      Height          =   2955
      Left            =   6360
      TabIndex        =   5
      Top             =   6900
      Visible         =   0   'False
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   5212
      _StockProps     =   14
      Caption         =   "Unused -- Invisible"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1500
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
         Left            =   390
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   750
         Width           =   1500
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2685
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5955
      _Version        =   65536
      _ExtentX        =   10504
      _ExtentY        =   4736
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Index           =   4
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1500
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
         Height          =   360
         Index           =   4
         Left            =   2685
         TabIndex        =   19
         Text            =   "txtData(4)"
         Top             =   2175
         Width           =   1635
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
         Height          =   360
         Index           =   3
         Left            =   2685
         TabIndex        =   17
         Text            =   "txtData(3)"
         Top             =   1515
         Width           =   1635
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
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1500
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
         Height          =   360
         Index           =   2
         Left            =   2685
         TabIndex        =   14
         Text            =   "txtData(2)"
         Top             =   1095
         Width           =   1635
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
         Height          =   360
         Index           =   1
         Left            =   2685
         TabIndex        =   12
         Text            =   "txtData(1)"
         Top             =   675
         Width           =   1635
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
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1500
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
         Height          =   360
         Index           =   0
         Left            =   2685
         TabIndex        =   9
         Text            =   "txtData(0)"
         Top             =   255
         Width           =   1635
      End
      Begin VB.Label lblDataUnits 
         Alignment       =   2  'Center
         Caption         =   "(mg VSS/mg BOD5)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   4350
         TabIndex        =   25
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label lblDataUnits 
         Alignment       =   2  'Center
         Caption         =   "(mg/L BOD5)"
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
         Left            =   4350
         TabIndex        =   24
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Aeration Basin Influent BOD5 Concentration:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1980
         Width           =   2505
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Yield Coefficient:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   2505
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Bacterial Decay Rate:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   2505
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Half Velocity Constant:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2505
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Max. Growth Rate Constant:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   2505
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   3915
      Width           =   7710
      _Version        =   65536
      _ExtentX        =   13600
      _ExtentY        =   714
      _StockProps     =   15
      ForeColor       =   -2147483640
      BackColor       =   12632256
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
         ForeColor       =   -2147483640
         BackColor       =   12632256
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
         Width           =   5000
         _Version        =   65536
         _ExtentX        =   8819
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Status"
         ForeColor       =   -2147483640
         BackColor       =   12632256
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   6270
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes to this window"
      Top             =   540
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   0
      Left            =   6270
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes on this window"
      Top             =   60
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Cancel"
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   2
      Left            =   6270
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Click here for help"
      Top             =   1200
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Help"
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
   Begin VB.Label Label1 
      Caption         =   "Before FaVOr can continue to calculate the basin biomass concentration(s), the following additional parameters are required"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   150
      TabIndex        =   4
      Top             =   90
      Width           =   5865
   End
End
Attribute VB_Name = "frmD5B_Biomass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmD5B_Biomass_Is_Dirty As Boolean

Dim Temp_Plant As TYPE_PlantDiagram




Const frmD5B_Biomass_declarations_end = True


Sub frmD5B_Biomass_Edit( _
    INPUT_UseWhichStructure As Integer, _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  '
  ' TRANSFER GLOBAL STRUCTURE TO LOCAL MEMORY.
  '
  Select Case INPUT_UseWhichStructure
    Case INPUT_UseWhichStructure_D5:
      Temp_Plant = frmD5_AerationBasin_Temp_Plant
    Case INPUT_UseWhichStructure_D5A:
      Temp_Plant = frmD5A_CSTR_Temp_Plant
  End Select
  '
  ' SHOW THE FORM.
  '
  frmD5B_Biomass.Show 1
  If (USER_HIT_OK) Then
    '
    ' TELL CALLER WINDOW TO RAISE DIRTY FLAG.
    '
    OUTPUT_Raise_Dirty_Flag = True
    '
    ' TRANSFER LOCAL STRUCTURE TO GLOBAL MEMORY.
    '
    Select Case INPUT_UseWhichStructure
      Case INPUT_UseWhichStructure_D5:
        frmD5_AerationBasin_Temp_Plant = Temp_Plant
      Case INPUT_UseWhichStructure_D5A:
        frmD5A_CSTR_Temp_Plant = Temp_Plant
    End Select
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub


Sub frmD5B_Biomass_PopulateUnits()
Dim Frm As Form
Set Frm = frmD5B_Biomass
  '
  ' MAIN DATA BLOCK.
  '
  With Temp_Plant.AerationBasin.BioTreat
    Call unitsys_register(Frm, lblData(0), txtData(0), cboUnits(0), "inverse_time", _
        .UnitsOfDisplay(0), "1/day", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(1), txtData(1), Nothing, "", _
        "", "", "", "", 100#, False)
    Call unitsys_register(Frm, lblData(2), txtData(2), cboUnits(2), "inverse_time", _
        .UnitsOfDisplay(2), "1/day", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(3), txtData(3), Nothing, "", _
        "", "", "", "", 100#, False)
    Call unitsys_register(Frm, lblData(4), txtData(4), cboUnits(4), "concentration", _
        .UnitsOfDisplay(4), "mg/L", "", "", 100#, True)
  End With
End Sub
Sub Store_Unit_Settings()
Dim i As Integer
  With Temp_Plant.AerationBasin.BioTreat
    For i = 0 To 4
      .UnitsOfDisplay(i) = unitsys_get_units(cboUnits(i))
    Next i
  End With
End Sub


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Private Sub cboUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub cboUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
Dim CalcSuccess As Boolean
Dim Temp_Store_Plant As TYPE_PlantDiagram
  Select Case Index
    Case 0:     'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      '
      ' RUN THE CALCULATIONS.
      '
      Temp_Store_Plant = NowProj.Plant
      NowProj.Plant = Temp_Plant
      If (ModelBIOCALC_Go(Temp_Plant) = False) Then
        Call Show_Error("Calculation failed.  You cannot " & _
            "save this data.  Modify the existing parameters " & _
            "and hit OK to try another calculation, or hit " & _
            "Cancel to lose the changes you have made to " & _
            "this window.")
        Exit Sub
      End If
      NowProj.Plant = Temp_Store_Plant
''''''      CalcSuccess = CalculateBioMass(Temp_Plant)
''''''      If (CalcSuccess = False) Then
''''''        Call Show_Error("Calculation failed.  You cannot " & _
''''''            "save this data.  Modify the existing parameters " & _
''''''            "and hit OK to try another calculation, or hit " & _
''''''            "Cancel to lose the changes you have made to " & _
''''''            "this window.")
''''''        Exit Sub
''''''      End If
      '
      ' STORE ALL UNIT SETTINGS.
      '
      Call Store_Unit_Settings
      '
      ' EXIT OUT OF HERE.
      '
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
    Case 2:     'HELP.
      SendKeys "{F1}"
  End Select
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  Call CenterOnForm(Me, frmMain)
  Call Local_DirtyStatus_Set(frmD5B_Biomass_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  '
  ' POPULATE UNIT CONTROLS.
  '
  Call frmD5B_Biomass_PopulateUnits
  '
  ' REFRESH DISPLAY.
  '
  Call frmD5B_Biomass_Refresh(Temp_Plant)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
    '
    ' MAIN DATA BLOCK.
    '
    Case 0:
      StatusMessagePanel = ""
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtData_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    ' MAIN DATA BLOCK.
    Case 0: Val_Low = 1E-20: Val_High = 1E+20
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
    Case 2: Val_Low = 1E-20: Val_High = 1E+20
    Case 3: Val_Low = 1E-20: Val_High = 1E+20
    Case 4: Val_Low = 1E-20: Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      With Temp_Plant.AerationBasin.BioTreat
        Select Case Index
          '
          ' MAIN DATA BLOCK.
          '
          Case 0: .MaxGrowthRate = NewValue
          Case 1: .HalfVelocityConst = NewValue
          Case 2: .BacterialDecay = NewValue
          Case 3: .YieldCoeff = NewValue
          Case 4: .BOD5Conc = NewValue
        End Select
      End With
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD5B_Biomass_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD5B_Biomass_Refresh(Temp_Plant)
    End If
  End If
End Sub



