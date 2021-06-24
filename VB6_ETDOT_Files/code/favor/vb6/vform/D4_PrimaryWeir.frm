VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmD4_PrimaryWeir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Parameters [Primary Clarifier Weir]"
   ClientHeight    =   4680
   ClientLeft      =   1470
   ClientTop       =   2025
   ClientWidth     =   7500
   ControlBox      =   0   'False
   HelpContextID   =   5000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   1725
      Left            =   240
      TabIndex        =   11
      Top             =   2340
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   3043
      _StockProps     =   14
      Caption         =   "Channel Specifications:"
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
         Index           =   2
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1140
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
         Height          =   300
         Index           =   2
         Left            =   3015
         TabIndex        =   2
         Text            =   "txtData(2)"
         Top             =   1170
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
         Left            =   3015
         TabIndex        =   0
         Text            =   "txtData(0)"
         Top             =   330
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
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
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
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
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
         Height          =   300
         Index           =   1
         Left            =   3015
         TabIndex        =   1
         Text            =   "txtData(1)"
         Top             =   750
         Width           =   1995
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Gas Flow Rate:"
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
         Left            =   90
         TabIndex        =   17
         Top             =   1200
         Width           =   2805
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Width:"
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
         Left            =   90
         TabIndex        =   15
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Distance betw. Water Levels:"
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
         Left            =   90
         TabIndex        =   14
         Top             =   780
         Width           =   2805
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   4275
      Width           =   7500
      _Version        =   65536
      _ExtentX        =   13229
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
         TabIndex        =   4
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
         TabIndex        =   5
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
      Left            =   5940
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes to this window"
      Top             =   630
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
      Left            =   5940
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes on this window"
      Top             =   150
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
      Left            =   5940
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Click here for help"
      Top             =   1290
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
   Begin Threed.SSFrame SSFrame6 
      Height          =   915
      Left            =   240
      TabIndex        =   9
      Top             =   1290
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Select Modeling Mechanism:"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cbo_Model_Type 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   5235
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"D4_PrimaryWeir.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   300
      TabIndex        =   18
      Top             =   150
      Width           =   4515
   End
End
Attribute VB_Name = "frmD4_PrimaryWeir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmD4_PrimaryWeir_Is_Dirty As Boolean

Dim Temp_Plant As TYPE_PlantDiagram

Public HALT_cbo_Model_Type As Boolean





Const frmD4_PrimaryWeir_declarations_end = True


Sub frmD4_PrimaryWeir_Edit( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  Temp_Plant = NowProj.Plant
  frmD4_PrimaryWeir.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
    NowProj.Plant = Temp_Plant
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub


Sub frmD4_PrimaryWeir_PopulateUnits()
Dim Frm As Form
Set Frm = frmD4_PrimaryWeir
  '
  ' MAIN DATA BLOCK.
  '
  Call unitsys_register(Frm, lblData(0), txtData(0), cboUnits(0), "length", _
      Temp_Plant.PrimaryWeir.UnitsOfDisplay(0), "m", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(1), txtData(1), cboUnits(1), "length", _
      Temp_Plant.PrimaryWeir.UnitsOfDisplay(1), "m", "", "", 100#, True)
  ''''Call unitsys_register(Frm, lblData(2), txtData(2), cboUnits(2), "flow_volumetric_per_length", _
      Temp_Plant.PrimaryWeir.UnitsOfDisplay(2), "m³/m-h", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(2), txtData(2), cboUnits(2), "flow_volumetric_per_length", _
      Temp_Plant.PrimaryWeir.UnitsOfDisplay(2), "m³/(m-h)", "", "", 100#, True)
End Sub
Sub Store_Unit_Settings()
Dim i As Integer
  With Temp_Plant.PrimaryWeir
    For i = 0 To 2
      .UnitsOfDisplay(i) = unitsys_get_units(cboUnits(i))
    Next i
  End With
End Sub


Sub Populate_cbo_Model_Type()
Dim Ctl As Control
Set Ctl = cbo_Model_Type
  HALT_cbo_Model_Type = True
  Ctl.Clear
  Ctl.AddItem "Nappe": Ctl.ItemData(Ctl.NewIndex) = WEIR_MODEL_TYPE_NAPPE
  Ctl.AddItem "Pool": Ctl.ItemData(Ctl.NewIndex) = WEIR_MODEL_TYPE_POOL
  HALT_cbo_Model_Type = False
End Sub


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Private Sub cbo_Model_Type_Click()
Dim Ctl As Control
Set Ctl = cbo_Model_Type
  If (HALT_cbo_Model_Type) Then Exit Sub
  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
  Temp_Plant.PrimaryWeir.ModelingMechanism = Ctl.ItemData(Ctl.ListIndex)
  'RAISE DIRTY FLAG AND REFRESH WINDOW.
  Call Local_DirtyStatus_Set(frmD4_PrimaryWeir_Is_Dirty, True)
  Call frmD4_PrimaryWeir_Refresh(Temp_Plant)
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
  Select Case Index
    Case 0:     'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
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
  Call Local_DirtyStatus_Set(frmD4_PrimaryWeir_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  HALT_cbo_Model_Type = False
  Call Populate_cbo_Model_Type
  '
  ' POPULATE UNIT CONTROLS.
  '
  Call frmD4_PrimaryWeir_PopulateUnits
  '
  ' REFRESH DISPLAY.
  '
  Call frmD4_PrimaryWeir_Refresh(Temp_Plant)
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
      Select Case Index
        '
        ' MAIN DATA BLOCK.
        '
        Case 0: Temp_Plant.PrimaryWeir.Width = NewValue
        Case 1: Temp_Plant.PrimaryWeir.WaterLevelDiff = NewValue
        Case 2: Temp_Plant.PrimaryWeir.GasFlow = NewValue
      End Select
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD4_PrimaryWeir_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD4_PrimaryWeir_Refresh(Temp_Plant)
    End If
  End If
End Sub

