VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmEditCarbonData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing an Adsorbent"
   ClientHeight    =   5640
   ClientLeft      =   1095
   ClientTop       =   2280
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   4860
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
      _ExtentY        =   6641
      _StockProps     =   14
      Caption         =   "Adsorbent Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   2010
         TabIndex        =   0
         Text            =   "txtData(7)"
         Top             =   285
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   2010
         TabIndex        =   7
         Text            =   "txtData(6)"
         Top             =   2805
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   2010
         TabIndex        =   6
         Text            =   "txtData(5)"
         Top             =   2445
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2010
         TabIndex        =   5
         Text            =   "txtData(4)"
         Top             =   2085
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2010
         TabIndex        =   4
         Text            =   "txtData(3)"
         Top             =   1725
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2010
         TabIndex        =   3
         Text            =   "txtData(2)"
         Top             =   1365
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2010
         TabIndex        =   2
         Text            =   "txtData(1)"
         Top             =   1005
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2010
         TabIndex        =   1
         Text            =   "txtData(0)"
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   26
         Top             =   315
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apparent Density"
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
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   675
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Radius"
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
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   1035
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Porosity"
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
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   20
         Top             =   1395
         Width           =   1755
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "g/cm3"
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
         Height          =   255
         Index           =   0
         Left            =   3210
         TabIndex        =   19
         Top             =   675
         Width           =   975
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
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
         Height          =   255
         Index           =   1
         Left            =   3210
         TabIndex        =   18
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Height          =   255
         Index           =   2
         Left            =   3210
         TabIndex        =   17
         Top             =   1395
         Width           =   975
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Adsorbent Type"
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
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   16
         Top             =   1755
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi W0"
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
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   15
         Top             =   2115
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi B"
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
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   14
         Top             =   2475
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi Exponent"
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
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   13
         Top             =   2835
         Width           =   1755
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm3/g"
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
         Height          =   255
         Index           =   4
         Left            =   3210
         TabIndex        =   12
         Top             =   2115
         Width           =   975
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Height          =   255
         Index           =   5
         Left            =   3210
         TabIndex        =   11
         Top             =   2475
         Width           =   975
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Height          =   255
         Index           =   6
         Left            =   3210
         TabIndex        =   10
         Top             =   2835
         Width           =   975
      End
      Begin VB.Label lblUnitB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "* in (mol/cal) ^(Polanyi Exponent)"
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
         Height          =   255
         Left            =   1230
         TabIndex        =   9
         Top             =   3315
         Width           =   2925
      End
   End
   Begin Threed.SSCommand cmdSaveCancel 
      Height          =   495
      Index           =   1
      Left            =   1260
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4980
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Save &As New Record"
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
   Begin Threed.SSCommand cmdSaveCancel 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4980
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Save"
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
   Begin Threed.SSCommand cmdSaveCancel 
      Height          =   495
      Index           =   2
      Left            =   3780
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4980
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   795
      Left            =   120
      TabIndex        =   27
      Top             =   90
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
      _ExtentY        =   1402
      _StockProps     =   14
      Caption         =   "Select Phase:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optPhase 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   300
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Liquid Phase"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Threed.SSOption optPhase 
         Height          =   375
         Index           =   2
         Left            =   2610
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   300
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Gas Phase"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmEditCarbonData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FORM_MODE As Integer
Const FORM_MODE_ADDNEW = 1
Const FORM_MODE_EDIT = 2

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_SAVE As Boolean
Dim USER_HIT_SAVEASNEW As Boolean

Dim DEFAULT_PHASE_IS_LIQUID As Boolean


Const frmEditCarbonData_declarations_end = True


Sub frmEditCarbonData_AddNew( _
    INPUT_DEFAULT_PHASE_IS_LIQUID As Boolean, _
    OUTPUT_USER_HIT_CANCEL As Boolean, _
    OUTPUT_USER_HIT_SAVE As Boolean)
  DEFAULT_PHASE_IS_LIQUID = INPUT_DEFAULT_PHASE_IS_LIQUID
  FORM_MODE = FORM_MODE_ADDNEW
  frmEditCarbonData.Show 1
  OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
  OUTPUT_USER_HIT_SAVE = USER_HIT_SAVE
End Sub
Sub frmEditCarbonData_Edit( _
    OUTPUT_USER_HIT_CANCEL As Boolean, _
    OUTPUT_USER_HIT_SAVE As Boolean, _
    OUTPUT_USER_HIT_SAVEASNEW As Boolean)
  FORM_MODE = FORM_MODE_EDIT
  frmEditCarbonData.Show 1
  OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
  OUTPUT_USER_HIT_SAVE = USER_HIT_SAVE
  OUTPUT_USER_HIT_SAVEASNEW = USER_HIT_SAVEASNEW
End Sub



Sub frmEditCarbonData_PopulateUnits()
  Call unitsys_register(frmEditCarbonData, lblDesc(0), _
      txtData(0), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditCarbonData, lblDesc(1), _
      txtData(1), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditCarbonData, lblDesc(2), _
      txtData(2), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditCarbonData, lblDesc(4), _
      txtData(4), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditCarbonData, lblDesc(5), _
      txtData(5), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditCarbonData, lblDesc(6), _
      txtData(6), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Private Sub cmdSaveAs_Click()

End Sub


Private Sub cmdSaveCancel_Click(Index As Integer)
  Select Case Index
    Case 0:       'SAVE.
      USER_HIT_CANCEL = False
      USER_HIT_SAVE = True
      USER_HIT_SAVEASNEW = False
      Unload Me
      Exit Sub
    Case 1:       'SAVE AS NEW RECORD.
      USER_HIT_CANCEL = False
      USER_HIT_SAVE = False
      USER_HIT_SAVEASNEW = True
      Unload Me
      Exit Sub
    Case 2:       'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_SAVE = False
      USER_HIT_SAVEASNEW = False
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub Form_Load()
  'MISC INITS.
  Me.Height = 6045
  Me.Width = 4995
  Call CenterOnForm(Me, frmEditAdsorber)
  lblUnit(0).Caption = "g/cm³"
  lblUnit(4).Caption = "cm³/g"
  'STRANGE THINGS CAN HAPPEN IF OPTION BOXES ARE ENABLED
  'BEFORE THE FORM IS LOADED/ACTIVATED.
  optPhase(1).Enabled = True
  optPhase(2).Enabled = True
  If (FORM_MODE = FORM_MODE_EDIT) Then
    'EDIT MODE.
    cmdSaveCancel(0).Visible = True
    cmdSaveCancel(1).Visible = True
    cmdSaveCancel(2).Visible = True
  Else
    'ADD NEW MODE.
    cmdSaveCancel(0).Visible = True
    cmdSaveCancel(1).Visible = False
    cmdSaveCancel(2).Visible = True
    'SET DEFAULTS FOR THE NEW RECORD.
    frmEditCarbonData_Record.Name = "New Adsorbent"
    frmEditCarbonData_Record.AppDen = 1#
    frmEditCarbonData_Record.ParticleRadius = 0.1
    frmEditCarbonData_Record.ParticlePorosity = 1#
    frmEditCarbonData_Record.AdsType = "GAC"
    frmEditCarbonData_Record.W0 = 0#
    frmEditCarbonData_Record.BB = 0#
    frmEditCarbonData_Record.PolanyiExponent = 0#
    If (DEFAULT_PHASE_IS_LIQUID) Then
      frmEditCarbonData_Record.PhaseIsLiquid = True
    Else
      frmEditCarbonData_Record.PhaseIsLiquid = False
    End If
  End If
  If (frmEditCarbonData_Record.PhaseIsLiquid) Then
    optPhase(1).Value = True
    optPhase(2).Value = False
  Else
    optPhase(1).Value = False
    optPhase(2).Value = True
  End If
  'POPULATE UNIT CONTROLS.
  Call frmEditCarbonData_PopulateUnits
  'REFRESH DISPLAY.
  Call frmEditCarbonData_Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub









Private Sub optPhase_Click(Index As Integer, Value As Integer)
  If (Index = 1) Then
    frmEditCarbonData_Record.PhaseIsLiquid = True
  Else
    frmEditCarbonData_Record.PhaseIsLiquid = False
  End If
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
  If (Index = 7) Or (Index = 3) Then
    Call Global_GotFocus(Ctl)
  Else
    Call unitsys_control_txtx_gotfocus(Ctl)
  End If
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  If (Index = 7) Or (Index = 3) Then
    KeyAscii = Global_TextKeyPress(KeyAscii)
  Else
    KeyAscii = Global_NumericKeyPress(KeyAscii)
  End If
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
Dim OldValueStr As String
  'HANDLE STRING FIELDS.
  If (Index = 7) Or (Index = 3) Then
    Select Case Index
      Case 7: OldValueStr = Trim$(frmEditCarbonData_Record.Name)
      Case 3: OldValueStr = Trim$(frmEditCarbonData_Record.AdsType)
    End Select
    If (Trim$(Ctl.Text) = "") Then
      Ctl.Text = OldValueStr
      'Call Show_Error("You must enter a non-blank string for the carbon name.")
      'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
      'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
    Else
      If (Trim$(OldValueStr) <> Trim$(Ctl.Text)) Then
        Select Case Index
          Case 7: frmEditCarbonData_Record.Name = Trim$(Ctl.Text)
          Case 3: frmEditCarbonData_Record.AdsType = Trim$(Ctl.Text)
        End Select
        ''THROW DIRTY FLAG.
        'Call DirtyStatus_Throw
      End If
    End If
    Call Global_LostFocus(Ctl)
    'Call GenericStatus_Set("")
    Exit Sub
  End If
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
  Select Case Index
    Case 0: Val_Low = 1E-20: Val_High = 1E+20
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
    Case 2: Val_Low = 1E-20: Val_High = 1E+20
    Case 4: Val_Low = 0#: Val_High = 1E+20
    Case 5: Val_Low = 0#: Val_High = 1E+20
    Case 6: Val_Low = 0#: Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 0:         'APPARENT DENSITY.
          frmEditCarbonData_Record.AppDen = NewValue
        Case 1:         'PARTICLE RADIUS.
          frmEditCarbonData_Record.ParticleRadius = NewValue
        Case 2:         'PARTICLE POROSITY.
          frmEditCarbonData_Record.ParticlePorosity = NewValue
        Case 4:         'POLANYI W0.
          frmEditCarbonData_Record.W0 = NewValue
        Case 5:         'POLANYI BB.
          frmEditCarbonData_Record.BB = NewValue
        Case 6:         'POLANYI EXPONENT.
          frmEditCarbonData_Record.PolanyiExponent = NewValue
      End Select
      'RAISE DIRTY FLAG IF NECESSARY.
      If (Raise_Dirty_Flag) Then
        ''THROW DIRTY FLAG.
        'Call frmCompoProp_DirtyStatus_Throw
      End If
      'REFRESH WINDOW.
      Call frmEditAdsorberData_Refresh
    End If
  End If
End Sub




