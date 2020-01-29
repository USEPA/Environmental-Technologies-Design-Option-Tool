VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEditIsothermData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing an Isotherm"
   ClientHeight    =   6630
   ClientLeft      =   6285
   ClientTop       =   1980
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   90
      TabIndex        =   12
      Top             =   3360
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13467
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   "Phase:"
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
         Height          =   255
         Index           =   2
         Left            =   5190
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   270
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Gas"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Threed.SSOption optPhase 
         Height          =   255
         Index           =   1
         Left            =   1470
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   270
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Liquid"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   765
      Left            =   90
      TabIndex        =   13
      Top             =   4170
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13467
      _ExtentY        =   1349
      _StockProps     =   14
      Caption         =   "Source:"
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
         Index           =   10
         Left            =   1440
         TabIndex        =   10
         Text            =   "txtData(10)"
         Top             =   300
         Width           =   6015
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2025
      Left            =   90
      TabIndex        =   14
      Top             =   1260
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13467
      _ExtentY        =   3572
      _StockProps     =   14
      Caption         =   "Data:"
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
         Left            =   1740
         TabIndex        =   2
         Text            =   "txtData(0)"
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   4
         Text            =   "txtData(1)"
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   6
         Text            =   "txtData(2)"
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   8
         Text            =   "txtData(3)"
         Top             =   1320
         Width           =   1635
      End
      Begin VB.TextBox txtData 
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
         Left            =   5280
         TabIndex        =   3
         Text            =   "txtData(4)"
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txtData 
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
         Left            =   5280
         TabIndex        =   5
         Text            =   "txtData(5)"
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txtData 
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
         Left            =   5280
         TabIndex        =   7
         Text            =   "txtData(6)"
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox txtData 
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
         Left            =   5280
         TabIndex        =   9
         Text            =   "txtData(7)"
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K (*)"
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
         Left            =   60
         TabIndex        =   27
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1/n"
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
         Left            =   3780
         TabIndex        =   26
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C min. (mg/L)"
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
         Left            =   60
         TabIndex        =   25
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C max. (mg/L)"
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
         Left            =   3780
         TabIndex        =   24
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "pH min."
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
         Left            =   60
         TabIndex        =   23
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "pH max."
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
         Left            =   3840
         TabIndex        =   22
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature (C)"
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
         Left            =   60
         TabIndex        =   21
         Top             =   1350
         Width           =   1575
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Carbon type"
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
         Left            =   3840
         TabIndex        =   20
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "* K in (mg/g)x(L/mg)^1/n"
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
         Left            =   600
         TabIndex        =   19
         Top             =   1710
         Width           =   3435
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   765
      Left            =   90
      TabIndex        =   15
      Top             =   5010
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13467
      _ExtentY        =   1349
      _StockProps     =   14
      Caption         =   "Comments:"
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
         Index           =   11
         Left            =   1440
         TabIndex        =   11
         Text            =   "txtData(11)"
         Top             =   300
         Width           =   6015
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   1065
      Left            =   90
      TabIndex        =   16
      Top             =   120
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13467
      _ExtentY        =   1879
      _StockProps     =   14
      Caption         =   "Chemical:"
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
         Index           =   9
         Left            =   1440
         TabIndex        =   1
         Text            =   "txtData(9)"
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox txtData 
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
         Index           =   8
         Left            =   1440
         TabIndex        =   0
         Text            =   "txtData(8)"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CAS Number"
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
         Left            =   120
         TabIndex        =   18
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   300
         TabIndex        =   17
         Top             =   630
         Width           =   1035
      End
   End
   Begin Threed.SSCommand cmdSaveCancel 
      Height          =   495
      Index           =   1
      Left            =   2730
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5970
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
      Left            =   90
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5970
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
      Left            =   6780
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5970
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
End
Attribute VB_Name = "frmEditIsothermData"
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
Dim DEFAULT_CHEMICALNAME As String
Dim DEFAULT_CHEMICALCAS As String



Const frmEditIsothermData_declarations_end = True


Sub frmEditIsothermData_AddNew( _
    INPUT_DEFAULT_PHASE_IS_LIQUID As Boolean, _
    INPUT_DEFAULT_CHEMICALNAME As String, _
    INPUT_DEFAULT_CHEMICALCAS As String, _
    OUTPUT_USER_HIT_CANCEL As Boolean, _
    OUTPUT_USER_HIT_SAVE As Boolean)
  DEFAULT_PHASE_IS_LIQUID = INPUT_DEFAULT_PHASE_IS_LIQUID
  DEFAULT_CHEMICALNAME = INPUT_DEFAULT_CHEMICALNAME
  DEFAULT_CHEMICALCAS = INPUT_DEFAULT_CHEMICALCAS
  FORM_MODE = FORM_MODE_ADDNEW
  frmEditIsothermData.Show 1
  OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
  OUTPUT_USER_HIT_SAVE = USER_HIT_SAVE
End Sub
Sub frmEditIsothermData_Edit( _
    OUTPUT_USER_HIT_CANCEL As Boolean, _
    OUTPUT_USER_HIT_SAVE As Boolean, _
    OUTPUT_USER_HIT_SAVEASNEW As Boolean)
  FORM_MODE = FORM_MODE_EDIT
  frmEditIsothermData.Show 1
  OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
  OUTPUT_USER_HIT_SAVE = USER_HIT_SAVE
  OUTPUT_USER_HIT_SAVEASNEW = USER_HIT_SAVEASNEW
End Sub



Sub frmEditIsothermData_PopulateUnits()
  Call unitsys_register(frmEditIsothermData, lblData(0), _
      txtData(0), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditIsothermData, lblData(1), _
      txtData(1), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditIsothermData, lblData(2), _
      txtData(2), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditIsothermData, lblData(3), _
      txtData(3), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditIsothermData, lblData(4), _
      txtData(4), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditIsothermData, lblData(5), _
      txtData(5), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditIsothermData, lblData(6), _
      txtData(6), Nothing, "", _
      "", "", "", "", 100#, False)
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
  Me.Height = 7035
  Me.Width = 7935
  Call CenterOnForm(Me, frmEditIsotherm)
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
    frmEditIsothermData_Record.PhaseIsLiquid = DEFAULT_PHASE_IS_LIQUID
    frmEditIsothermData_Record.Name = DEFAULT_CHEMICALNAME
    frmEditIsothermData_Record.CAS = DEFAULT_CHEMICALCAS
    frmEditIsothermData_Record.k = 1#
    frmEditIsothermData_Record.OneOverN = 1#
    frmEditIsothermData_Record.Cmin = 0#
    frmEditIsothermData_Record.Cmax = 0#
    frmEditIsothermData_Record.pHmin = 0#
    frmEditIsothermData_Record.pHmax = 0#
    frmEditIsothermData_Record.Source = "Type Source Here"
    frmEditIsothermData_Record.CarbonName = "Type Carbon Here"
    frmEditIsothermData_Record.Tmin = 25#
    frmEditIsothermData_Record.Comments = ""
  End If
  If (frmEditIsothermData_Record.PhaseIsLiquid) Then
    optPhase(1).Value = True
    optPhase(2).Value = False
  Else
    optPhase(1).Value = False
    optPhase(2).Value = True
  End If
  'POPULATE UNIT CONTROLS.
  Call frmEditIsothermData_PopulateUnits
  'REFRESH DISPLAY.
  Call frmEditIsothermData_Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub









Private Sub optPhase_Click(Index As Integer, Value As Integer)
  If (Index = 1) Then
    frmEditIsothermData_Record.PhaseIsLiquid = True
  Else
    frmEditIsothermData_Record.PhaseIsLiquid = False
  End If
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
  If (Index >= 7) And (Index <= 11) Then
    Call Global_GotFocus(Ctl)
  Else
    Call unitsys_control_txtx_gotfocus(Ctl)
  End If
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  If (Index >= 7) And (Index <= 11) Then
    If (Index = 8) Then
      KeyAscii = Global_Numeric0123456789KeyPress(KeyAscii)
    Else
      KeyAscii = Global_TextKeyPress(KeyAscii)
    End If
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
  If (Index >= 7) And (Index <= 11) Then
    Select Case Index
      Case 7: OldValueStr = Trim$(frmEditIsothermData_Record.CarbonName)
      Case 8: OldValueStr = Trim$(frmEditIsothermData_Record.CAS)
      Case 9: OldValueStr = Trim$(frmEditIsothermData_Record.Name)
      Case 10: OldValueStr = Trim$(frmEditIsothermData_Record.Source)
      Case 11: OldValueStr = Trim$(frmEditIsothermData_Record.Comments)
    End Select
    'NOTE: ZERO-LENGTH STRINGS FOR 8 AND 11 ARE ALLOWED.
    If (Trim$(Ctl.Text) = "") And _
        (Index <> 8) And (Index <> 11) Then
      Ctl.Text = OldValueStr
      'Call Show_Error("You must enter a non-blank string for the carbon name.")
      'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
      'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
    Else
      If (Trim$(OldValueStr) <> Trim$(Ctl.Text)) Then
        Select Case Index
          Case 7: frmEditIsothermData_Record.CarbonName = Trim$(Ctl.Text)
          Case 8: frmEditIsothermData_Record.CAS = Trim$(Ctl.Text)
          Case 9: frmEditIsothermData_Record.Name = Trim$(Ctl.Text)
          Case 10: frmEditIsothermData_Record.Source = Trim$(Ctl.Text)
          Case 11: frmEditIsothermData_Record.Comments = Trim$(Ctl.Text)
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
    Case 1: Val_Low = 0#: Val_High = 1E+20
    Case 2: Val_Low = 0#: Val_High = 1E+20
    Case 3: Val_Low = 1E-20: Val_High = 1E+20
    Case 4: Val_Low = 1E-20: Val_High = 1E+20
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
        Case 0:         'FREUNDLICH K.
          frmEditIsothermData_Record.k = NewValue
        Case 1:         'MINIMUM CONCENTRATION.
          frmEditIsothermData_Record.Cmin = NewValue
        Case 2:         'MINIMUM pH.
          frmEditIsothermData_Record.pHmin = NewValue
        Case 3:         'TEMPERATURE.
          frmEditIsothermData_Record.Tmin = NewValue
        Case 4:         'FREUNDLICH 1/n.
          frmEditIsothermData_Record.OneOverN = NewValue
        Case 5:         'MAXIMUM CONCENTRATION.
          frmEditIsothermData_Record.Cmax = NewValue
        Case 6:         'MAXIMUM pH.
          frmEditIsothermData_Record.pHmax = NewValue
      End Select
      'RAISE DIRTY FLAG IF NECESSARY.
      If (Raise_Dirty_Flag) Then
        ''THROW DIRTY FLAG.
        'Call frmCompoProp_DirtyStatus_Throw
      End If
      'REFRESH WINDOW.
      Call frmEditIsothermData_Refresh
    End If
  End If
End Sub






