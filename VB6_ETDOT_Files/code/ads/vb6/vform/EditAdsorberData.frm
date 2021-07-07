VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEditAdsorberData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing an Adsorber"
   ClientHeight    =   5025
   ClientLeft      =   1275
   ClientTop       =   2550
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5085
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   5040
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   795
      Left            =   180
      TabIndex        =   8
      Top             =   90
      Width           =   4755
      _Version        =   65536
      _ExtentX        =   8387
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
         TabIndex        =   10
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
      End
      Begin Threed.SSOption optPhase 
         Height          =   375
         Index           =   2
         Left            =   2610
         TabIndex        =   11
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
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3135
      Left            =   180
      TabIndex        =   9
      Top             =   1020
      Width           =   4755
      _Version        =   65536
      _ExtentX        =   8387
      _ExtentY        =   5530
      _StockProps     =   14
      Caption         =   "Adsorber Properties:"
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
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Text            =   "txtData(1)"
         Top             =   900
         Width           =   1395
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
         Left            =   2160
         TabIndex        =   3
         Text            =   "txtData(2)"
         Top             =   1200
         Width           =   1395
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
         Left            =   2160
         TabIndex        =   4
         Text            =   "txtData(3)"
         Top             =   1500
         Width           =   1395
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
         Left            =   2160
         TabIndex        =   5
         Text            =   "txtData(4)"
         Top             =   1800
         Width           =   1395
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
         Left            =   2160
         TabIndex        =   6
         Text            =   "txtData(5)"
         Top             =   2100
         Width           =   1395
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
         Left            =   600
         TabIndex        =   7
         Text            =   "txtData(6)"
         Top             =   2670
         Width           =   3915
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
         Left            =   2160
         TabIndex        =   1
         Text            =   "txtData(0)"
         Top             =   600
         Width           =   1395
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
         Index           =   7
         Left            =   2160
         TabIndex        =   0
         Text            =   "txtData(7)"
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label lblDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Notes:"
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
         Left            =   600
         TabIndex        =   25
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(see code)"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   2130
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Default Flow Rate:"
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
         Left            =   60
         TabIndex        =   23
         Top             =   2130
         Width           =   2055
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(see code)"
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
         Left            =   3600
         TabIndex        =   22
         Top             =   1830
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Design Flow Range:"
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
         Left            =   60
         TabIndex        =   21
         Top             =   1830
         Width           =   2055
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(psig)"
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
         Left            =   3600
         TabIndex        =   20
         Top             =   1530
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Design Pressure:"
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
         TabIndex        =   19
         Top             =   1530
         Width           =   2055
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(feet)"
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
         Left            =   3600
         TabIndex        =   18
         Top             =   1230
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Outside Diameter:"
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
         TabIndex        =   17
         Top             =   1230
         Width           =   2055
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(lbs)"
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
         Left            =   3600
         TabIndex        =   16
         Top             =   930
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Maximum Capacity:"
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
         TabIndex        =   15
         Top             =   930
         Width           =   2055
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(ft²)"
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
         Left            =   3600
         TabIndex        =   14
         Top             =   630
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Internal Area:"
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
         TabIndex        =   13
         Top             =   630
         Width           =   2055
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Part Number:"
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
         Left            =   60
         TabIndex        =   12
         Top             =   330
         Width           =   2055
      End
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   495
      Left            =   480
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4410
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
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
   Begin Threed.SSCommand cmdCancel 
      Height          =   495
      Left            =   3240
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4410
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
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
Attribute VB_Name = "frmEditAdsorberData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim frmEditAdsorber_Cancelled As Integer
'Dim frmEditAdsorber_RunMode As Integer
'Const frmEditAdsorber_RunMode_QUERY_DATABASE = 1
'Const frmEditAdsorber_RunMode_EDIT_DATABASE = 2

'Dim frmEditAdsorberData_Cancelled As Integer
Dim frmEditAdsorberData_RunMode As Integer
Const frmEditAdsorberData_RunMode_NEW = 1
Const frmEditAdsorberData_RunMode_EDIT = 2
Dim frmEditAdsorberData_UsePhase As Integer

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_SAVE As Boolean




Const frmEditAdsorberData_declarations_end = True


Sub frmEditAdsorberData_AddNew( _
    INPUT_PHASE As Integer, _
    OUTPUT_USER_HIT_CANCEL As Boolean)
  frmEditAdsorberData_RunMode = frmEditAdsorberData_RunMode_NEW
  frmEditAdsorberData_UsePhase = INPUT_PHASE
  frmEditAdsorberData.Show 1
  If (USER_HIT_CANCEL) Then
    OUTPUT_USER_HIT_CANCEL = True
  Else
    OUTPUT_USER_HIT_CANCEL = False
  End If
End Sub
Sub frmEditAdsorberData_Edit( _
    INPUT_PHASE As Integer, _
    OUTPUT_USER_HIT_CANCEL As Boolean)
  frmEditAdsorberData_RunMode = frmEditAdsorberData_RunMode_EDIT
  frmEditAdsorberData_UsePhase = INPUT_PHASE
  frmEditAdsorberData.Show 1
  If (USER_HIT_CANCEL) Then
    OUTPUT_USER_HIT_CANCEL = True
  Else
    OUTPUT_USER_HIT_CANCEL = False
  End If
End Sub


Sub frmEditAdsorberData_PopulateUnits()
  Call unitsys_register(frmEditAdsorberData, lblDesc(0), _
      txtData(0), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditAdsorberData, lblDesc(1), _
      txtData(1), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditAdsorberData, lblDesc(2), _
      txtData(2), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmEditAdsorberData, lblDesc(5), _
      txtData(5), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Private Sub cmdCancel_Click()
  'frmEditAdsorberData_Cancelled = True
  USER_HIT_CANCEL = True
  USER_HIT_SAVE = False
  Unload Me
End Sub
Private Sub cmdSave_Click()
Dim i As Integer
  For i = 0 To 7
    If (Trim$(txtData(i)) = "") Then
      Beep
      MsgBox "No data item can be set to an empty string; enter a non-empty string or hit Cancel.", vbExclamation, AppName_For_Display_Long
      Exit Sub
    End If
  Next i
  If (optPhase(1).Value) Then
    frmEditAdsorberData_Record.Phase = 1
  Else
    frmEditAdsorberData_Record.Phase = 2
  End If
  frmEditAdsorberData_Record.InternalArea = txtData(0)
  frmEditAdsorberData_Record.MaxCapacity = txtData(1)
  frmEditAdsorberData_Record.OutsideDiameter = txtData(2)
  frmEditAdsorberData_Record.DesignPressure = txtData(3)
  frmEditAdsorberData_Record.DesignFlowRange = txtData(4)
  frmEditAdsorberData_Record.DefaultFlowRate = txtData(5)
  frmEditAdsorberData_Record.Note = txtData(6)
  frmEditAdsorberData_Record.PartNumber = txtData(7)
  'frmEditAdsorberData_Cancelled = False
  USER_HIT_CANCEL = False
  USER_HIT_SAVE = True
  Unload Me
End Sub


Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
Dim now_phase As Integer
Dim i As Integer
  'MISC INITS.
  Me.Height = 5505
  Me.Width = 5205
  Call CenterOnForm(Me, frmEditAdsorber)
  now_phase = frmEditAdsorberData_UsePhase
  If (now_phase = 1) Then
    optPhase(1).Value = True
    optPhase(2).Value = False
  Else
    optPhase(1).Value = False
    optPhase(2).Value = True
  End If
  If (frmEditAdsorberData_RunMode = frmEditAdsorberData_RunMode_NEW) Then
    'CREATE NEW RECORD
    frmEditAdsorberData_Record.PartNumber = "New Adsorber Code"
    frmEditAdsorberData_Record.InternalArea = "1"
    frmEditAdsorberData_Record.MaxCapacity = "1000"
    frmEditAdsorberData_Record.OutsideDiameter = "1"
    frmEditAdsorberData_Record.DesignPressure = "not available"
    frmEditAdsorberData_Record.DesignFlowRange = "1-10"
    frmEditAdsorberData_Record.DefaultFlowRate = "10"
    frmEditAdsorberData_Record.Note = "none"
  End If
  If (frmEditAdsorberData_RunMode = frmEditAdsorberData_RunMode_EDIT) Then
    'MODIFY EXISTING RECORD
'    txtData(0) = Trim$(frmEditAdsorberData_Record.InternalArea)
'    txtData(1) = Trim$(frmEditAdsorberData_Record.MaxCapacity)
'    txtData(2) = Trim$(frmEditAdsorberData_Record.OutsideDiameter)
'    txtData(3) = Trim$(frmEditAdsorberData_Record.DesignPressure)
'    txtData(4) = Trim$(frmEditAdsorberData_Record.DesignFlowRange)
'    txtData(5) = Trim$(frmEditAdsorberData_Record.DefaultFlowRate)
'    txtData(6) = Trim$(frmEditAdsorberData_Record.Note)
'    txtData(7) = Trim$(frmEditAdsorberData_Record.PartNumber)
  End If
  'POPULATE UNIT CONTROLS.
  Call frmEditAdsorberData_PopulateUnits
  'REFRESH DISPLAY.
  Call frmEditAdsorberData_Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub optPhase_Click(Index As Integer, Value As Integer)
  If (optPhase(1).Value) Then
    'LIQUID PHASE
    lblUnits(4).Caption = "(gal/min)"
    lblUnits(5).Caption = "(gal/min)"
  Else
    'GAS PHASE
    lblUnits(4).Caption = "(ft³/min)"
    lblUnits(5).Caption = "(ft³/min)"
  End If
End Sub






Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
  If (Index = 7) Or (Index = 3) Or _
      (Index = 4) Or (Index = 6) Then
    Call Global_GotFocus(Ctl)
  Else
    Call unitsys_control_txtx_gotfocus(Ctl)
  End If
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  If (Index = 7) Or (Index = 3) Or _
      (Index = 4) Or (Index = 6) Then
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
  If (Index = 7) Or (Index = 3) Or _
      (Index = 4) Or (Index = 6) Then
    Select Case Index
      Case 7: OldValueStr = Trim$(frmEditAdsorberData_Record.PartNumber)
      Case 3: OldValueStr = Trim$(frmEditAdsorberData_Record.DesignPressure)
      Case 4: OldValueStr = Trim$(frmEditAdsorberData_Record.DesignFlowRange)
      Case 6: OldValueStr = Trim$(frmEditAdsorberData_Record.Note)
    End Select
    If (Trim$(Ctl.Text) = "") Then
      Ctl.Text = OldValueStr
      'Call Show_Error("You must enter a non-blank string for the carbon name.")
      'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
      'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
    Else
      If (Trim$(OldValueStr) <> Trim$(Ctl.Text)) Then
        Select Case Index
          Case 7: frmEditAdsorberData_Record.PartNumber = Trim$(Ctl.Text)
          Case 3: frmEditAdsorberData_Record.DesignPressure = Trim$(Ctl.Text)
          Case 4: frmEditAdsorberData_Record.DesignFlowRange = Trim$(Ctl.Text)
          Case 6: frmEditAdsorberData_Record.Note = Trim$(Ctl.Text)
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
    Case 5: Val_Low = 1E-20: Val_High = 1E+20
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
        Case 0:         'INTERNAL AREA.
          frmEditAdsorberData_Record.InternalArea = Trim$(Str$(NewValue))
        Case 1:         'MAXIMUM CAPACITY.
          frmEditAdsorberData_Record.MaxCapacity = Trim$(Str$(NewValue))
        Case 2:         'OUTSIDE DIAMETER.
          frmEditAdsorberData_Record.OutsideDiameter = Trim$(Str$(NewValue))
        Case 5:         'DEFAULT FLOW RATE.
          frmEditAdsorberData_Record.DefaultFlowRate = Trim$(Str$(NewValue))
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



