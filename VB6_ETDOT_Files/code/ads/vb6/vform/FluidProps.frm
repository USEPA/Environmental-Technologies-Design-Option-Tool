VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFluidProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "{Fluid} Properties"
   ClientHeight    =   2730
   ClientLeft      =   4230
   ClientTop       =   3480
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5085
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   4920
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   11
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   1680
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1395
      Left            =   240
      TabIndex        =   4
      Top             =   90
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   2461
      _StockProps     =   14
      Caption         =   "Density and Viscosity"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSFrame SSFrame2 
         Height          =   1005
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   1773
         _StockProps     =   14
         Caption         =   "Use correlation for:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCheck chkCorr 
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   0
            Top             =   270
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "&Density"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkCorr 
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   1
            Top             =   630
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "&Viscosity"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   1005
         Left            =   2010
         TabIndex        =   6
         Top             =   270
         Width           =   2445
         _Version        =   65536
         _ExtentX        =   4313
         _ExtentY        =   1773
         _StockProps     =   14
         Caption         =   "Values:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtWater 
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
            Height          =   288
            Index           =   0
            Left            =   90
            TabIndex        =   2
            Text            =   "txtWater(0)"
            Top             =   270
            Width           =   1365
         End
         Begin VB.TextBox txtWater 
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
            Height          =   288
            Index           =   1
            Left            =   90
            TabIndex        =   3
            Text            =   "txtWater(1)"
            Top             =   630
            Width           =   1365
         End
         Begin VB.Label lblUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "lblUnit(0)"
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
            Left            =   1620
            TabIndex        =   8
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "lblUnit(1)"
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
            Left            =   1620
            TabIndex        =   7
            Top             =   660
            Width           =   735
         End
      End
   End
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
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
      Left            =   2400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
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
Attribute VB_Name = "frmFluidProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_OK As Boolean
Dim USER_HIT_CANCEL As Boolean

Dim Save_Density As Double
Dim Save_Viscosity As Double
Dim Save_State_Check_Water(1 To 2) As Integer




Const frmFluidProps_declarations_end = True


Sub frmFluidProps_Edit( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  frmFluidProps.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub
Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdCancelOK(1).Enabled = False
  End If
End Sub


Sub frmFluidProps_PopulateUnits()
  Call unitsys_register(frmFluidProps, lblUnit(0), _
      txtWater(0), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmFluidProps, lblUnit(1), _
      txtWater(1), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Private Sub chkCorr_Click(Index As Integer, Value As Integer)
  'UPDATE MEMORY.
  State_Check_Water(Index + 1) = Value
  'IF TURNED CORRELATION ON, RE-CALCULATE DENSITY/VISCOSITY.
  If (Value) Then
    Select Case Index
      Case 0:     'DENSITY.
        Call Update_FluidDensity(Bed.Temperature, Bed.Pressure, Bed.WaterDensity)
      Case 1:     'VISCOSITY.
        Call Update_FluidViscosity(Bed.Temperature, Bed.WaterViscosity)
    End Select
  End If
  'REFRESH DISPLAY.
  Call frmFluidProps_Refresh
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
  Select Case Index
    Case 0:     'CANCEL.
      'ROLLBACK TO ORIGINAL VALUES.
      
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
  'MISC INITS.
  Me.Height = 3210
  Me.Width = 5235
  Call CenterOnForm(Me, frmMain)
  If (Bed.Phase = 0) Then
    Me.Caption = "Water Properties"
  Else
    Me.Caption = "Air Properties"
  End If
  lblUnit(0).Caption = "g/cm" & Chr$(179)
  lblUnit(1).Caption = "g/cm-s"

'txtWater(2) = Format$(Bed.WaterDensity, "0.000E+00")
'txtWater(3) = Format$(Bed.WaterViscosity, "0.00E+00")
'State_Check_Water(1) = chkCorr(0).Value
'State_Check_Water(2) = chkCorr(1).Value
  
  'SAVE OLD VALUES FOR CANCEL ROLLBACK.
  Save_Density = Bed.WaterDensity
  Save_Viscosity = Bed.WaterViscosity
  Save_State_Check_Water(1) = State_Check_Water(1)
  Save_State_Check_Water(2) = State_Check_Water(2)
  'POPULATE UNIT CONTROLS.
  Call frmFluidProps_PopulateUnits
  'REFRESH DISPLAY.
  Call frmFluidProps_Refresh
  'DEMO SETTINGS.
  Call LOCAL___Reset_DemoVersionDisablings
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub txtWater_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtWater(Index)
  Call unitsys_control_txtx_gotfocus(Ctl)
End Sub
Private Sub txtWater_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtWater_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtWater(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
  Select Case Index
    Case 0: Val_Low = 1E-20: Val_High = 1E+20
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
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
        Case 0:         'DENSITY.
          Bed.WaterDensity = NewValue
        Case 1:         'VISCOSITY.
          Bed.WaterViscosity = NewValue
      End Select
      'RAISE DIRTY FLAG IF NECESSARY.
      If (Raise_Dirty_Flag) Then
        ''THROW DIRTY FLAG.
        'Call frmCompoProp_DirtyStatus_Throw
      End If
      'REFRESH WINDOW.
      Call frmFluidProps_Refresh
    End If
  End If
End Sub

