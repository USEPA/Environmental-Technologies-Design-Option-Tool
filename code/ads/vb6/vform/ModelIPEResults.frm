VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmModelIPEResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Isotherm Parameter Estimation (IPE) Results"
   ClientHeight    =   6660
   ClientLeft      =   4530
   ClientTop       =   2160
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6705
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   6600
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   30
      Top             =   4680
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
      Left            =   960
      TabIndex        =   29
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   6105
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2235
      Left            =   510
      TabIndex        =   0
      Top             =   90
      Width           =   5685
      _Version        =   65536
      _ExtentX        =   10028
      _ExtentY        =   3942
      _StockProps     =   14
      Caption         =   "Input Data:"
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
         Height          =   555
         Index           =   3
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "ModelIPEResults.frx":0000
         Top             =   1530
         Width           =   5265
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
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "txtData(2)"
         Top             =   930
         Width           =   1600
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
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "txtData(1)"
         Top             =   600
         Width           =   1600
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
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "txtData(0)"
         Top             =   270
         Width           =   1600
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Model Used:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   9
         Top             =   1290
         Width           =   1485
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi GM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   7
         Top             =   960
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi BB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   630
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   300
         Width           =   3600
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3525
      Left            =   510
      TabIndex        =   1
      Top             =   2430
      Width           =   5685
      _Version        =   65536
      _ExtentX        =   10028
      _ExtentY        =   6218
      _StockProps     =   14
      Caption         =   "Results:"
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
         Index           =   12
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "txtData(12)"
         Top             =   3090
         Width           =   1600
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
         Index           =   11
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "txtData(11)"
         Top             =   2760
         Width           =   1600
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
         Index           =   10
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "txtData(10)"
         Top             =   2430
         Width           =   1600
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
         Index           =   9
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "txtData(9)"
         Top             =   2100
         Width           =   1600
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
         Index           =   8
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "txtData(8)"
         Top             =   1620
         Width           =   1600
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
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "txtData(7)"
         Top             =   1290
         Width           =   1600
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
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "txtData(6)"
         Top             =   960
         Width           =   1600
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
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "txtData(5)"
         Top             =   630
         Width           =   1600
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
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "txtData(4)"
         Top             =   300
         Width           =   1600
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   5640
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Root Mean Square Error"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   100
         TabIndex        =   27
         Top             =   3120
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Regression R Squared"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   100
         TabIndex        =   25
         Top             =   2790
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Upper Correlation Limit (mg/L)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   100
         TabIndex        =   23
         Top             =   2460
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lower Correlation Limit (mg/L)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   100
         TabIndex        =   21
         Top             =   2130
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Freundlich 1/n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   100
         TabIndex        =   19
         Top             =   1650
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Freundlich K (mmol/g)*(L/mmol)^(1/n)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   100
         TabIndex        =   17
         Top             =   1320
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Freundlich K (mg/g)*(L/mg)^(1/n)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   100
         TabIndex        =   15
         Top             =   990
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi Adsorption Capacity (mg/g)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   100
         TabIndex        =   13
         Top             =   660
         Width           =   3600
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Average Chemical Conc. (mg/L)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   11
         Top             =   330
         Width           =   3600
      End
   End
   Begin Threed.SSCommand cmdClose 
      Height          =   495
      Left            =   3900
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6060
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Close"
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
Attribute VB_Name = "frmModelIPEResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WHICH_MODEL As Integer




Const frmModelIPEResults_declarations_end = True


Sub frmModelIPEResults_Run( _
    INPUT_WHICH_MODEL As Integer)
  WHICH_MODEL = INPUT_WHICH_MODEL
  frmModelIPEResults.Show 1
End Sub


Sub frmModelIPEResults_PopulateUnits()
Dim i As Integer
  For i = 0 To 12
    If (i <> 3) Then
      Call unitsys_register(frmModelIPEResults, lblDesc(i), _
          txtData(i), Nothing, "", _
          "", "", "", "", 100#, False)
    End If
  Next i
End Sub
Sub frmModelIPEResults_Refresh()
Dim Frm As Form
Set Frm = frmModelIPEResults
Dim i As Integer
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(0), IPES_Data.Input.W0)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(1), IPES_Data.Input.BB)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(2), IPES_Data.Input.GM)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(4), IPES_Data.Output.CSAV)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(5), IPES_Data.Output.QSAV)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(6), IPES_Data.Output.XK1)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(7), IPES_Data.Output.XK2)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(8), IPES_Data.Output.XN)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(9), IPES_Data.Output.CBEG)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(10), IPES_Data.Output.CEND)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(11), IPES_Data.Output.RSQD)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(12), IPES_Data.Output.RMSE)
  '
  ' REMOVE 9-12 IF NO DATA AVAILABLE.
  '
  If ((IPES_Data.Output.CBEG = 0#) And _
      (IPES_Data.Output.CEND = 0#) And _
      (IPES_Data.Output.RSQD = 0#) And _
      (IPES_Data.Output.RMSE = 0#)) Then
    For i = 9 To 12
      Frm.lblDesc(i).Visible = False
      Frm.txtData(i).Visible = False
    Next i
  End If
End Sub


Private Sub cmdClose_Click()
  Unload Me
  Exit Sub
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
  Call CenterOnForm(Me, frmFreundlich)
  Select Case WHICH_MODEL
    Case MODULECODE_ADLIQ:
      txtData(3).Text = "3-Parameter Polanyi Correlation"
    Case MODULECODE_SPEQ:
      txtData(3).Text = "D-R Equal Spreading Pressure Calculation"
    Case MODULECODE_HOFMAN:
      txtData(3).Text = "Estimated From Gas-Phase D-R Isotherm" & _
          vbCrLf & "Hansen-Fackler model, uniform adsorbate"
    Case Else:
  End Select
  'POPULATE UNIT CONTROLS.
  Call frmModelIPEResults_PopulateUnits
  'REFRESH DISPLAY.
  Call frmModelIPEResults_Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
  If (Index = 3) Then
    Call Global_GotFocus(Ctl)
    Exit Sub
  End If
  Call unitsys_control_txtx_gotfocus(Ctl)
End Sub
'Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
'  KeyAscii = Global_NumericKeyPress(KeyAscii)
'End Sub
Private Sub txtData_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
'Dim Too_Small As Integer
  If (Index = 3) Then
    Call Global_LostFocus(Ctl)
    Exit Sub
  End If
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
  Val_Low = -1E+20
  Val_High = 1E+20
  Select Case Index
    Case 0: Val_Low = 0.05: Val_High = 2.5
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
    Case 2: Val_Low = 1E-20: Val_High = 1E+20
  End Select
  'NOTE: THE VALUES SHOULD NEVER CHANGE BECAUSE ALL TEXT BOXES
  'ON THIS FORM ARE LOCKED!
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
'  If (NewValue_Okay) Then
'    If (Raise_Dirty_Flag) Then
'      'STORE TO MEMORY.
'      Select Case Index
'        Case 0:         'W0.
'          Carbon.W0 = NewValue
'        Case 1:         'BB.
'          Carbon.BB = NewValue
'        Case 2:         'GM.
'          Carbon.PolanyiExponent = NewValue
'      End Select
'      'RAISE DIRTY FLAG IF NECESSARY.
'      If (Raise_Dirty_Flag) Then
'        ''THROW DIRTY FLAG.
'        'Call frmCompoProp_DirtyStatus_Throw
'      End If
'      'REFRESH WINDOW.
'      Call frmPolanyi_Refresh
'    End If
'  End If
End Sub

