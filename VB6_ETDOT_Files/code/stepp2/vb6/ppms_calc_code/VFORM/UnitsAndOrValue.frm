VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmUnitsAndOrValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "{Set by Form_Load}"
   ClientHeight    =   4560
   ClientLeft      =   8415
   ClientTop       =   1890
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelOK 
      Caption         =   "&Accept"
      Height          =   345
      Index           =   1
      Left            =   1148
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Accept changes (if any) and return to previous window"
      Top             =   4110
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancelOK 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   0
      Left            =   1148
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Cancel changes (if any) and return to previous window"
      Top             =   3750
      Width           =   1485
   End
   Begin Threed.SSFrame ssfUnits 
      Height          =   2685
      Left            =   155
      TabIndex        =   2
      Top             =   960
      Width           =   3465
      _Version        =   65536
      _ExtentX        =   6112
      _ExtentY        =   4736
      _StockProps     =   14
      Caption         =   "Units:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstUnits 
         BackColor       =   &H8000000F&
         Height          =   2205
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   3195
      End
   End
   Begin Threed.SSFrame ssfValue 
      Height          =   825
      Left            =   155
      TabIndex        =   3
      Top             =   60
      Width           =   3465
      _Version        =   65536
      _ExtentX        =   6112
      _ExtentY        =   1455
      _StockProps     =   14
      Caption         =   "Value:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Text            =   "txtData()"
         Top             =   300
         Width           =   2115
      End
      Begin VB.Label lblData 
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   2460
         TabIndex        =   6
         Top             =   330
         Visible         =   0   'False
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmUnitsAndOrValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FORM_MODE As Integer
Const FORM_MODE_UNITSONLY = 1
Const FORM_MODE_UNITSANDVALUE = 2

Public UnitType As String
Public UnitBase As String
Public UnitDisplayed As String
Public ValueInBaseUnits As Double
Public HALT_ALL_CONTROLS As Boolean

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean




Const frmUnitsAndOrValue_decl_end = True


Function frmUnitsAndOrValue_GoUnitsOnly( _
    in_UnitType As String, _
    in_UnitBase As String, _
    inout_UnitDisplayed As String, _
    out_HitCancel As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmUnitsAndOrValue
  FORM_MODE = FORM_MODE_UNITSONLY
  UnitType = in_UnitType
  UnitBase = in_UnitBase
  UnitDisplayed = inout_UnitDisplayed
  Frm.Show 1
  out_HitCancel = IIf(USER_HIT_CANCEL = True, True, False)
  If (out_HitCancel = False) Then
    inout_UnitDisplayed = UnitDisplayed
  End If
exit_normally_ThisFunc:
  frmUnitsAndOrValue_GoUnitsOnly = True
  Exit Function
exit_err_ThisFunc:
  frmUnitsAndOrValue_GoUnitsOnly = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmUnitsAndOrValue_GoUnitsOnly")
  Resume exit_err_ThisFunc
End Function
Function frmUnitsAndOrValue_GoUnitsAndValue( _
    in_UnitType As String, _
    in_UnitBase As String, _
    inout_UnitDisplayed As String, _
    inout_ValueInBaseUnits As Double, _
    out_HitCancel As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmUnitsAndOrValue
  FORM_MODE = FORM_MODE_UNITSANDVALUE
  UnitType = in_UnitType
  UnitBase = in_UnitBase
  UnitDisplayed = inout_UnitDisplayed
  ValueInBaseUnits = inout_ValueInBaseUnits
  Frm.Show 1
  out_HitCancel = IIf(USER_HIT_CANCEL = True, True, False)
  If (out_HitCancel = False) Then
    inout_UnitDisplayed = UnitDisplayed
    inout_ValueInBaseUnits = ValueInBaseUnits
  End If
exit_normally_ThisFunc:
  frmUnitsAndOrValue_GoUnitsAndValue = True
  Exit Function
exit_err_ThisFunc:
  frmUnitsAndOrValue_GoUnitsAndValue = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmUnitsAndOrValue_GoUnitsAndValue")
  Resume exit_err_ThisFunc
End Function


Sub frmUnitsAndOrValue_Populate_Units()
Dim Frm As Form
Set Frm = frmUnitsAndOrValue
  Call unitsys_register(Frm, lblData(0), txtData(0), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Function frmUnitsAndOrValue_Resize() _
    As Boolean
On Error GoTo err_ThisFunc
Const USE_MARGIN = 100
  '
  ' MAKE VALUE FRAME VISIBLE OR INVISIBLE.
  '
  If (FORM_MODE = FORM_MODE_UNITSONLY) Then
    ssfValue.Visible = False
    ssfUnits.Top = ssfValue.Top
  Else
    ssfValue.Visible = True
  End If
  cmdCancelOK(0).Top = _
      ssfUnits.Top + ssfUnits.Height + USE_MARGIN
  cmdCancelOK(1).Top = _
      cmdCancelOK(0).Top + cmdCancelOK(0).Height
  Me.Height = _
      cmdCancelOK(1).Top + cmdCancelOK(1).Height + _
      USE_MARGIN + _
      (Me.Height - Me.ScaleHeight)



exit_normally_ThisFunc:
  frmUnitsAndOrValue_Resize = True
  Exit Function
exit_err_ThisFunc:
  frmUnitsAndOrValue_Resize = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmUnitsAndOrValue_Resize")
  Resume exit_err_ThisFunc
End Function


Private Sub cmdCancelOK_Click(Index As Integer)
  Select Case Index
    Case 0:       'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:       'OK.
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  USER_HIT_CANCEL = False
  USER_HIT_OK = False
  HALT_ALL_CONTROLS = False
  Call CenterOnForm(Me, frmMain)
  Select Case FORM_MODE
    Case FORM_MODE_UNITSONLY:
      Me.Caption = "Enter New Units"
    Case FORM_MODE_UNITSANDVALUE:
      Me.Caption = "Enter New Value and Units"
  End Select
  '
  ' POPULATE THE UNITS.
  '
  Call unitsys_populate_units0( _
      lstUnits, _
      UnitType, _
      UnitDisplayed)
  Call frmUnitsAndOrValue_Populate_Units
  '
  ' FIRST RESIZE AND REFRESH.
  '
  Call frmUnitsAndOrValue_Resize
  Call frmUnitsAndOrValue_Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub lstUnits_Click()
Dim Ctl As Control
Set Ctl = lstUnits
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  If (Ctl.ListCount < 1) Then Exit Sub
  UnitDisplayed = Ctl.List(Ctl.ListIndex)
  '
  ' REFRESH WINDOW.
  '
  Call frmUnitsAndOrValue_Refresh
End Sub
Private Sub lstUnits_KeyPress(KeyAscii As Integer)
  Call lstUnits_Click
  If (KeyAscii = 13) Then
    Call cmdCancelOK_Click(1)
    Exit Sub
  End If
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
''''    Case 0
''''      StatusMessagePanel = "Type in the bed diameter"
  End Select
  ''''Call Local_GenericStatus_Set(StatusMessagePanel)
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
Dim ValueInDisplayedUnits As Double
Dim out_Found As Integer
  '
  ' NOTE: LOW AND HIGH VALUES IN BASE UNITS
  '
  Val_Low = -1E+20
  Val_High = 1E+20
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  ''''Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 0:
          ValueInDisplayedUnits = NewValue
          Call unitsys_convert0( _
              UnitType, _
              UnitDisplayed, _
              UnitBase, _
              ValueInDisplayedUnits, _
              ValueInBaseUnits, _
              out_Found)
          ValueInBaseUnits = NewValue
      End Select
      'If (Raise_Dirty_Flag) Then
      '  'THROW DIRTY FLAG.
      '  Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
      'End If
      '
      ' REFRESH WINDOW.
      '
      Call frmUnitsAndOrValue_Refresh
    End If
  End If
End Sub


