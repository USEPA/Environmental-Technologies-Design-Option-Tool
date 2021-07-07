VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmPrefEnvironment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Environment Preferences"
   ClientHeight    =   2460
   ClientLeft      =   1785
   ClientTop       =   3780
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRestoreDefaults 
      Caption         =   "Restore Defaults"
      Height          =   345
      Left            =   4170
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Restore all default settings"
      Top             =   1890
      Width           =   1485
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1515
      Left            =   345
      TabIndex        =   2
      Top             =   150
      Width           =   4305
      _Version        =   65536
      _ExtentX        =   7594
      _ExtentY        =   2672
      _StockProps     =   14
      Caption         =   "Numerical Display Format:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboSigFig 
         Height          =   315
         Index           =   2
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1020
         Width           =   1695
      End
      Begin VB.ComboBox cboSigFig 
         Height          =   315
         Index           =   1
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   1695
      End
      Begin VB.ComboBox cboSigFig 
         Height          =   315
         Index           =   0
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lbl_cboSigFig 
         Alignment       =   1  'Right Justify
         Caption         =   "All other numbers:"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   1080
         Width           =   2265
      End
      Begin VB.Label lbl_cboSigFig 
         Alignment       =   1  'Right Justify
         Caption         =   "Numbers less than 0.001:"
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   720
         Width           =   2265
      End
      Begin VB.Label lbl_cboSigFig 
         Alignment       =   1  'Right Justify
         Caption         =   "Numbers greater than 1000:"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   2265
      End
   End
   Begin VB.CommandButton cmdCancelOK 
      Caption         =   "&Accept"
      Height          =   345
      Index           =   1
      Left            =   7530
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Accept changes (if any) and return to main window"
      Top             =   1890
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancelOK 
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   0
      Left            =   6030
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Cancel changes (if any) and return to main window"
      Top             =   1890
      Width           =   1485
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1515
      Left            =   4710
      TabIndex        =   9
      Top             =   150
      Width           =   4305
      _Version        =   65536
      _ExtentX        =   7594
      _ExtentY        =   2672
      _StockProps     =   14
      Caption         =   "Miscellaneous:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lbl_cboFontSize 
         Alignment       =   1  'Right Justify
         Caption         =   "Font size of all lists:"
         Height          =   225
         Left            =   60
         TabIndex        =   11
         Top             =   360
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmPrefEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Save_TempCopy_PrefEnvironment As PrefEnvironment_Type

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Public HALT_ALL_CONTROLS As Boolean





Const frmPrefEnvironment_decl_end = True


Function frmPrefEnvironment_Go( _
    out_HitCancel As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmPrefEnvironment
  Save_TempCopy_PrefEnvironment = PrefEnvironment
  Frm.Show 1
  out_HitCancel = IIf(USER_HIT_CANCEL = True, True, False)
  If (out_HitCancel = True) Then
    PrefEnvironment = Save_TempCopy_PrefEnvironment
  Else
    Call PrefEnvironment_SaveToINI
  End If
exit_normally_ThisFunc:
  frmPrefEnvironment_Go = True
  Exit Function
exit_err_ThisFunc:
  frmPrefEnvironment_Go = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmPrefEnvironment_Go")
  Resume exit_err_ThisFunc
End Function


Function frmPrefEnvironment_PopulateFirstTime_SeveralControls() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmPrefEnvironment
Dim Ctl As Control
Dim i As Integer
  '
  ' NUMERICAL DISPLAY FORMAT SCROLLBOXES.
  '
  For i = 0 To 2
    Set Ctl = Frm.cboSigFig(i)
    Ctl.Clear
    Ctl.AddItem "3 Signif. Figures": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_3SIGFIG
    Ctl.AddItem "4 Signif. Figures": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_4SIGFIG
    Ctl.AddItem "5 Signif. Figures": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_5SIGFIG
    Ctl.AddItem "6 Signif. Figures": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_6SIGFIG
    Ctl.AddItem "0.000E+00": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_EXP3
    Ctl.AddItem "0.0000E+00": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_EXP4
    Ctl.AddItem "0.00000E+00": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_EXP5
    Ctl.AddItem "0.000": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_3PASTDEC
    Ctl.AddItem "0.0000": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_4PASTDEC
    Ctl.AddItem "0.00000": Ctl.ItemData(Ctl.NewIndex) = NUMFORMAT_5PASTDEC
  Next i
  '
  ' FONT SIZES IN LISTS.
  '
  Set Ctl = Frm.cboFontSize
  Ctl.Clear
  Ctl.AddItem "8 Point": Ctl.ItemData(Ctl.NewIndex) = 8
  Ctl.AddItem "10 Point": Ctl.ItemData(Ctl.NewIndex) = 10
  Ctl.AddItem "12 Point": Ctl.ItemData(Ctl.NewIndex) = 12
  Ctl.AddItem "14 Point": Ctl.ItemData(Ctl.NewIndex) = 14
  Ctl.AddItem "16 Point": Ctl.ItemData(Ctl.NewIndex) = 16
  Ctl.AddItem "18 Point": Ctl.ItemData(Ctl.NewIndex) = 18
  Ctl.AddItem "20 Point": Ctl.ItemData(Ctl.NewIndex) = 20
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
  frmPrefEnvironment_PopulateFirstTime_SeveralControls = True
  Exit Function
exit_err_ThisFunc:
  frmPrefEnvironment_PopulateFirstTime_SeveralControls = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmPrefEnvironment_PopulateFirstTime_SeveralControls")
  Resume exit_err_ThisFunc
End Function


Private Sub cboFontSize_Click()
Dim Ctl As Control
Set Ctl = cboFontSize
Dim New_Data As Integer
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
  New_Data = Ctl.ItemData(Ctl.ListIndex)
  With PrefEnvironment
    .FontSize_Lists = New_Data
  End With
  '
  ' REFRESH WINDOW.
  '
  ''''Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
  Call frmPrefEnvironment_Refresh
End Sub
Private Sub cboSigFig_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboSigFig(Index)
Dim New_Data As Integer
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
  New_Data = Ctl.ItemData(Ctl.ListIndex)
  With PrefEnvironment
    Select Case Index
      Case 0: .NumFormat_Greater1000 = New_Data
      Case 1: .NumFormat_Less0_001 = New_Data
      Case 2: .NumFormat_Other = New_Data
    End Select
  End With
  '
  ' REFRESH WINDOW.
  '
  ''''Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
  Call frmPrefEnvironment_Refresh
End Sub


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


Private Sub cmdRestoreDefaults_Click()
  Call PrefEnvironment_SetDefaults
  Call frmPrefEnvironment_Refresh
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  USER_HIT_CANCEL = False
  USER_HIT_OK = False
  HALT_ALL_CONTROLS = False
  Call CenterOnForm(Me, frmMain)
  Call frmPrefEnvironment_PopulateFirstTime_SeveralControls
  '
  ' FIRST REFRESH.
  '
  Call frmPrefEnvironment_Refresh
End Sub




