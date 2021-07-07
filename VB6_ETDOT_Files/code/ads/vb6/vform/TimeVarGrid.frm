VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmTimeVarGrid 
   Caption         =   "{Caption}"
   ClientHeight    =   5685
   ClientLeft      =   960
   ClientTop       =   1980
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7590
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
      Left            =   3750
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   1545
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
      Left            =   3750
      Style           =   2  'Dropdown List
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   1545
   End
   Begin VCIF1Lib.F1Book foUser 
      Height          =   3975
      Left            =   90
      OleObjectBlob   =   "TimeVarGrid.frx":0000
      TabIndex        =   0
      Top             =   960
      Width           =   5985
   End
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes to this grid"
      Top             =   450
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2293
      _ExtentY        =   714
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
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes you have made to this grid"
      Top             =   60
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2293
      _ExtentY        =   714
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
   Begin VCIF1Lib.F1Book foHidden 
      Height          =   3975
      Left            =   6390
      OleObjectBlob   =   "TimeVarGrid.frx":053F
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   5985
   End
   Begin Threed.SSPanel sspanel_Holder 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   5280
      Width           =   7590
      _Version        =   65536
      _ExtentX        =   13388
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
      Begin Threed.SSPanel sspanel_Status 
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Width           =   9855
         _Version        =   65536
         _ExtentX        =   17383
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
         Alignment       =   1
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2805
      Left            =   6270
      TabIndex        =   10
      Top             =   4590
      Visible         =   0   'False
      Width           =   2805
      _Version        =   65536
      _ExtentX        =   4948
      _ExtentY        =   4948
      _StockProps     =   14
      Caption         =   "Invisible"
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
      Begin Threed.SSPanel sspanel_Dirty 
         Height          =   285
         Left            =   270
         TabIndex        =   11
         Top             =   510
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
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      Caption         =   "{Whatever} Units:"
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
      Left            =   1470
      TabIndex        =   4
      Top             =   540
      Width           =   2200
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Units:"
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
      Left            =   1470
      TabIndex        =   2
      Top             =   120
      Width           =   2200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As ..."
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   49
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Page Setup ..."
         Enabled         =   0   'False
         Index           =   50
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Printer Setup ..."
         Enabled         =   0   'False
         Index           =   60
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print ..."
         Enabled         =   0   'False
         Index           =   70
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   99
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Close"
         Index           =   100
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Copy"
         Index           =   10
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Paste"
         Index           =   20
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmTimeVarGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FormCaption As String
Dim UnitType(1 To 2) As String
Dim BaseUnits(1 To 2) As String
Dim CurrentUnits(1 To 2) As String
Dim lblUnitType(1 To 2) As String
Dim DataRowCount As Integer
Dim MaxRows As Integer
Dim ColumnCount As Integer
Dim ColumnNames() As String
Dim foStoreTo As Control
Dim USER_HIT_CANCEL As Boolean

Dim frm_ActivatedYet As Boolean
Dim frmTimeVarGrid_Is_Dirty As Boolean
Dim HALT_cboUnits As Boolean
Dim READY_TO_UNLOAD As Boolean




Const frmTimeVarGrid_declarations_end = True


Sub frmTimeVarGrid_Run( _
    in_FormCaption As String, _
    in_UnitType() As String, _
    in_BaseUnits() As String, _
    inout_CurrentUnits() As String, _
    in_lblUnitType() As String, _
    inout_DataRowCount As Integer, _
    in_MaxRows As Integer, _
    in_ColumnCount As Integer, _
    in_ColumnNames() As String, _
    in_foStoreTo As Control, _
    out_HitCancel As Boolean)
Dim i As Integer
  FormCaption = in_FormCaption
  For i = 1 To 2
    UnitType(i) = in_UnitType(i)
    BaseUnits(i) = in_BaseUnits(i)
    CurrentUnits(i) = inout_CurrentUnits(i)
    lblUnitType(i) = in_lblUnitType(i)
  Next i
  DataRowCount = inout_DataRowCount
  MaxRows = in_MaxRows
  ColumnCount = in_ColumnCount
  ReDim ColumnNames(1 To in_ColumnCount)
  For i = 1 To in_ColumnCount
    ColumnNames(i) = in_ColumnNames(i)
  Next i
  Set foStoreTo = in_foStoreTo
  USER_HIT_CANCEL = False
  frmTimeVarGrid.Show 1
  If (USER_HIT_CANCEL) Then
    out_HitCancel = True
  Else
    out_HitCancel = False
    For i = 1 To 2
      inout_CurrentUnits(i) = CurrentUnits(i)
    Next i
    inout_DataRowCount = DataRowCount
  End If
End Sub


Sub frmTimeVarGrid_GenericStatus_Set(fn_Text As String)
  Me.sspanel_Status = fn_Text
End Sub
Sub frmTimeVarGrid_DirtyStatus_Set(newVal As Boolean)
  If (newVal) Then
    frmTimeVarGrid.sspanel_Dirty = "Data Changed"
    frmTimeVarGrid.sspanel_Dirty.ForeColor = QBColor(12)
  Else
    frmTimeVarGrid.sspanel_Dirty = "Unchanged"
    frmTimeVarGrid.sspanel_Dirty.ForeColor = QBColor(0)
  End If
End Sub
Sub frmTimeVarGrid_DirtyStatus_Set_Current()
  Call frmTimeVarGrid_DirtyStatus_Set(frmTimeVarGrid_Is_Dirty)
End Sub
Sub frmTimeVarGrid_DirtyStatus_Throw()
  frmTimeVarGrid_Is_Dirty = True
  Call frmTimeVarGrid_DirtyStatus_Set_Current
End Sub
Sub frmTimeVarGrid_DirtyStatus_Clear()
  frmTimeVarGrid_Is_Dirty = False
  Call frmTimeVarGrid_DirtyStatus_Set_Current
End Sub


Sub Copy_Hidden_to_User()
Dim ConvFactor_Time As Double
Dim ConvFactor_Other As Double
Dim CurrentRows_Hidden As Integer
Dim i As Integer
  'On Error GoTo err_ThisSub
  foUser.Visible = False
  foUser.NumSheets = 1
  foHidden.NumSheets = 1
  CurrentRows_Hidden = foHidden.MaxRow
  Call GridFunc_CopyGrid(foHidden, foUser)
  foUser.MaxCol = ColumnCount
  foUser.MaxRow = MaxRows
  '---- CONVERT FROM BASE-UNIT DATA TO DISPLAYED-UNIT DATA.
  'COPY VALUES FROM SHEET 1 TO SHEET 2.
  foUser.NumSheets = 2
  foUser.Sheet = 1
  foUser.SelStartRow = 1
  foUser.SelStartCol = 1
  foUser.SelEndRow = CurrentRows_Hidden
  foUser.SelEndCol = ColumnCount
  foUser.EditCopy
  foUser.Sheet = 2
  foUser.SelStartRow = 1
  foUser.SelStartCol = 1
  foUser.SelEndRow = CurrentRows_Hidden
  foUser.SelEndCol = ColumnCount
  foUser.EditPasteValues
  '
  ' DETERMINE ConvFactor_Time.
  '
  ConvFactor_Time = _
      unitsys_convert_getfactor(UnitType(1), BaseUnits(1)) / _
      unitsys_convert_getfactor(UnitType(1), CurrentUnits(1))
  '
  ' DETERMINE ConvFactor_Other.
  '
  ''''
  ''''ConvFactor_Other = _
      unitsys_convert_getfactor(UnitType(2), BaseUnits(2)) / _
      unitsys_convert_getfactor(UnitType(2), CurrentUnits(2))
  ''''
  Call unitsys_convert( _
      UnitType(2), _
      BaseUnits(2), _
      CurrentUnits(2), _
      1#, _
      ConvFactor_Other)
  '
  'CONVERT TIME DATA FROM SHEET 2 TO SHEET 1.
  '
  foUser.Sheet = 1
  foUser.EntryRC(1, 1) = "=(Sheet2!A1)*" & Trim$(Str$(ConvFactor_Time))
  foUser.SelStartRow = 1
  foUser.SelStartCol = 1
  foUser.SelEndRow = 1
  foUser.SelEndCol = 1
  foUser.EditCopy
  foUser.SelStartRow = 1
  foUser.SelStartCol = 1
  foUser.SelEndRow = CurrentRows_Hidden
  foUser.SelEndCol = 1
  foUser.EditPaste
  foUser.EditCopy
  foUser.EditPasteValues
  'CONVERT OTHER DATA FROM SHEET 2 TO SHEET 1.
  foUser.Sheet = 1
  foUser.EntryRC(1, 2) = "=(Sheet2!B1)*" & Trim$(Str$(ConvFactor_Other))
  foUser.SelStartRow = 1
  foUser.SelStartCol = 2
  foUser.SelEndRow = 1
  foUser.SelEndCol = 2
  foUser.EditCopy
  foUser.SelStartRow = 1
  foUser.SelStartCol = 2
  foUser.SelEndRow = CurrentRows_Hidden
  foUser.SelEndCol = ColumnCount
  foUser.EditPaste
  foUser.EditCopy
  foUser.EditPasteValues
  'CLEAR ALL NON-DATA ROWS ON USER GRID, IF NECESSARY.
  If (CurrentRows_Hidden < MaxRows) Then
    foUser.SelStartRow = CurrentRows_Hidden + 1
    foUser.SelStartCol = 1
    foUser.SelEndRow = MaxRows
    foUser.SelEndCol = ColumnCount
    foUser.EditClear F1ClearAll
  End If
  'REPLACE HIGHLIGHT WITH R,C=1,1 HIGHLIGHT.
  foUser.SelStartRow = 1
  foUser.SelStartCol = 1
  foUser.SelEndRow = 1
  foUser.SelEndCol = 1
  'FINISH UP THE PROCESS.
  foUser.NumSheets = 1
  For i = 1 To ColumnCount
    foUser.ColText(i) = Trim$(ColumnNames(i))
  Next i
  foUser.ShowHScrollBar = F1On
  foUser.ShowVScrollBar = F1On
  foUser.ShowTabs = F1TabsOff
  foUser.Visible = True
exit_normally_ThisSub:
  'Copy_Hidden_to_User = True
  Exit Sub
exit_err_ThisSub:
  'Copy_Hidden_to_User = False
  Exit Sub
err_ThisSub:
  Call Show_Trapped_Error("Copy_Hidden_to_User")
  Resume exit_err_ThisSub
End Sub
'RETURNS:
'    FALSE = USER CANCELLED.
'    TRUE = USER OKAYED.
Function Copy_User_to_Hidden( _
    Ask_Question As Boolean) As Boolean
Dim ConvFactor_Time As Double
Dim ConvFactor_Other As Double
Dim CurrentRows_User As Integer
Dim i As Integer
Dim J As Integer
Dim AllZeros As Boolean
Dim RetVal As Integer
  '---- DETERMINE NUMBER OF DATA-CONTAINING ROWS.
  'A ROW WITH ANY NON-ZERO VALUE IS CONSIDERED DATA-CONTAINING.
  'ALL BLANK CELLS ARE ASSUMED TO BE ZEROS.
  Call frmTimeVarGrid_GenericStatus_Set("Detecting data-containing rows, please wait ...")
  Me.MousePointer = 11
  CurrentRows_User = MaxRows
  For i = 1 To MaxRows
    AllZeros = True
    For J = 1 To ColumnCount
      If (foUser.NumberRC(i, J) <> 0#) Then
        AllZeros = False
        Exit For
      End If
    Next J
    If (AllZeros) Then
      CurrentRows_User = i - 1
      If (CurrentRows_User < 1) Then
        CurrentRows_User = 1
      End If
      Exit For
    End If
  Next i
  Me.MousePointer = 0
  Call frmTimeVarGrid_GenericStatus_Set("")
  If (Ask_Question) Then
    RetVal = MsgBox("There are " & Trim$(Str$(CurrentRows_User)) & _
        " data-containing rows detected.  Click Yes to save, " & _
        "or No to continue data-entry.", _
        vbYesNo + vbQuestion, _
        App.Title & " : Save " & Trim$(Str$(CurrentRows_User)) & _
        " Rows ?")
    If (RetVal = vbNo) Then
      Copy_User_to_Hidden = False
      Exit Function
    End If
  End If
  '---- PERFORM THE COPY AND CONVERSION.
  Call frmTimeVarGrid_GenericStatus_Set("Storing data, please wait ...")
  foUser.Visible = False
  foUser.NumSheets = 1
  foHidden.NumSheets = 1
  Call GridFunc_CopyGrid(foUser, foHidden)
  foHidden.MaxCol = foStoreTo.MaxCol
  foHidden.MaxRow = CurrentRows_User
  '---- CONVERT FROM DISPLAYED-UNIT DATA TO BASE-UNIT DATA.
  'COPY VALUES FROM SHEET 1 TO SHEET 2.
  foHidden.NumSheets = 2
  foHidden.Sheet = 1
  foHidden.SelStartRow = 1
  foHidden.SelStartCol = 1
  foHidden.SelEndRow = CurrentRows_User
  foHidden.SelEndCol = ColumnCount
  foHidden.EditCopy
  foHidden.Sheet = 2
  foHidden.SelStartRow = 1
  foHidden.SelStartCol = 1
  foHidden.SelEndRow = CurrentRows_User
  foHidden.SelEndCol = ColumnCount
  foHidden.EditPasteValues
  '
  ' DETERMINE ConvFactor_Time.
  '
  ConvFactor_Time = 1# / _
      (unitsys_convert_getfactor(UnitType(1), BaseUnits(1)) / _
      unitsys_convert_getfactor(UnitType(1), CurrentUnits(1)))
  '
  ' DETERMINE ConvFactor_Other.
  '
''''
''''  ConvFactor_Other = 1# / _
''''      (unitsys_convert_getfactor(UnitType(2), BaseUnits(2)) / _
''''      unitsys_convert_getfactor(UnitType(2), CurrentUnits(2)))
''''
  Call unitsys_convert( _
      UnitType(2), _
      CurrentUnits(2), _
      BaseUnits(2), _
      1#, _
      ConvFactor_Other)
  '
  ' CONVERT TIME DATA FROM SHEET 2 TO SHEET 1.
  '
  foHidden.Sheet = 1
  foHidden.EntryRC(1, 1) = "=(Sheet2!A1)*" & Trim$(Str$(ConvFactor_Time))
  foHidden.SelStartRow = 1
  foHidden.SelStartCol = 1
  foHidden.SelEndRow = 1
  foHidden.SelEndCol = 1
  foHidden.EditCopy
  foHidden.SelStartRow = 1
  foHidden.SelStartCol = 1
  foHidden.SelEndRow = CurrentRows_User
  foHidden.SelEndCol = 1
  foHidden.EditPaste
  foHidden.EditCopy
  foHidden.EditPasteValues
  'CONVERT OTHER DATA FROM SHEET 2 TO SHEET 1.
  foHidden.Sheet = 1
  foHidden.EntryRC(1, 2) = "=(Sheet2!B1)*" & Trim$(Str$(ConvFactor_Other))
  foHidden.SelStartRow = 1
  foHidden.SelStartCol = 2
  foHidden.SelEndRow = 1
  foHidden.SelEndCol = 2
  foHidden.EditCopy
  foHidden.SelStartRow = 1
  foHidden.SelStartCol = 2
  foHidden.SelEndRow = CurrentRows_User
  foHidden.SelEndCol = ColumnCount
  foHidden.EditPaste
  foHidden.EditCopy
  foHidden.EditPasteValues
  'REPLACE HIGHLIGHT WITH R,C=1,1 HIGHLIGHT.
  foHidden.SelStartRow = 1
  foHidden.SelStartCol = 1
  foHidden.SelEndRow = 1
  foHidden.SelEndCol = 1
  'FINISH UP THE PROCESS.
  foHidden.NumSheets = 1
  foUser.Visible = True
  Call frmTimeVarGrid_GenericStatus_Set("")
  'UPDATE ROW COUNT.
  DataRowCount = CurrentRows_User
  'RETURN "OKAY" MESSAGE.
  Copy_User_to_Hidden = True
End Function


Private Sub cboUnits_Click(Index As Integer)
Dim RetValBool As Boolean
  If (HALT_cboUnits) Then Exit Sub
  Me.MousePointer = 11
  Call frmTimeVarGrid_GenericStatus_Set("Converting units, please wait ...")
  RetValBool = Copy_User_to_Hidden(False)
  If (RetValBool = False) Then Exit Sub
  CurrentUnits(Index + 1) = cboUnits(Index).List(cboUnits(Index).ListIndex)
  Call Copy_Hidden_to_User
  Me.MousePointer = 0
  Call frmTimeVarGrid_GenericStatus_Set("")
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
  Select Case Index
    Case 0:     'CANCEL.
      READY_TO_UNLOAD = True
      USER_HIT_CANCEL = True
      Unload Me
      Exit Sub
    Case 1:     'OK.
      Call frmTimeVarGrid_GenericStatus_Set("Storing data, please wait ...")
      If (Copy_User_to_Hidden(True) = False) Then
        Exit Sub
      End If
      Call GridFunc_CopyGrid(Me.foHidden, foStoreTo)
      Call frmTimeVarGrid_GenericStatus_Set("")
      READY_TO_UNLOAD = True
      USER_HIT_CANCEL = False
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub Form_Activate()
  If (frm_ActivatedYet = False) Then
    frm_ActivatedYet = True
    DoEvents
    foUser.Visible = False
    Call frmTimeVarGrid_GenericStatus_Set("Loading data, please wait ...")
    DoEvents
    Call GridFunc_CopyGrid(foStoreTo, Me.foHidden)
    DoEvents
    Call frmTimeVarGrid_GenericStatus_Set("Loading data, please wait ...")
    Call Copy_Hidden_to_User
    DoEvents
    If (DataRowCount = 0) Then
      foUser.NumberRC(1, 1) = 0#
      foUser.NumberRC(1, 2) = 0#
    End If
    Call frmTimeVarGrid_GenericStatus_Set("")
    foUser.Visible = True
  End If
End Sub
Private Sub Form_Load()
Dim i As Integer
  '
  ' MISC INITS.
  Call CenterOnScreen(Me)   ', frmInfluentEdit)
  frm_ActivatedYet = False
  Call frmTimeVarGrid_DirtyStatus_Clear
  sspanel_Dirty = ""          'DIRTY FUNCTIONALITY NOT ADDED YET.
  Call frmTimeVarGrid_GenericStatus_Set("")
  Me.Caption = FormCaption
  lblData(0).Caption = lblUnitType(1)
  lblData(1).Caption = lblUnitType(2)
  HALT_cboUnits = False
  READY_TO_UNLOAD = False
  '
  ' CLEAR OUT FIRST ROW (IF ANYTHING IS THERE).
  foUser.NumberRC(1, 1) = 0#
  foUser.NumberRC(1, 2) = 0#
  '
  ' SETUP GRID.
  foUser.MaxRow = MaxRows
  foUser.MaxCol = ColumnCount
  HALT_cboUnits = True
  Call unitsys_populate_units0(cboUnits(0), UnitType(1), CurrentUnits(1))
  Call unitsys_populate_units0(cboUnits(1), UnitType(2), CurrentUnits(2))
  HALT_cboUnits = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (READY_TO_UNLOAD = False) Then
    Cancel = True
  End If
End Sub
Private Sub Form_Resize()
Dim USE_MARGIN As Long
Dim XXX As Long
  If (Me.WindowState = 1) Then
    'CAN'T RESIZE WHEN MINIMIZED.
    Exit Sub
  End If
  USE_MARGIN = foUser.Left
  XXX = Me.ScaleWidth - USE_MARGIN * 2
  If (XXX < 1000) Then XXX = 1000
  foUser.Width = XXX
  XXX = Me.ScaleHeight - foUser.Top - USE_MARGIN - sspanel_Holder.Height
  If (XXX < 1000) Then XXX = 1000
  foUser.Height = XXX
End Sub


Private Sub mnuEditItem_Click(Index As Integer)
  Select Case Index
    Case 10:    'COPY.
      foUser.EditCopy
    Case 20:    'PASTE.
      foUser.EditPaste
  End Select
End Sub
