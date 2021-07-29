Option Strict Off
Option Explicit On
Friend Class frmTimeVarGrid
    Inherits System.Windows.Forms.Form
    Dim rs As New Resizer

    Dim FormCaption As String
    'UPGRADE_WARNING: Lower bound of array UnitType was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim UnitType(2) As String
    'UPGRADE_WARNING: Lower bound of array BaseUnits was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim BaseUnits(2) As String
    'UPGRADE_WARNING: Lower bound of array CurrentUnits was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim CurrentUnits(2) As String
    'UPGRADE_WARNING: Lower bound of array lblUnitType was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim lblUnitType(2) As String
    Dim DataRowCount As Short
    Dim MaxRows As Short
    Dim ColumnCount As Short
    Dim ColumnNames() As String
    Dim foStoreTo As VCIF1Lib.F1Book
    Dim USER_HIT_CANCEL As Boolean

    Dim frm_ActivatedYet As Boolean
    Dim frmTimeVarGrid_Is_Dirty As Boolean
    Dim HALT_cboUnits As Boolean
    Dim READY_TO_UNLOAD As Boolean




    Const frmTimeVarGrid_declarations_end As Boolean = True


    Sub frmTimeVarGrid_Run(ByRef in_FormCaption As String, ByRef in_UnitType() As String, ByRef in_BaseUnits() As String, ByRef inout_CurrentUnits() As String, ByRef in_lblUnitType() As String, ByRef inout_DataRowCount As Short, ByRef in_MaxRows As Short, ByRef in_ColumnCount As Short, ByRef in_ColumnNames() As String, ByRef in_foStoreTo As System.Windows.Forms.Control, ByRef out_HitCancel As Boolean)
        Dim i As Short
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
        'UPGRADE_WARNING: Lower bound of array ColumnNames was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        ReDim ColumnNames(in_ColumnCount)
        For i = 1 To in_ColumnCount
            ColumnNames(i) = in_ColumnNames(i)
        Next i
        foStoreTo = in_foStoreTo
        USER_HIT_CANCEL = False
        Me.ShowDialog()
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


    Sub frmTimeVarGrid_GenericStatus_Set(ByRef fn_Text As String)
        'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '   Me.sspanel_Status.Caption = fn_Text
    End Sub
    Sub frmTimeVarGrid_DirtyStatus_Set(ByRef newVal As Boolean)
        If (newVal) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object frmTimeVarGrid.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '  Me.sspanel_Dirty.Caption = "Data Changed"
            'UPGRADE_WARNING: Couldn't resolve default property of object frmTimeVarGrid.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '  Me.sspanel_Dirty.ForeColor = Color.FromArgb(QBColor(12))
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object frmTimeVarGrid.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ' Me.sspanel_Dirty.Caption = "Unchanged"
            'UPGRADE_WARNING: Couldn't resolve default property of object frmTimeVarGrid.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ' Me.sspanel_Dirty.ForeColor = Color.FromArgb(QBColor(0))
        End If
    End Sub
    Sub frmTimeVarGrid_DirtyStatus_Set_Current()
        Call frmTimeVarGrid_DirtyStatus_Set(frmTimeVarGrid_Is_Dirty)
    End Sub
    Sub frmTimeVarGrid_DirtyStatus_Throw()
        frmTimeVarGrid_Is_Dirty = True
        Call frmTimeVarGrid_DirtyStatus_Set_Current()
    End Sub
    Sub frmTimeVarGrid_DirtyStatus_Clear()
        frmTimeVarGrid_Is_Dirty = False
        Call frmTimeVarGrid_DirtyStatus_Set_Current()
    End Sub


    Sub Copy_Hidden_to_User()
        Dim F1TabsOff As Object
        Dim F1On As Object
        Dim F1ClearAll As Object
        Dim ConvFactor_Time As Double
        Dim ConvFactor_Other As Double
        Dim CurrentRows_Hidden As Short
        Dim i As Short
        'On Error GoTo err_ThisSub
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Visible = False
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.NumSheets = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foHidden.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.NumSheets = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foHidden.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'CurrentRows_Hidden = foHidden.MaxRow
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object foHidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Call GridFunc_CopyGrid(foHidden, foUser)
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.MaxCol = ColumnCount
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.MaxRow = MaxRows
        '---- CONVERT FROM BASE-UNIT DATA TO DISPLAYED-UNIT DATA.
        'COPY VALUES FROM SHEET 1 TO SHEET 2.
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.NumSheets = 2
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Sheet = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartRow = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartCol = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndRow = CurrentRows_Hidden
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndCol = ColumnCount
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditCopy()
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Sheet = 2
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartRow = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartCol = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndRow = CurrentRows_Hidden
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndCol = ColumnCount
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditPasteValues. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditPasteValues()
        '
        ' DETERMINE ConvFactor_Time.
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_convert_getfactor(UnitType(1), CurrentUnits(1)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_convert_getfactor(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ConvFactor_Time = unitsys_convert_getfactor(UnitType(1), BaseUnits(1)) / unitsys_convert_getfactor(UnitType(1), CurrentUnits(1))
        '
        ' DETERMINE ConvFactor_Other.
        '
        ''''
        ''''ConvFactor_Other = _
        'unitsys_convert_getfactor(UnitType(2), BaseUnits(2)) / _
        'unitsys_convert_getfactor(UnitType(2), CurrentUnits(2))
        ''''
        Call unitsys_convert(UnitType(2), BaseUnits(2), CurrentUnits(2), 1.0#, ConvFactor_Other)
        '
        'CONVERT TIME DATA FROM SHEET 2 TO SHEET 1.
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Sheet = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EntryRC(1, 1) = "=(Sheet2!A1)*" & Trim(Str(ConvFactor_Time))
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartRow = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartCol = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndRow = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndCol = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditCopy()
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartRow = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartCol = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndRow = CurrentRows_Hidden
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndCol = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditPaste. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditPaste()
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditCopy()
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditPasteValues. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditPasteValues()
        'CONVERT OTHER DATA FROM SHEET 2 TO SHEET 1.
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Sheet = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EntryRC(1, 2) = "=(Sheet2!B1)*" & Trim(Str(ConvFactor_Other))
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartCol = 2
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndCol = 2
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditCopy()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartCol = 2
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndRow = CurrentRows_Hidden
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndCol = ColumnCount
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditPaste. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditPaste()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditCopy()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditPasteValues. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.EditPasteValues()
        'CLEAR ALL NON-DATA ROWS ON USER GRID, IF NECESSARY.
        If (CurrentRows_Hidden < MaxRows) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'foUser.SelStartRow = CurrentRows_Hidden + 1
            ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'foUser.SelStartCol = 1
            ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'foUser.SelEndRow = MaxRows
            ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'foUser.SelEndCol = ColumnCount
            ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditClear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'foUser.EditClear(F1ClearAll)
        End If
        'REPLACE HIGHLIGHT WITH R,C=1,1 HIGHLIGHT.
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelStartCol = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.SelEndCol = 1
        ''FINISH UP THE PROCESS.
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.NumSheets = 1
        For i = 1 To ColumnCount
            'UPGRADE_WARNING: Couldn't resolve default property of object foUser.ColText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'foUser.ColText(i) = Trim(ColumnNames(i))
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.ShowHScrollBar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object F1On. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.ShowHScrollBar = F1On
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.ShowVScrollBar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''UPGRADE_WARNING: Couldn't resolve default property of object F1On. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.ShowVScrollBar = F1On
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.ShowTabs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''UPGRADE_WARNING: Couldn't resolve default property of object F1TabsOff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.ShowTabs = F1TabsOff
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Visible = True
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
    Function Copy_User_to_Hidden(ByRef Ask_Question As Boolean) As Boolean
        Dim ConvFactor_Time As Double
        Dim ConvFactor_Other As Double
        Dim CurrentRows_User As Short
        Dim i As Short
        Dim J As Short
        Dim AllZeros As Boolean
        Dim RetVal As Short
        '---- DETERMINE NUMBER OF DATA-CONTAINING ROWS.
        'A ROW WITH ANY NON-ZERO VALUE IS CONSIDERED DATA-CONTAINING.
        'ALL BLANK CELLS ARE ASSUMED TO BE ZEROS.
        Call frmTimeVarGrid_GenericStatus_Set("Detecting data-containing rows, please wait ...")
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        CurrentRows_User = MaxRows
        For i = 1 To MaxRows
            AllZeros = True
            For J = 1 To ColumnCount
                'UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumberRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'If (foUser.NumberRC(i, J) <> 0#) Then
                '	AllZeros = False
                '	Exit For
                'End If
            Next J
            If (AllZeros) Then
                CurrentRows_User = i - 1
                If (CurrentRows_User < 1) Then
                    CurrentRows_User = 1
                End If
                Exit For
            End If
        Next i
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Call frmTimeVarGrid_GenericStatus_Set("")
        If (Ask_Question) Then
            RetVal = MsgBox("There are " & Trim(Str(CurrentRows_User)) & " data-containing rows detected.  Click Yes to save, " & "or No to continue data-entry.", MsgBoxStyle.YesNo + MsgBoxStyle.Question, My.Application.Info.Title & " : Save " & Trim(Str(CurrentRows_User)) & " Rows ?")
            If (RetVal = MsgBoxResult.No) Then
                Copy_User_to_Hidden = False
                Exit Function
            End If
        End If
        '---- PERFORM THE COPY AND CONVERSION.
        Call frmTimeVarGrid_GenericStatus_Set("Storing data, please wait ...")
        'UPGRADE_WARNING: Couldn 't resolve default property of object foUser.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Visible = False
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.NumSheets = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.NumSheets = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Call GridFunc_CopyGrid(foUser, foHidden)
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''UPGRADE_WARNING: Couldn't resolve default property of object foStoreTo.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.MaxCol = foStoreTo.MaxCol
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.MaxRow = CurrentRows_User
        ''---- CONVERT FROM DISPLAYED-UNIT DATA TO BASE-UNIT DATA.
        ''COPY VALUES FROM SHEET 1 TO SHEET 2.
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.NumSheets = 2
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.Sheet = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartCol = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndRow = CurrentRows_User
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndCol = ColumnCount
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditCopy()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.Sheet = 2
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartCol = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndRow = CurrentRows_User
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndCol = ColumnCount
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditPasteValues. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditPasteValues()
        ''
        ' DETERMINE ConvFactor_Time.
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_convert_getfactor(UnitType(1), CurrentUnits(1)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_convert_getfactor(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ConvFactor_Time = 1.0# / (unitsys_convert_getfactor(UnitType(1), BaseUnits(1)) / unitsys_convert_getfactor(UnitType(1), CurrentUnits(1)))
        '
        ' DETERMINE ConvFactor_Other.
        '
        ''''
        ''''  ConvFactor_Other = 1# / _
        '''''      (unitsys_convert_getfactor(UnitType(2), BaseUnits(2)) / _
        '''''      unitsys_convert_getfactor(UnitType(2), CurrentUnits(2)))
        ''''
        Call unitsys_convert(UnitType(2), CurrentUnits(2), BaseUnits(2), 1.0#, ConvFactor_Other)
        '
        ' CONVERT TIME DATA FROM SHEET 2 TO SHEET 1.
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object foHidden.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.Sheet = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EntryRC(1, 1) = "=(Sheet2!A1)*" & Trim(Str(ConvFactor_Time))
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartCol = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndCol = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditCopy()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartCol = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndRow = CurrentRows_User
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndCol = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditPaste. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditPaste()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditCopy()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditPasteValues. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditPasteValues()
        ''CONVERT OTHER DATA FROM SHEET 2 TO SHEET 1.
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.Sheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.Sheet = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EntryRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EntryRC(1, 2) = "=(Sheet2!B1)*" & Trim(Str(ConvFactor_Other))
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartCol = 2
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndCol = 2
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditCopy()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartCol = 2
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndRow = CurrentRows_User
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndCol = ColumnCount
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditPaste. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditPaste()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditCopy()
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.EditPasteValues. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.EditPasteValues()
        ''REPLACE HIGHLIGHT WITH R,C=1,1 HIGHLIGHT.
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelStartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelStartCol = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndRow = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.SelEndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.SelEndCol = 1
        ''FINISH UP THE PROCESS.
        ''UPGRADE_WARNING: Couldn't resolve default property of object foHidden.NumSheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foHidden.NumSheets = 1
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Visible = True
        Call frmTimeVarGrid_GenericStatus_Set("")
        'UPDATE ROW COUNT.
        DataRowCount = CurrentRows_User
        'RETURN "OKAY" MESSAGE.
        Copy_User_to_Hidden = True
    End Function


    'UPGRADE_WARNING: Event cboUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboUnits.SelectedIndexChanged
        Dim Index As Short = cboUnits.GetIndex(eventSender)
        Dim RetValBool As Boolean
        If (HALT_cboUnits) Then Exit Sub
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Call frmTimeVarGrid_GenericStatus_Set("Converting units, please wait ...")
        RetValBool = Copy_User_to_Hidden(False)
        If (RetValBool = False) Then Exit Sub
        CurrentUnits(Index + 1) = VB6.GetItemString(cboUnits(Index), cboUnits(Index).SelectedIndex)
        Call Copy_Hidden_to_User()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Call frmTimeVarGrid_GenericStatus_Set("")
    End Sub


    Private Sub cmdCancelOK_Click(ByRef Index As Short)
        Select Case Index
            Case 0 'CANCEL.
                READY_TO_UNLOAD = True
                USER_HIT_CANCEL = True
                Me.Close()
                Exit Sub
            Case 1 'OK.
                Call frmTimeVarGrid_GenericStatus_Set("Storing data, please wait ...")
                If (Copy_User_to_Hidden(True) = False) Then
                    Exit Sub
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object Me.foHidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'Call GridFunc_CopyGrid((Me.foHidden), foStoreTo)
                Call frmTimeVarGrid_GenericStatus_Set("")
                READY_TO_UNLOAD = True
                USER_HIT_CANCEL = False
                Me.Close()
                Exit Sub
        End Select
    End Sub


    'UPGRADE_WARNING: Form event frmTimeVarGrid.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmTimeVarGrid_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        If (frm_ActivatedYet = False) Then
            frm_ActivatedYet = True
            System.Windows.Forms.Application.DoEvents()
            'UPGRADE_WARNING: Couldn't resolve default property of object foUser.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'foUser.Visible = False
            Call frmTimeVarGrid_GenericStatus_Set("Loading data, please wait ...")
            System.Windows.Forms.Application.DoEvents()
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.foHidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ' Call GridFunc_CopyGrid(foStoreTo, (Me.foHidden))
            System.Windows.Forms.Application.DoEvents()
            Call frmTimeVarGrid_GenericStatus_Set("Loading data, please wait ...")
            Call Copy_Hidden_to_User()
            System.Windows.Forms.Application.DoEvents()
            If (DataRowCount = 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumberRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ' foUser.NumberRC(1, 1) = 0#
                'UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumberRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                '  foUser.NumberRC(1, 2) = 0#
            End If
            Call frmTimeVarGrid_GenericStatus_Set("")
            'UPGRADE_WARNING: Couldn't resolve default property of object foUser.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ' foUser.Visible = True
        End If
    End Sub
    Private Sub frmTimeVarGrid_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        rs.FindAllControls(Me)

        Dim i As Short
        '
        ' MISC INITS.
        Call CenterOnScreen(Me) ', frmInfluentEdit)
        frm_ActivatedYet = False
        Call frmTimeVarGrid_DirtyStatus_Clear()
        'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'sspanel_Dirty.Caption = "" 'DIRTY FUNCTIONALITY NOT ADDED YET.
        Call frmTimeVarGrid_GenericStatus_Set("")
        Me.Text = FormCaption
        lblData(0).Text = lblUnitType(1)
        lblData(1).Text = lblUnitType(2)
        HALT_cboUnits = False
        READY_TO_UNLOAD = False
        '
        ' CLEAR OUT FIRST ROW (IF ANYTHING IS THERE).
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumberRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.NumberRC(1, 1) = 0#
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.NumberRC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ' foUser.NumberRC(1, 2) = 0#
        '
        ' SETUP GRID.
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.MaxRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ' foUser.MaxRow = MaxRows
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.MaxCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ' foUser.MaxCol = ColumnCount
        HALT_cboUnits = True
        Call unitsys_populate_units0(cboUnits(0), UnitType(1), CurrentUnits(1))
        Call unitsys_populate_units0(cboUnits(1), UnitType(2), CurrentUnits(2))
        HALT_cboUnits = False
    End Sub
    Private Sub frmTimeVarGrid_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If (READY_TO_UNLOAD = False) Then
            Cancel = True
        End If
        eventArgs.Cancel = Cancel
    End Sub
    'UPGRADE_WARNING: Event frmTimeVarGrid.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub frmTimeVarGrid_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        rs.ResizeAllControls(Me)

        Dim USE_MARGIN As Integer
        Dim XXX As Integer
        If (Me.WindowState = 1) Then
            'CAN'T RESIZE WHEN MINIMIZED.
            Exit Sub
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object foUser.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'USE_MARGIN = foUser.Left
        'XXX = VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - USE_MARGIN * 2
        'If (XXX < 1000) Then XXX = 1000
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Width = XXX
        ''UPGRADE_WARNING: Couldn't resolve default property of object sspanel_Holder.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'XXX = VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - foUser.Top - USE_MARGIN - sspanel_Holder.Height
        'If (XXX < 1000) Then XXX = 1000
        ''UPGRADE_WARNING: Couldn't resolve default property of object foUser.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'foUser.Height = XXX
    End Sub


    Public Sub mnuEditItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditItem.Click
        Dim Index As Short = mnuEditItem.GetIndex(eventSender)
        Select Case Index
            Case 10 'COPY.
                'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditCopy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'foUser.EditCopy()
            Case 20 'PASTE.
                'UPGRADE_WARNING: Couldn't resolve default property of object foUser.EditPaste. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ' foUser.EditPaste()
        End Select
    End Sub

    Private Sub _cmdCancelOK_0_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_0.Click

    End Sub

    Private Sub frmTimeVarGrid_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles Me.MouseDoubleClick

    End Sub
End Class