Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmCompoProp
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Dim FORM_MODE As Short
	Const FORM_MODE_ADDNEW As Short = 1
	Const FORM_MODE_EDIT As Short = 2
	Dim USER_HIT_CANCEL As Boolean = False
	Dim USER_HIT_OK As Boolean = False
	Dim frmCompoProp_Is_Dirty As Boolean
	
	Dim START_AT_COMPNUMBER As Short
	Dim CurrentCompNumber As Short
	'UPGRADE_WARNING: Lower bound of array TempComponents was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	'UPGRADE_WARNING: Array TempComponents may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
	Dim TempComponents(Number_Compo_Max) As ComponentPropertyType
	'USED TO STORE TEMPORARY COPIES OF COMPONENTS.
	'THIS ALLOWS A ROLLBACK IF THE USER HITS CANCEL.
	
	Dim HALT_CBOCHEMNAME As Boolean
	Dim HALT_CBOSOURCE As Boolean
	
	
	
	
	
	Const frmCompoProp_declarations_end As Boolean = True
	
	
	Sub frmCompoProp_Add(ByRef OUTPUT_Raise_Dirty_Flag As Boolean)
		FORM_MODE = FORM_MODE_ADDNEW
		Me.ShowDialog()

		If (USER_HIT_OK) Then
			OUTPUT_Raise_Dirty_Flag = True
		Else
			OUTPUT_Raise_Dirty_Flag = False
		End If
	End Sub
	Sub frmCompoProp_Edit(ByRef OUTPUT_Raise_Dirty_Flag As Boolean, ByRef INPUT_Start_At_CompNumber As Object)
		FORM_MODE = FORM_MODE_EDIT
		'UPGRADE_WARNING: Couldn't resolve default property of object INPUT_Start_At_CompNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		START_AT_COMPNUMBER = INPUT_Start_At_CompNumber
		Me.ShowDialog()
		If (USER_HIT_OK) Then
			OUTPUT_Raise_Dirty_Flag = True
		Else
			OUTPUT_Raise_Dirty_Flag = False
		End If
	End Sub
	'RETURNS:
	'- true = it's okay to unload now.
	'- false = cancel the unload.
	Function frmCompoProp_Query_Unload() As Short
		Dim RetVal As Short
		Dim msg As String
		If (Not frmCompoProp_Is_Dirty) Then
			frmCompoProp_Query_Unload = True
			Exit Function
		End If
		msg = "Are you sure you want to abandon the changes " & "made to " & IIf(Number_Component = 1, "this component", "these components") & " ?" & vbCrLf & vbCrLf & "If you want to abandon the changes, click Yes." & vbCrLf & "If you want to save the " & "changes, click No, then click OK."
		RetVal = MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.YesNo, AppName_For_Display_Short & " : Abandon Changes ?")
		Select Case RetVal
			Case MsgBoxResult.Yes
				frmCompoProp_Query_Unload = True
				Exit Function
			Case MsgBoxResult.No
				frmCompoProp_Query_Unload = False
				Exit Function
		End Select
	End Function
	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.CancelButton.Enabled = False
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdImportFromFile.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdImportFromFile.Enabled = False
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdImportClipboard.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdImportClipboard.Enabled = False
		End If
	End Sub
	
	
	Sub Update_Display_of_KFreundlich(ByRef IndexOfChangedProperty As Short)
		''Sub Update_Display_of_KFreundlich(IndexOfChangedProperty As Integer, OldPropertyValue As Double)
		''NOTE: THIS SUBROUTINE IS _NEVER_ CALLED UNLESS THE PROPERTY HAS CHANGED,
		''SO THE VALUE OldPropertyValue IS NO LONGER NEEDED.
		'Dim NoUnitsOnChangedProperty As Integer
		'Dim OldK As String
		'Dim NewK As String
		'Dim KUnits As String
		'Dim temp As String
		'Dim CheckINI As String
		'
		'Dim A_Property_Changed As Integer
		'Dim The_MW_Changed As Integer
		'Dim Invalidate_Isotherm_K As Integer
		'Dim Invalidate_IPES_K As Integer
		'Dim K_Reverted As Integer
		'
		'Dim ReverseConversionFactor As Double
		'Dim ThisValue As Double
		'Dim i As Integer
		'
		'Dim Current_CheckValue As Integer
		'
		'  'THE PHILOSOPHY OF FREUNDLICH K DATA ENTRY.
		'  '==========================================
		'  'If the source of the K and 1/n for a chemical is:
		'  '(1). User-entry.  They may do what they wish with any chemical
		'  '     property.  If they change the 1/n or MW, the internally-stored
		'  '     value for Freundlich K is updated so that it remains constant
		'  '     on the screen.
		'  '(2). IPES Routines.  They should not modify any chemical property.
		'  '     If they do, a message box will pop up which tells them that
		'  '     their K and 1/n from IPES have been invalidated, and their
		'  '     K and 1/n will revert to the user-entered numbers.
		'  '(3). Isotherm Database.  They should not modify molecular weight.
		'  '     If they do, a message box will pop up which tells them that
		'  '     their K and 1/n from the Isotherm Database have been
		'  '     invalidated, and their K and 1/n will revert to the user-entered
		'  '     numbers.
		'
		'  'Update display of Freundlich K, which is dependent upon MW and 1/n
		'  '(depending on which units for Freundlich K have been chosen).
		'  OldK = txtDataComponentProperty(5).Text
		'  Call txtPropUnits_Click(5)
		'  NewK = txtDataComponentProperty(5).Text
		'  KUnits = txtPropUnits(5).List(txtPropUnits(5).ListIndex)
		'
		'  '----- INVALIDATE K AND 1/N FROM IPES OR ISOTHERM DATABASE IF NECESSARY
		'  A_Property_Changed = False
		'  The_MW_Changed = False
		'  Invalidate_Isotherm_K = False
		'  Invalidate_IPES_K = False
		'  K_Reverted = False
		'  If ((IndexOfChangedProperty = 1) Or (IndexOfChangedProperty = 2) Or (IndexOfChangedProperty = 3) Or (IndexOfChangedProperty = 4) Or (IndexOfChangedProperty = 7) Or (IndexOfChangedProperty = 8) Or (IndexOfChangedProperty = 9) Or (IndexOfChangedProperty = 10)) Then
		'    If (OldPropertyValue <> txtDataComponentProperty(IndexOfChangedProperty).Text) Then
		'      A_Property_Changed = True
		'    End If
		'  End If
		'  If (IndexOfChangedProperty = 1) Then
		'    If (OldPropertyValue <> txtDataComponentProperty(IndexOfChangedProperty).Text) Then
		'      The_MW_Changed = True
		'    End If
		'  End If
		'  If (A_Property_Changed) Then
		'    'The IPES K and 1/n are invalidated if any property has changed.
		'    If (Component(0).IPESResult_K <> -1) Then
		'      Invalidate_IPES_K = True
		'    End If
		'  End If
		'  If (The_MW_Changed) Then
		'    'The Isotherm Database K and 1/n are invalidated if MW has changed.
		'    If (Component(0).IsothermDB_K <> -1) Then
		'      Invalidate_Isotherm_K = True
		'    End If
		'  End If
		'
		'  If (Invalidate_IPES_K) Then
		'    '-- IPES K and 1/n are now invalid!
		'    Component(0).IPESResult_K = -1
		'    Component(0).IPESResult_OneOverN = -1
		'    If (cboSource.ListIndex = 1) Then
		'      '-- Change from IPES to user-entry.
		'      cboSource.ListIndex = 2
		'      K_Reverted = True
		'    End If
		'    Call Update_cboSource
		'  End If
		'  If (Invalidate_Isotherm_K) Then
		'    '-- Isotherm Database K and 1/n are now invalid!
		'    Component(0).IsothermDB_K = -1
		'    Component(0).IsothermDB_OneOverN = -1
		'    If (cboSource.ListIndex = 0) Then
		'      '-- Change from IPES to user-entry.
		'      cboSource.ListIndex = 2
		'      K_Reverted = True
		'    End If
		'    Call Update_cboSource
		'  End If
		'
		'  '-- Inform user of what the hell just happened.
		'  If ((Invalidate_IPES_K) Or (Invalidate_Isotherm_K)) Then
		'    temp = "A property has changed.  "
		'    temp = temp & "This has caused the values of Freundlich K and 1/n from the "
		'    If (Invalidate_IPES_K) Then
		'      temp = temp & "IPES Results"
		'    End If
		'    If (Invalidate_Isotherm_K) Then
		'      If (Invalidate_IPES_K) Then
		'        temp = temp & " and the "
		'      End If
		'      temp = temp & "Isotherm Database"
		'    End If
		'    temp = temp & " to become invalidated for this chemical."
		'    If (K_Reverted) Then
		'      temp = temp & "  This chemical has reverted to user-entered K and 1/n."
		'    End If
		'    MsgBox temp, MB_ICONEXCLAMATION, AppName_For_Display_long
		'''''''    Exit Sub
		'  End If
		'
		'  '----- UPON CHANGE OF MW OR FREUNDLICH 1/N, UPDATE K
		'  If ((IndexOfChangedProperty = 1) Or (IndexOfChangedProperty = 6)) Then
		'    If (OldPropertyValue <> txtDataComponentProperty(IndexOfChangedProperty).Text) Then
		'      '--- Note: Freundlich K is stored internally in (mg/g)*(L/mg)^(1/n).
		'      '... In order to keep K constant in the currently-displayed set of units,
		'      '... it is (sometimes) necessary to change the internally-stored value.
		'      '... The user should be informed of what the heck is going on here.
		'
		'      'STEP ONE: Determine the new K in (mg/g)*(L/mg)^(1/n) units
		'      ReverseConversionFactor = 1 / KFreundlichConversionFactor(CInt(txtPropUnits(5).ListIndex), Component(0).Use_OneOverN, Component(0).MW)
		'      Component(0).Use_K = CDbl(OldK) * ReverseConversionFactor
		'      'Update display of K:
		'      Call txtPropUnits_Click(5)
		'
		'      'STEP TWO: Inform the user of what just happened
		'      'frmGoAway_Caption = "Warning: Freundlich K Updated"
		'      If (IndexOfChangedProperty = 1) Then
		'        temp = "molecular weight"
		'        NoUnitsOnChangedProperty = False
		'      ElseIf (IndexOfChangedProperty = 6) Then
		'        temp = "Freundlich 1/n"
		'        NoUnitsOnChangedProperty = True
		'      End If
		'      temp = "The property of " & temp & " was manually changed to "
		'      temp = temp & txtDataComponentProperty(IndexOfChangedProperty).Text
		'      If (Not NoUnitsOnChangedProperty) Then
		'        temp = temp & " ("
		'        temp = temp & txtPropUnits(IndexOfChangedProperty).List(txtPropUnits(IndexOfChangedProperty).ListIndex)
		'        temp = temp & ")"
		'      End If
		'      temp = temp & "." & Chr$(13)
		'      temp = temp & "This causes a change in the value of Freundlich K:" & Chr$(13)
		'      temp = temp & "    Old value = " & OldK & " (" & KUnits & ")" & Chr$(13)
		'      temp = temp & "    New values:" & Chr$(13)
		'      For i = 0 To 3
		'        ThisValue = _
		''            Component(0).Use_K * _
		''            KFreundlichConversionFactor(i, Component(0).Use_OneOverN, Component(0).MW)
		'        temp = temp & "        " & _
		''            Format$(ThisValue, NumericalFormat(5)) & _
		''            " (" & txtPropUnits(5).List(i) & ")" & Chr$(13)
		'      Next i
		'      'frmGoAway_Text = temp
		'
		'      'The property of molecular weight was manually changed to 156.77 mg/mmol.
		'      'This causes a change in the value of Freundlich K.
		'      '    Old value = ___________ (unit)
		'      '    New values:
		'      '        ___________ (unit)
		'      '        ___________ (unit)
		'      '        ___________ (unit)
		'      '        ___________ (unit)
		'
		'      CheckINI = ini_getsetting("has_seen_freundlichK_warning")
		'      If (CheckINI = "1") Then
		'        Current_CheckValue = 1
		'      Else
		'        Current_CheckValue = 0
		'      End If
		'      'frmGoAway_CheckText = "Never display this warning again"
		'      If (Current_CheckValue <> 1) Then
		'        Call frmGoAway.frmGoAway_Run( _
		''            frmCompoProp, _
		''            "Warning: Freundlich K Was Updated", _
		''            temp, _
		''            "Never display this warning again", _
		''            Current_CheckValue)
		'      End If
		'      Call ini_putsetting("has_seen_freundlichK_warning", Trim$(Str$(frmGoAway_CheckValue)))
		'      Exit Sub
		'    End If
		'  End If
	End Sub
	Sub Update_cboSource()
		HALT_CBOSOURCE = True
		'POPULATE FREUNDLICH K AND 1/N SOURCE BOX.
		cboSource.Items.Clear()
		If (Component(0).IsothermDB_K > 0#) And (Component(0).IsothermDB_OneOverN > 0#) Then
			'ENABLE ISOTHERM DATABASE AS SOURCE.
			cboSource.Items.Add("Isotherm Database")
		Else
			cboSource.Items.Add("(Isotherm Database)")
		End If
		If (Component(0).IPESResult_K > 0#) And (Component(0).IPESResult_OneOverN > 0#) Then
			'ENABLE IPES AS SOURCE.
			cboSource.Items.Add("Isotherm Parameter Estimation")
		Else
			cboSource.Items.Add("(Isotherm Parameter Estimation)")
		End If
		cboSource.Items.Add("User Entry")
		'DISPLAY CURRENT SOURCE.
		Select Case Component(0).Source_KandOneOverN
			Case KNSOURCE_ISOTHERMDB : cboSource.SelectedIndex = 0
			Case KNSOURCE_IPES : cboSource.SelectedIndex = 1
			Case KNSOURCE_USERINPUT : cboSource.SelectedIndex = 2
		End Select
		HALT_CBOSOURCE = False
	End Sub
	Sub frmCompoProp_PopulateUnits()
		'MAIN BLOCK.
		Call unitsys_register(Me, lblComponentProperty(1), txtDataComponentProperty(1), txtPropUnits(1), "molecular_weight", PropertyUnits.MW, "g/gmol", "", "", 100#, True)
		Call unitsys_register(Me, lblComponentProperty(2), txtDataComponentProperty(2), txtPropUnits(2), "molar_volume", PropertyUnits.MolarVolume, "mL/gmol", "", "", 100#, True)
		Call unitsys_register(Me, lblComponentProperty(3), txtDataComponentProperty(3), txtPropUnits(3), "temperature", PropertyUnits.BP, "C", "", "", 100#, True)
		Call unitsys_register(Me, lblComponentProperty(4), txtDataComponentProperty(4), txtPropUnits(4), "concentration", PropertyUnits.InitialConcentration, "mg/L", "", "", 100#, True)
		Call unitsys_register(Me, lblComponentProperty(10), txtDataComponentProperty(10), txtPropUnits(10), "density", PropertyUnits.Liquid_Density, "g/mL", "", "", 100#, True)
		Call unitsys_register(Me, lblComponentProperty(9), txtDataComponentProperty(9), txtPropUnits(9), "concentration", PropertyUnits.Aqueous_Solubility, "mg/L", "", "", 100#, True)
		Call unitsys_register(Me, lblComponentProperty(7), txtDataComponentProperty(7), txtPropUnits(7), "pressure", PropertyUnits.Vapor_Pressure, "Pa", "", "", 100#, True)
		Call unitsys_register(Me, lblComponentProperty(8), txtDataComponentProperty(8), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblComponentProperty(11), txtDataComponentProperty(11), Nothing, "", "", "", "0", "0", 100.0#, False)
		'FREUNDLICH K AND 1/N BLOCK.
		Call unitsys_register(Me, lblComponentProperty(5), txtDataComponentProperty(5), txtPropUnits(5), "freundlich_k", PropertyUnits.k, "(mg/g)*(L/mg)^(1/n)", "", "", 100#, True)
		Call unitsys_register(Me, lblComponentProperty(6), txtDataComponentProperty(6), Nothing, "", "", "", "", "", 100#, False)
	End Sub
	Sub Store_Unit_Settings()
		PropertyUnits.MW = unitsys_get_units(txtPropUnits(1))
		PropertyUnits.MolarVolume = unitsys_get_units(txtPropUnits(2))
		PropertyUnits.BP = unitsys_get_units(txtPropUnits(3))
		PropertyUnits.InitialConcentration = unitsys_get_units(txtPropUnits(4))
		PropertyUnits.Liquid_Density = unitsys_get_units(txtPropUnits(10))
		PropertyUnits.Aqueous_Solubility = unitsys_get_units(txtPropUnits(9))
		PropertyUnits.Vapor_Pressure = unitsys_get_units(txtPropUnits(7))
		PropertyUnits.k = unitsys_get_units(txtPropUnits(5))
	End Sub
	
	
	'UPGRADE_WARNING: Event cboChemName.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboChemName_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboChemName.SelectedIndexChanged
		If (HALT_CBOCHEMNAME) Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object TempComponents(CurrentCompNumber). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempComponents(CurrentCompNumber) = Component(0)
		CurrentCompNumber = cboChemName.SelectedIndex + 1
		'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Component(0) = TempComponents(CurrentCompNumber)
		Call frmCompoProp_Refresh()
		Call Update_cboSource()
	End Sub
	'UPGRADE_WARNING: Event cboSource.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboSource_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSource.SelectedIndexChanged
		Dim KandOneOverN_Enabled As Short
		Dim X As Short
		Dim temp As String
		If (HALT_CBOSOURCE) Then Exit Sub
		If (VB.Left(VB6.GetItemString(cboSource, cboSource.SelectedIndex), 1) = "(") Then
			'UNABLE TO USE THAT SOURCE!
			X = cboSource.SelectedIndex
			cboSource.SelectedIndex = 2
			Select Case X
				Case 0
					temp = "You must first select an isotherm from the isotherm database.  "
					temp = temp & "Click on the button marked " & Chr(34) & "Freundlich K and 1/n" & Chr(34) & " to do so."
					Call Show_Error(temp)
				Case 1
					temp = "You must first calculate K and 1/n using IPES.  "
					temp = temp & "To do so, click on the button marked " & Chr(34) & "Freundlich K and 1/n" & Chr(34) & " and then click on " & Chr(34) & "Re-calculate" & Chr(34) & " from within the next window."
					Call Show_Error(temp)
			End Select
		End If
		'UPDATE INTERNAL RECORDS.
		Select Case cboSource.SelectedIndex
			Case 0 'ISOTHERM DB.
				Component(0).Source_KandOneOverN = KNSOURCE_ISOTHERMDB
				Component(0).Use_K = Component(0).IsothermDB_K
				Component(0).Use_OneOverN = Component(0).IsothermDB_OneOverN
				KandOneOverN_Enabled = False
			Case 1 'IP ESTIMATION.
				Component(0).Source_KandOneOverN = KNSOURCE_IPES
				Component(0).Use_K = Component(0).IPESResult_K
				Component(0).Use_OneOverN = Component(0).IPESResult_OneOverN
				KandOneOverN_Enabled = False
			Case 2 'USER INPUT.
				Component(0).Source_KandOneOverN = KNSOURCE_USERINPUT
				Component(0).Use_K = Component(0).UserEntered_K
				Component(0).Use_OneOverN = Component(0).UserEntered_OneOverN
				KandOneOverN_Enabled = True
		End Select
		'UPDATE WINDOW.
		Call txtPropUnits_SelectedIndexChanged(txtPropUnits.Item(5), New System.EventArgs())
		'txtDataComponentProperty(5) = Format$(Component(0).Use_K, "0.000")
		'txtDataComponentProperty(6) = Format$(Component(0).Use_OneOverN, "0.000")
		txtDataComponentProperty(5).ReadOnly = Not KandOneOverN_Enabled
		txtDataComponentProperty(6).ReadOnly = Not KandOneOverN_Enabled
		Call frmCompoProp_Refresh() 'THIS CALL REDISPLAYS FREUNDLICH 1/N.
	End Sub


	Private Sub cmdFreundlich_Click()
		Dim Raise_Dirty_Flag As Boolean
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		Call frmFreundlich.frmFreundlich_Run(Raise_Dirty_Flag)
		If (Raise_Dirty_Flag) Then
			'UPDATE SOURCE OF K AND 1/N SCROLLBOX.
			Call Update_cboSource()
			'THROW DIRTY FLAG.
			Call frmCompoProp_DirtyStatus_Throw()
		End If
		'REFRESH VALUES, ESPECIALLY K AND 1/n.
		Call frmCompoProp_Refresh()
	End Sub
	
	
	Private Sub cmdImportClipboard_Click()
		Dim Was_Aborted As Boolean
		Call Do_ImportClipboard(Was_Aborted)
		If (Was_Aborted) Then
			Call Show_Error("There is no valid StEPP data copied to the clipboard.")
			Exit Sub
		Else
			'STORE ALL UNIT SETTINGS.
			Call Store_Unit_Settings()
			'REFRESH MAIN WINDOW.
			Call frmMain_Refresh()
			'EXIT OUT OF HERE.
			USER_HIT_CANCEL = False
			USER_HIT_OK = True
			Me.Close()
			Exit Sub
		End If
	End Sub
	Private Sub cmdImportFromFile_Click()
		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNFileMustExist As Object
		Dim f As Short
		Dim LineCount As Short
		Dim ThisLine As String
		Dim AllLines As String
		Dim InvalidFile As Boolean
		Const MAX_LINE_COUNT As Short = 1000 'SOMEWHAT ARBITRARY.
		On Error GoTo err_cmdImportFromFile_Click
		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'frmMain.CommonDialog1.CancelError = True
		'cancel error appears useless
		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'frmMain.CommonDialog1.DialogTitle = "Load StEPP Export File"

		frmMain.OpenFileDialog1.Title = "Load StEPP Export File"

		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'frmMain.CommonDialog1.Filter = "All Files (*.*)|*.*|StEPP Export Files (*.exp)|*.exp"

		frmMain.OpenFileDialog1.Filter = "All Files (*.*)|*.*|StEPP Export Files (*.exp)|*.exp"

		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'frmMain.CommonDialog1.FilterIndex = 2

		frmMain.OpenFileDialog1.FilterIndex = 2

		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNFileMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'frmMain.CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist

		'Not sure what to replace flags with

		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'frmMain.CommonDialog1.Action = 1

		'Not sure what to replace action with

		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

		frmMain.OpenFileDialog1.ShowDialog()

		If (frmMain.OpenFileDialog1.FileName = "") Then
			Exit Sub
		End If
		f = FreeFile
		LineCount = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.CommonDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(f, frmMain.OpenFileDialog1.FileName, OpenMode.Input)
		InvalidFile = False
		Do While (1 = 1)
			If (EOF(f)) Then Exit Do
			ThisLine = LineInput(f)
			AllLines = AllLines & ThisLine & Chr(13) & Chr(10)
			LineCount = LineCount + 1
			If (LineCount > MAX_LINE_COUNT) Then
				InvalidFile = True
				Exit Do
			End If
		Loop 
		FileClose(f)
		If (InvalidFile) Then
			Call Show_Error("This is not a valid StEPP export file.")
			Exit Sub
		End If
		'DO THE IMPORT.
		My.Computer.Clipboard.SetText(AllLines)
		Call cmdImportClipboard_Click()
		Exit Sub
exit_err_cmdImportFromFile_Click: 
		Exit Sub
err_cmdImportFromFile_Click: 
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = cdlCancel) Then
			'DO NOTHING.
		Else
			Call Show_Trapped_Error("cmdImportFromFile_Click")
		End If
		Resume exit_err_cmdImportFromFile_Click
	End Sub
	
	
	Private Sub cmdKinetics_Click()
		Dim Raise_Dirty_Flag As Boolean
		Call frmKinetic.frmKinetic_Run(Raise_Dirty_Flag)
		If (Raise_Dirty_Flag) Then
			'THROW DIRTY FLAG.
			Call frmCompoProp_DirtyStatus_Throw()
		End If
	End Sub


	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub

	Private Sub frmCompoProp_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

		rs.FindAllControls(Me)

		Dim i As Short
		'MISC INITS.
		'	Me.Height = VB6.TwipsToPixelsY(6885)
		'	Me.Width = VB6.TwipsToPixelsX(8940)
		Call CenterOnForm(Me, frmMain)
		'ADD/EDIT MODE RELATED.
		If (FORM_MODE = FORM_MODE_ADDNEW) Then
			'ADD MODE RELATED INITIALIZATION.
			'---- CREATE DEFAULT COMPONENT.
			'---- NOTE: GLOBAL COMPONENT(0) IS USED FOR THE CURRENT COMPONENT.
			Call SetComponentDefaults(Component(0), 0)
			'---- VARIOUS VISIBILITY SETTINGS.
			cboChemName.Visible = False
			'UPGRADE_WARNING: Couldn't resolve default property of object ssframe_StEPP.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			steppframe.Visible = True
			''---- DATA ALREADY CONSIDERED CHANGED (DUE TO ADD MODE).
			'Call frmCompoProp_DirtyStatus_Throw
			'---- DATA UNCHANGED AS YET.
			Call frmCompoProp_DirtyStatus_Clear()
		Else
			'EDIT MODE RELATED INITIALIZATION.
			'---- TRANSFER COMPONENTS TO LOCAL STORAGE.
			'---- NOTE: GLOBAL COMPONENT(0) IS USED FOR THE CURRENT COMPONENT.
			For i = 1 To Number_Component
				'UPGRADE_WARNING: Couldn't resolve default property of object TempComponents(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TempComponents(i) = Component(i)
			Next i
			'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Component(0) = Component(START_AT_COMPNUMBER)
			CurrentCompNumber = START_AT_COMPNUMBER
			'---- POPULATE COMPONENT NAME SCROLLBOX.
			HALT_CBOCHEMNAME = True
			cboChemName.Items.Clear()
			For i = 1 To Number_Component
				cboChemName.Items.Add(Trim(Component(i).Name))
			Next i
			cboChemName.SelectedIndex = START_AT_COMPNUMBER - 1
			HALT_CBOCHEMNAME = False
			'---- VARIOUS VISIBILITY SETTINGS.
			cboChemName.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object ssframe_StEPP.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			steppframe.Visible = False
			'---- DATA UNCHANGED AS YET.
			Call frmCompoProp_DirtyStatus_Clear()
		End If
		'POPULATE UNIT CONTROLS.
		Call frmCompoProp_PopulateUnits()
		'REFRESH DISPLAY.
		Call frmCompoProp_Refresh()
		'POPULATE SOURCE OF K AND 1/N SCROLLBOX.
		Call Update_cboSource()
		'DEMO SETTINGS.
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	Private Sub frmCompoProp_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub

	Sub frmCompoProp_GenericStatus_Set(ByRef fn_Text As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.StatusLabel.Text = ""
		Me.StatusLabel.Text = fn_Text
	End Sub
	Sub frmCompoProp_DirtyStatus_Set(ByRef newVal As Boolean)
		If (newVal) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object frmCompoProp.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.DirtyLabel.Text = "Data Changed"
			'UPGRADE_WARNING: Couldn't resolve default property of object frmCompoProp.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.DirtyLabel.ForeColor = Color.FromArgb(QBColor(12))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object frmCompoProp.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.DirtyLabel.Text = "Unchanged"
			'UPGRADE_WARNING: Couldn't resolve default property of object frmCompoProp.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.DirtyLabel.ForeColor = Color.FromArgb(QBColor(0))
		End If
	End Sub
	Sub frmCompoProp_DirtyStatus_Set_Current()
		Call frmCompoProp_DirtyStatus_Set(frmCompoProp_Is_Dirty)
	End Sub
	Sub frmCompoProp_DirtyStatus_Throw()
		frmCompoProp_Is_Dirty = True
		Call frmCompoProp_DirtyStatus_Set_Current()
	End Sub
	Sub frmCompoProp_DirtyStatus_Clear()
		frmCompoProp_Is_Dirty = False
		Call frmCompoProp_DirtyStatus_Set_Current()
	End Sub
	
	
	Private Sub txtDataComponentProperty_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDataComponentProperty.Enter
		Dim Index As Short = txtDataComponentProperty.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtDataComponentProperty(Index)
		Dim StatusMessagePanel As String
		If (Index = 0) Then
			Call Global_GotFocus(Ctl)
		Else
			Call unitsys_control_txtx_gotfocus(Ctl)
		End If
		Select Case Index
			Case 0
				StatusMessagePanel = "Type in the component name"
			Case 1
				StatusMessagePanel = "Type in the molecular weight"
			Case 2
				StatusMessagePanel = "Type in the molar volume at the normal boiling point"
			Case 3
				StatusMessagePanel = "Type in the boiling point temperature"
			Case 4
				StatusMessagePanel = "Type in the inlet concentration"
			Case 10
				StatusMessagePanel = "Type in the liquid density"
			Case 9
				StatusMessagePanel = "Type in the aqueous solubility"
			Case 7
				StatusMessagePanel = "Type in the vapor pressure"
			Case 8
				StatusMessagePanel = "Type in the refractive index"
			Case 11
				StatusMessagePanel = "Type in the CAS number, with no hyphen characters"
			Case 5
				StatusMessagePanel = "Type in the Freundlich K value"
			Case 6
				StatusMessagePanel = "Type in the Freundlich 1/n value"
		End Select
		Call frmCompoProp_GenericStatus_Set(StatusMessagePanel)
	End Sub
	Private Sub txtDataComponentProperty_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDataComponentProperty.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtDataComponentProperty.GetIndex(eventSender)
		If (Index = 0) Then
			KeyAscii = Global_TextKeyPress(KeyAscii)
		Else
			KeyAscii = Global_NumericKeyPress(KeyAscii)
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtDataComponentProperty_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDataComponentProperty.Leave
		Dim Index As Short = txtDataComponentProperty.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As TextBox
		Ctl = txtDataComponentProperty(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		'HANDLE THE COMPONENT NAME TEXTBOX.
		If (Index = 0) Then
			If (Trim(Ctl.Text) = "") Then
				Ctl.Text = Component(0).Name
				'Call Show_Error("You must enter a non-blank string for the component name.")
				'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
				'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
			Else
				If (Trim(Component(0).Name) <> Trim(Ctl.Text)) Then
					Component(0).Name = Trim(Ctl.Text)
					'THROW DIRTY FLAG.
					Call frmCompoProp_DirtyStatus_Throw()
				End If
			End If
			Call Global_LostFocus(Ctl)
			Call frmCompoProp_GenericStatus_Set("")
			Exit Sub
		End If
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
		Select Case Index
			Case 1 : Val_Low = 2# : Val_High = 10000000000#
				'Case 2: Val_Low = 0.01 / 1000#: Val_High = 100000# / 1000#
			Case 2 : Val_Low = 0.01 : Val_High = 100000#
			Case 3 : Val_Low = -273# : Val_High = 1000#
			Case 4 : Val_Low = 1E-20 : Val_High = 1000#
			Case 10 : Val_Low = 0.001 : Val_High = 100#
			Case 9 : Val_Low = 0.0001 : Val_High = 10000000#
				''''Case 7: Val_Low = 0.01: Val_High = 1000000#
			Case 7 : Val_Low = 0.0000000001 : Val_High = 1000000#
			Case 8 : Val_Low = 0.01 : Val_High = 100000#
			Case 11 : Val_Low = 0# : Val_High = 2000000000#
			Case 5 : Val_Low = 0.0001 : Val_High = 1000000#
			Case 6 : Val_Low = 0.00001 : Val_High = 10#
		End Select
		NewValue_Okay = False
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		Call frmCompoProp_GenericStatus_Set("")
		If (NewValue_Okay) Then
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				Select Case Index
					Case 1 'MOLECULAR WEIGHT.
						Component(0).MW = NewValue
					Case 2 'MOLAR VOLUME.
						Component(0).MolarVolume = NewValue
					Case 3 'BOILING POINT TEMPERATURE.
						Component(0).BP = NewValue
					Case 4 'INLET CONCENTRATION.
						Component(0).InitialConcentration = NewValue
					Case 10 'LIQUID DENSITY.
						Component(0).Liquid_Density = NewValue
					Case 9 'AQUEOUS SOLUBILITY.
						Component(0).Aqueous_Solubility = NewValue
					Case 7 'VAPOR PRESSURE.
						Component(0).Vapor_Pressure = NewValue
					Case 8 'REFRACTIVE INDEX.
						Component(0).Refractive_Index = NewValue
					Case 11 'CAS NUMBER.
						Component(0).CAS = NewValue
					Case 5 'FREUNDLICH K.
						Component(0).UserEntered_K = NewValue
						Component(0).Use_K = NewValue
					Case 6 'FREUNDLICH 1/N.
						Component(0).UserEntered_OneOverN = NewValue
						Component(0).Use_OneOverN = NewValue
				End Select
				'RAISE DIRTY FLAG IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					'THROW DIRTY FLAG.
					Call frmCompoProp_DirtyStatus_Throw()
					'BASED ON THIS CHANGE, UPDATE FREUNDLICH K IF NECESSARY.
					Call Update_Display_of_KFreundlich(Index)
				End If
				'REFRESH WINDOW.
				Call frmCompoProp_Refresh()
			End If
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event txtPropUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPropUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPropUnits.SelectedIndexChanged
		Dim Index As Short = txtPropUnits.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtPropUnits(Index)
		Call unitsys_control_cbox_click(Ctl)
	End Sub
	Private Sub txtPropUnits_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPropUnits.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtPropUnits.GetIndex(eventSender)
		KeyAscii = Global_TextKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub

	Private Sub _cmdCancelOK_1_Enter(sender As Object, e As EventArgs)
		If (frmCompoProp_Query_Unload() = False) Then
			'THE CANCEL WAS CANCELLED.
			Exit Sub
		End If
		USER_HIT_CANCEL = True
		USER_HIT_OK = False
		Me.Close()
		Exit Sub
	End Sub

	Private Sub _cmdCancelOK_0_Enter(sender As Object, e As EventArgs)
		'UPDATE KINETIC COEFFICIENTS.
		If Component(0).Corr(1) Then
			Component(0).kf = kf(0)
		End If
		If Component(0).Corr(2) Then
			Component(0).Ds = Ds(0)
		End If
		If Component(0).Corr(3) Then
			Component(0).Dp = Dp(0)
		End If
		If (FORM_MODE = FORM_MODE_ADDNEW) Then
			'/////////////////// ADD NEW COMPONENT CODE //////////////////////////////////////////////////////////////////////////////////////
			'ADD COMPONENT TO ROOM PROPERTIES DATA AREA.
			RoomParams.COUNT_CONTAMINANT = RoomParams.COUNT_CONTAMINANT + 1
			RoomParams.ROOM_C0(RoomParams.COUNT_CONTAMINANT) = 0#
			RoomParams.ROOM_EMIT(RoomParams.COUNT_CONTAMINANT) = 1.7
			RoomParams.ROOM_SS_VALUE(RoomParams.COUNT_CONTAMINANT) = 0#
			RoomParams.INITIAL_ROOM_CONC(RoomParams.COUNT_CONTAMINANT) = 0#
			RoomParams.RXN_RATE_CONSTANT(RoomParams.COUNT_CONTAMINANT) = 0#
			RoomParams.RXN_PRODUCT(RoomParams.COUNT_CONTAMINANT) = 0
			RoomParams.RXN_RATIO(RoomParams.COUNT_CONTAMINANT) = 0#

			'ADD COMPONENT TO MAIN DATA AREA.
			Number_Component = Number_Component + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Component(Number_Component). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Component(Number_Component) = Component(0)
			'FRMMAIN.cmdViewDimensionless.Enabled = True
			'FRMMAIN.cmdEditComponent.Enabled = True
			'FRMMAIN.cmdDeleteComponent.Enabled = True
			'FRMMAIN.lstComponents.AddItem txtDataComponentProperty(0)
			'frmMain.cboSelectCompo.Enabled = True
			'frmMain.cboSelectCompo.AddItem txtDataComponentProperty(0)
			'If (Number_Component = Number_Compo_Max) Then
			'  frmMain.cmdAddComponent.Enabled = False
			'End If
			'frmMain.cboSelectCompo.ListIndex = frmMain.cboSelectCompo.ListCount - 1
			''Update the corresponding kinetic data displayed
			'Call Update_Display_Kinetic
		Else
			'/////////////////// EDIT EXISTING COMPONENT(S) CODE //////////////////////////////////////////////////////////////////////////////////////
			'UPGRADE_WARNING: Couldn't resolve default property of object TempComponents(CurrentCompNumber). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TempComponents(CurrentCompNumber) = Component(0)
			Dim i As Integer
			For i = 1 To Number_Component
				'UPGRADE_WARNING: Couldn't resolve default property of object Component(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Component(i) = TempComponents(i)
			Next i
			''Update the display of the name
			'For N = 1 To frmMain.cboSelectCompo.ListCount
			'  frmMain.lstComponents.List(N - 1) = Component(N).name
			'  frmMain.cboSelectCompo.List(N - 1) = Component(N).name
			'Next
			'frmMain.cboSelectCompo.ListIndex = cboChemName.ListIndex
		End If
		'If (Number_Component > 0) Then
		'  frmMain.mnuRunItem(0).Enabled = True
		'  frmMain.mnuRunItem(1).Enabled = True
		'  frmMain.mnuRunItem(2).Enabled = True
		'  frmMain.mnuOptionsItem(0).Enabled = True
		'  frmMain.mnuOptionsItem(1).Enabled = True  'Variable Influent concentration
		'  frmMain.mnuOptionsItem(2).Enabled = True  'Variable Effluent concentration
		'End If
		'STORE ALL UNIT SETTINGS.
		Call Store_Unit_Settings()
		'REFRESH MAIN WINDOW.
		Call frmMain_Refresh()
		'EXIT OUT OF HERE.
		USER_HIT_CANCEL = False
		USER_HIT_OK = True
		Me.Close()
		Exit Sub
	End Sub

	Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
		If (frmCompoProp_Query_Unload() = False) Then
			'THE CANCEL WAS CANCELLED.
			Exit Sub
		End If
		USER_HIT_CANCEL = True
		USER_HIT_OK = False
		Me.Close()
		Exit Sub
	End Sub

	Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
		'UPDATE KINETIC COEFFICIENTS.
		If Component(0).Corr(1) Then
			Component(0).kf = kf(0)
		End If
		If Component(0).Corr(2) Then
			Component(0).Ds = Ds(0)
		End If
		If Component(0).Corr(3) Then
			Component(0).Dp = Dp(0)
		End If
		If (FORM_MODE = FORM_MODE_ADDNEW) Then
			'/////////////////// ADD NEW COMPONENT CODE //////////////////////////////////////////////////////////////////////////////////////
			'ADD COMPONENT TO ROOM PROPERTIES DATA AREA.
			RoomParams.COUNT_CONTAMINANT = RoomParams.COUNT_CONTAMINANT + 1
			RoomParams.ROOM_C0(RoomParams.COUNT_CONTAMINANT) = 0#
			RoomParams.ROOM_EMIT(RoomParams.COUNT_CONTAMINANT) = 1.7
			RoomParams.ROOM_SS_VALUE(RoomParams.COUNT_CONTAMINANT) = 0#
			RoomParams.INITIAL_ROOM_CONC(RoomParams.COUNT_CONTAMINANT) = 0#
			RoomParams.RXN_RATE_CONSTANT(RoomParams.COUNT_CONTAMINANT) = 0#
			RoomParams.RXN_PRODUCT(RoomParams.COUNT_CONTAMINANT) = 0
			RoomParams.RXN_RATIO(RoomParams.COUNT_CONTAMINANT) = 0#

			'ADD COMPONENT TO MAIN DATA AREA.
			Number_Component = Number_Component + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Component(Number_Component). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Component(Number_Component) = Component(0)
			'FRMMAIN.cmdViewDimensionless.Enabled = True
			'FRMMAIN.cmdEditComponent.Enabled = True
			'FRMMAIN.cmdDeleteComponent.Enabled = True
			'FRMMAIN.lstComponents.AddItem txtDataComponentProperty(0)
			'frmMain.cboSelectCompo.Enabled = True
			'frmMain.cboSelectCompo.AddItem txtDataComponentProperty(0)
			'If (Number_Component = Number_Compo_Max) Then
			'  frmMain.cmdAddComponent.Enabled = False
			'End If
			'frmMain.cboSelectCompo.ListIndex = frmMain.cboSelectCompo.ListCount - 1
			''Update the corresponding kinetic data displayed
			'Call Update_Display_Kinetic
		Else
			'/////////////////// EDIT EXISTING COMPONENT(S) CODE //////////////////////////////////////////////////////////////////////////////////////
			'UPGRADE_WARNING: Couldn't resolve default property of object TempComponents(CurrentCompNumber). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TempComponents(CurrentCompNumber) = Component(0)
			Dim i As Integer
			For i = 1 To Number_Component
				'UPGRADE_WARNING: Couldn't resolve default property of object Component(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Component(i) = TempComponents(i)
			Next i
			''Update the display of the name
			'For N = 1 To frmMain.cboSelectCompo.ListCount
			'  frmMain.lstComponents.List(N - 1) = Component(N).name
			'  frmMain.cboSelectCompo.List(N - 1) = Component(N).name
			'Next
			'frmMain.cboSelectCompo.ListIndex = cboChemName.ListIndex
		End If
		'If (Number_Component > 0) Then
		'  frmMain.mnuRunItem(0).Enabled = True
		'  frmMain.mnuRunItem(1).Enabled = True
		'  frmMain.mnuRunItem(2).Enabled = True
		'  frmMain.mnuOptionsItem(0).Enabled = True
		'  frmMain.mnuOptionsItem(1).Enabled = True  'Variable Influent concentration
		'  frmMain.mnuOptionsItem(2).Enabled = True  'Variable Effluent concentration
		'End If
		'STORE ALL UNIT SETTINGS.
		Call Store_Unit_Settings()
		'REFRESH MAIN WINDOW.
		Call frmMain_Refresh()
		'EXIT OUT OF HERE.
		USER_HIT_CANCEL = False
		USER_HIT_OK = True
		Me.Close()
		Exit Sub
	End Sub

	Private Sub cmdKinetics_ClickEvent(sender As Object, e As EventArgs)
		Call cmdKinetics_Click()
	End Sub

	Private Sub cmdFreundlich_ClickEvent(sender As Object, e As EventArgs)
		Call cmdFreundlich_Click()
	End Sub


	Private Sub cmdImportFromFile_ClickEvent(sender As Object, e As EventArgs)
		Call cmdImportFromFile_Click()
	End Sub


	Private Sub Freundlich_Click(sender As Object, e As EventArgs) Handles cmdFreundlich.Click
		Call cmdFreundlich_Click()
	End Sub

	Private Sub Kinetics_Click(sender As Object, e As EventArgs) Handles cmdKinetics.Click
		Call cmdKinetics_Click()
	End Sub

	Private Sub Clipboard_Click(sender As Object, e As EventArgs) Handles cmdImportClipboard.Click
		Call cmdImportClipboard_Click()
	End Sub

	Private Sub ImportFromFile_Click(sender As Object, e As EventArgs) Handles cmdImportFromFile.Click
		Call cmdImportFromFile_Click()
	End Sub

	Private Sub ToolStripStatusLabel1_Click(sender As Object, e As EventArgs) Handles DirtyLabel.Click

	End Sub

	Private Sub _txtPropUnits_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles _txtPropUnits_2.SelectedIndexChanged

	End Sub

	Private Sub frmCompoProp_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)
	End Sub
End Class