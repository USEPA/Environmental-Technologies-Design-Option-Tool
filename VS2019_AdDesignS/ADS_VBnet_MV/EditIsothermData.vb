Option Strict Off
Option Explicit On
Friend Class frmEditIsothermData
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer


	Dim FORM_MODE As Short
	Const FORM_MODE_ADDNEW As Short = 1
	Const FORM_MODE_EDIT As Short = 2
	
	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_SAVE As Boolean
	Dim USER_HIT_SAVEASNEW As Boolean
	
	Dim DEFAULT_PHASE_IS_LIQUID As Boolean
	Dim DEFAULT_CHEMICALNAME As String
	Dim DEFAULT_CHEMICALCAS As String
	
	
	
	Const frmEditIsothermData_declarations_end As Boolean = True
	
	
	Sub frmEditIsothermData_AddNew(ByRef INPUT_DEFAULT_PHASE_IS_LIQUID As Boolean, ByRef INPUT_DEFAULT_CHEMICALNAME As String, ByRef INPUT_DEFAULT_CHEMICALCAS As String, ByRef OUTPUT_USER_HIT_CANCEL As Boolean, ByRef OUTPUT_USER_HIT_SAVE As Boolean)
		DEFAULT_PHASE_IS_LIQUID = INPUT_DEFAULT_PHASE_IS_LIQUID
		DEFAULT_CHEMICALNAME = INPUT_DEFAULT_CHEMICALNAME
		DEFAULT_CHEMICALCAS = INPUT_DEFAULT_CHEMICALCAS
		FORM_MODE = FORM_MODE_ADDNEW
		Me.ShowDialog()
		OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
		OUTPUT_USER_HIT_SAVE = USER_HIT_SAVE
	End Sub
	Sub frmEditIsothermData_Edit(ByRef OUTPUT_USER_HIT_CANCEL As Boolean, ByRef OUTPUT_USER_HIT_SAVE As Boolean, ByRef OUTPUT_USER_HIT_SAVEASNEW As Boolean)
		FORM_MODE = FORM_MODE_EDIT
		Me.ShowDialog()
		OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
		OUTPUT_USER_HIT_SAVE = USER_HIT_SAVE
		OUTPUT_USER_HIT_SAVEASNEW = USER_HIT_SAVEASNEW
	End Sub
	
	
	
	Sub frmEditIsothermData_PopulateUnits()
		Call unitsys_register(Me, lblData(0), txtData(0), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblData(1), txtData(1), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblData(2), txtData(2), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblData(3), txtData(3), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblData(4), txtData(4), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblData(5), txtData(5), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblData(6), txtData(6), Nothing, "", "", "", "", "", 100#, False)
	End Sub
	
	
	Private Sub cmdSaveCancel_Click(ByRef Index As Short)
		Select Case Index
			Case 0 'SAVE.
				USER_HIT_CANCEL = False
				USER_HIT_SAVE = True
				USER_HIT_SAVEASNEW = False
				Me.Close()
				Exit Sub
			Case 1 'SAVE AS NEW RECORD.
				USER_HIT_CANCEL = False
				USER_HIT_SAVE = False
				USER_HIT_SAVEASNEW = True
				Me.Close()
				Exit Sub
			Case 2 'CANCEL.
				USER_HIT_CANCEL = True
				USER_HIT_SAVE = False
				USER_HIT_SAVEASNEW = False
				Me.Close()
				Exit Sub
		End Select
	End Sub
	
	
	Private Sub frmEditIsothermData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		'MISC INITS.
		'	Me.Height = VB6.TwipsToPixelsY(7035)
		'	Me.Width = VB6.TwipsToPixelsX(7935)
		Call CenterOnForm(Me, frmEditIsotherm)
		'STRANGE THINGS CAN HAPPEN IF OPTION BOXES ARE ENABLED
		'BEFORE THE FORM IS LOADED/ACTIVATED.
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_optPhase_1.Enabled = True
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_optPhase_2.Enabled = True
		If (FORM_MODE = FORM_MODE_EDIT) Then
			'EDIT MODE.
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdSaveCancel_0.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdSaveCancel_1.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdSaveCancel_2.Visible = True
		Else
			'ADD NEW MODE.
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdSaveCancel_0.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdSaveCancel_1.Visible = False
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdSaveCancel_2.Visible = True
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
			frmEditIsothermData_Record.Tmin = CStr(25#)
			frmEditIsothermData_Record.Comments = ""
		End If
		If (frmEditIsothermData_Record.PhaseIsLiquid) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_1.Checked = True
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_2.Checked = False
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_1.Checked = False
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_2.Checked = True
		End If
		'POPULATE UNIT CONTROLS.
		Call frmEditIsothermData_PopulateUnits()
		'REFRESH DISPLAY.
		Call frmEditIsothermData_Refresh()
	End Sub
	Private Sub frmEditIsothermData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub
	
	
	
	
	
	
	
	
	
	Private Sub optPhase_Click(ByRef Index As Short, ByRef Value As Short)
		If (Index = 1) Then
			frmEditIsothermData_Record.PhaseIsLiquid = True
		Else
			frmEditIsothermData_Record.PhaseIsLiquid = False
		End If
	End Sub
	
	
	Private Sub txtData_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Enter
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtData(Index)
		If (Index >= 7) And (Index <= 11) Then
			Call Global_GotFocus(Ctl)
		Else
			Call unitsys_control_txtx_gotfocus(Ctl)
		End If
	End Sub
	Private Sub txtData_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtData.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtData.GetIndex(eventSender)
		If (Index >= 7) And (Index <= 11) Then
			If (Index = 8) Then
				KeyAscii = Global_Numeric0123456789KeyPress(KeyAscii)
			Else
				KeyAscii = Global_TextKeyPress(KeyAscii)
			End If
		Else
			KeyAscii = Global_NumericKeyPress(KeyAscii)
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtData_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Leave
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtData(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		Dim OldValueStr As String
		'HANDLE STRING FIELDS.
		If (Index >= 7) And (Index <= 11) Then
			Select Case Index
				Case 7 : OldValueStr = Trim(frmEditIsothermData_Record.CarbonName)
				Case 8 : OldValueStr = Trim(frmEditIsothermData_Record.CAS)
				Case 9 : OldValueStr = Trim(frmEditIsothermData_Record.Name)
				Case 10 : OldValueStr = Trim(frmEditIsothermData_Record.Source)
				Case 11 : OldValueStr = Trim(frmEditIsothermData_Record.Comments)
			End Select
			'NOTE: ZERO-LENGTH STRINGS FOR 8 AND 11 ARE ALLOWED.
			If (Trim(Ctl.Text) = "") And (Index <> 8) And (Index <> 11) Then
				Ctl.Text = OldValueStr
				'Call Show_Error("You must enter a non-blank string for the carbon name.")
				'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
				'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
			Else
				If (Trim(OldValueStr) <> Trim(Ctl.Text)) Then
					Select Case Index
						Case 7 : frmEditIsothermData_Record.CarbonName = Trim(Ctl.Text)
						Case 8 : frmEditIsothermData_Record.CAS = Trim(Ctl.Text)
						Case 9 : frmEditIsothermData_Record.Name = Trim(Ctl.Text)
						Case 10 : frmEditIsothermData_Record.Source = Trim(Ctl.Text)
						Case 11 : frmEditIsothermData_Record.Comments = Trim(Ctl.Text)
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
			Case 0 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 1 : Val_Low = 0# : Val_High = 1E+20
			Case 2 : Val_Low = 0# : Val_High = 1E+20
			Case 3 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 4 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 5 : Val_Low = 0# : Val_High = 1E+20
			Case 6 : Val_Low = 0# : Val_High = 1E+20
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
					Case 0 'FREUNDLICH K.
						frmEditIsothermData_Record.k = NewValue
					Case 1 'MINIMUM CONCENTRATION.
						frmEditIsothermData_Record.Cmin = NewValue
					Case 2 'MINIMUM pH.
						frmEditIsothermData_Record.pHmin = NewValue
					Case 3 'TEMPERATURE.
						frmEditIsothermData_Record.Tmin = CStr(NewValue)
					Case 4 'FREUNDLICH 1/n.
						frmEditIsothermData_Record.OneOverN = NewValue
					Case 5 'MAXIMUM CONCENTRATION.
						frmEditIsothermData_Record.Cmax = NewValue
					Case 6 'MAXIMUM pH.
						frmEditIsothermData_Record.pHmax = NewValue
				End Select
				'RAISE DIRTY FLAG IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					''THROW DIRTY FLAG.
					'Call frmCompoProp_DirtyStatus_Throw
				End If
				'REFRESH WINDOW.
				Call frmEditIsothermData_Refresh()
			End If
		End If
	End Sub

	Private Sub _cmdSaveCancel_0_ClickEvent(sender As Object, e As EventArgs)
		Call cmdSaveCancel_Click(0)
	End Sub

	Private Sub _cmdSaveCancel_1_ClickEvent(sender As Object, e As EventArgs)
		Call cmdSaveCancel_Click(1)
	End Sub

	Private Sub _cmdSaveCancel_2_ClickEvent(sender As Object, e As EventArgs)
		Call cmdSaveCancel_Click(2)
	End Sub

	Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles _cmdSaveCancel_2.Click
		Call cmdSaveCancel_Click(2)
	End Sub

	Private Sub SaveNew_Click(sender As Object, e As EventArgs) Handles _cmdSaveCancel_1.Click
		Call cmdSaveCancel_Click(1)
	End Sub

	Private Sub Save_Click(sender As Object, e As EventArgs) Handles _cmdSaveCancel_0.Click
		Call cmdSaveCancel_Click(0)
	End Sub

	Private Sub frmEditIsothermData_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class