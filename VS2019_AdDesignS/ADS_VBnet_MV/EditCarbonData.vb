Option Strict Off
Option Explicit On
Friend Class frmEditCarbonData
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Dim FORM_MODE As Short
	Const FORM_MODE_ADDNEW As Short = 1
	Const FORM_MODE_EDIT As Short = 2
	
	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_SAVE As Boolean
	Dim USER_HIT_SAVEASNEW As Boolean
	
	Dim DEFAULT_PHASE_IS_LIQUID As Boolean
	
	
	Const frmEditCarbonData_declarations_end As Boolean = True
	
	
	Sub frmEditCarbonData_AddNew(ByRef INPUT_DEFAULT_PHASE_IS_LIQUID As Boolean, ByRef OUTPUT_USER_HIT_CANCEL As Boolean, ByRef OUTPUT_USER_HIT_SAVE As Boolean)
		DEFAULT_PHASE_IS_LIQUID = INPUT_DEFAULT_PHASE_IS_LIQUID
		FORM_MODE = FORM_MODE_ADDNEW
		Me.ShowDialog()
		OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
		OUTPUT_USER_HIT_SAVE = USER_HIT_SAVE
	End Sub
	Sub frmEditCarbonData_Edit(ByRef OUTPUT_USER_HIT_CANCEL As Boolean, ByRef OUTPUT_USER_HIT_SAVE As Boolean, ByRef OUTPUT_USER_HIT_SAVEASNEW As Boolean)
		FORM_MODE = FORM_MODE_EDIT
		Me.ShowDialog()
		OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
		OUTPUT_USER_HIT_SAVE = USER_HIT_SAVE
		OUTPUT_USER_HIT_SAVEASNEW = USER_HIT_SAVEASNEW
	End Sub
	
	
	
	Sub frmEditCarbonData_PopulateUnits()
		Call unitsys_register(Me, lblDesc(0), txtData(0), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(1), txtData(1), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(2), txtData(2), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(4), txtData(4), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(5), txtData(5), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(6), txtData(6), Nothing, "", "", "", "", "", 100#, False)
	End Sub
	
	
	Private Sub cmdSaveAs_Click()
		
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
	
	
	Private Sub frmEditCarbonData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		'MISC INITS.
		'	Me.Height = VB6.TwipsToPixelsY(6045)
		'	Me.Width = VB6.TwipsToPixelsX(4995)
		Call CenterOnForm(Me, frmEditAdsorber)
		lblUnit(0).Text = "g/cm³"
		lblUnit(4).Text = "cm³/g"
		'STRANGE THINGS CAN HAPPEN IF OPTION BOXES ARE ENABLED
		'BEFORE THE FORM IS LOADED/ACTIVATED.
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	optPhase(1).Enabled = True
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	optPhase(2).Enabled = True
		If (FORM_MODE = FORM_MODE_EDIT) Then
			'EDIT MODE.
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		cmdSaveCancel(0).Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		cmdSaveCancel(1).Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		cmdSaveCancel(2).Visible = True
		Else
			'ADD NEW MODE.
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		cmdSaveCancel(0).Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		cmdSaveCancel(1).Visible = False
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		cmdSaveCancel(2).Visible = True
			'SET DEFAULTS FOR THE NEW RECORD.
			frmEditCarbonData_Record.Name = "New Adsorbent"
			frmEditCarbonData_Record.AppDen = 1#
			frmEditCarbonData_Record.ParticleRadius = 0.1
			frmEditCarbonData_Record.ParticlePorosity = 1#
			frmEditCarbonData_Record.AdsType = "GAC"
			frmEditCarbonData_Record.W0 = 0#
			frmEditCarbonData_Record.BB = 0#
			frmEditCarbonData_Record.PolanyiExponent = 0#
			If (DEFAULT_PHASE_IS_LIQUID) Then
				frmEditCarbonData_Record.PhaseIsLiquid = True
			Else
				frmEditCarbonData_Record.PhaseIsLiquid = False
			End If
		End If
		If (frmEditCarbonData_Record.PhaseIsLiquid) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		optPhase(1).Value = True
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		optPhase(2).Value = False
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		optPhase(1).Value = False
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		optPhase(2).Value = True
		End If
		'POPULATE UNIT CONTROLS.
		Call frmEditCarbonData_PopulateUnits()
		'REFRESH DISPLAY.
		Call frmEditCarbonData_Refresh()
	End Sub
	Private Sub frmEditCarbonData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub
	
	
	
	
	
	
	
	
	
	Private Sub optPhase_Click(ByRef Index As Short, ByRef Value As Short)
		If (Index = 1) Then
			frmEditCarbonData_Record.PhaseIsLiquid = True
		Else
			frmEditCarbonData_Record.PhaseIsLiquid = False
		End If
	End Sub
	
	
	Private Sub txtData_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Enter
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim Ctl As TextBox
		Ctl = txtData(Index)
		If (Index = 7) Or (Index = 3) Then
			Call Global_GotFocus(Ctl)
		Else
			Call unitsys_control_txtx_gotfocus(Ctl)
		End If
	End Sub
	Private Sub txtData_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtData.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtData.GetIndex(eventSender)
		If (Index = 7) Or (Index = 3) Then
			KeyAscii = Global_TextKeyPress(KeyAscii)
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
		If (Index = 7) Or (Index = 3) Then
			Select Case Index
				Case 7 : OldValueStr = Trim(frmEditCarbonData_Record.Name)
				Case 3 : OldValueStr = Trim(frmEditCarbonData_Record.AdsType)
			End Select
			If (Trim(Ctl.Text) = "") Then
				Ctl.Text = OldValueStr
				'Call Show_Error("You must enter a non-blank string for the carbon name.")
				'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
				'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
			Else
				If (Trim(OldValueStr) <> Trim(Ctl.Text)) Then
					Select Case Index
						Case 7 : frmEditCarbonData_Record.Name = Trim(Ctl.Text)
						Case 3 : frmEditCarbonData_Record.AdsType = Trim(Ctl.Text)
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
			Case 1 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 2 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 4 : Val_Low = 0# : Val_High = 1E+20
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
					Case 0 'APPARENT DENSITY.
						frmEditCarbonData_Record.AppDen = NewValue
					Case 1 'PARTICLE RADIUS.
						frmEditCarbonData_Record.ParticleRadius = NewValue
					Case 2 'PARTICLE POROSITY.
						frmEditCarbonData_Record.ParticlePorosity = NewValue
					Case 4 'POLANYI W0.
						frmEditCarbonData_Record.W0 = NewValue
					Case 5 'POLANYI BB.
						frmEditCarbonData_Record.BB = NewValue
					Case 6 'POLANYI EXPONENT.
						frmEditCarbonData_Record.PolanyiExponent = NewValue
				End Select
				'RAISE DIRTY FLAG IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					''THROW DIRTY FLAG.
					'Call frmCompoProp_DirtyStatus_Throw
				End If
				'REFRESH WINDOW.
				Call frmEditAdsorberData_Refresh()
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


	Private Sub _optPhase_1_CheckedChanged(sender As Object, e As EventArgs) Handles _optPhase_1.CheckedChanged
		Call optPhase_Click(1, 0)
	End Sub

	Private Sub _optPhase_2_CheckedChanged(sender As Object, e As EventArgs) Handles _optPhase_2.CheckedChanged
		Call optPhase_Click(2, 0)
	End Sub

	Private Sub Save_Click(sender As Object, e As EventArgs) Handles _cmdSaveCancel_0.Click
		Call cmdSaveCancel_Click(0)
	End Sub

	Private Sub SaveNew_Click(sender As Object, e As EventArgs) Handles _cmdSaveCancel_1.Click
		Call cmdSaveCancel_Click(1)
	End Sub

	Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles _cmdSaveCancel_2.Click
		Call cmdSaveCancel_Click(2)
	End Sub

	Private Sub frmEditCarbonData_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)
	End Sub
End Class