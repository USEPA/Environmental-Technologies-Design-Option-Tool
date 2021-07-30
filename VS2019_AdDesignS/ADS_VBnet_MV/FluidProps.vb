Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmFluidProps
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer

	Dim USER_HIT_OK As Boolean
	Dim USER_HIT_CANCEL As Boolean
	
	Dim Save_Density As Double
	Dim Save_Viscosity As Double
	'UPGRADE_WARNING: Lower bound of array Save_State_Check_Water was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim Save_State_Check_Water(2) As Short
	
	
	
	
	Const frmFluidProps_declarations_end As Boolean = True
	
	
	Sub frmFluidProps_Edit(ByRef OUTPUT_Raise_Dirty_Flag As Boolean)
		Me.ShowDialog()
		If (USER_HIT_OK) Then
			OUTPUT_Raise_Dirty_Flag = True
		Else
			OUTPUT_Raise_Dirty_Flag = False
		End If
	End Sub
	Sub LOCAL___Reset_DemoVersionDisablings()
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdCancelOK().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_cmdCancelOK_1.Enabled = False
		End If
	End Sub
	
	
	Sub frmFluidProps_PopulateUnits()
		Call unitsys_register(Me, lblUnit(0), txtWater(0), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblUnit(1), txtWater(1), Nothing, "", "", "", "", "", 100#, False)
	End Sub
	
	
	Private Sub chkCorr_Click(ByRef Index As Short, ByRef Value As Short)
		'UPDATE MEMORY.
		State_Check_Water(Index + 1) = Value
		'IF TURNED CORRELATION ON, RE-CALCULATE DENSITY/VISCOSITY.
		If (Value) Then
			Select Case Index
				Case 0 'DENSITY.
					Call Update_FluidDensity(Bed.Temperature, Bed.Pressure, Bed.WaterDensity)
				Case 1 'VISCOSITY.
					Call Update_FluidViscosity(Bed.Temperature, Bed.WaterViscosity)
			End Select
		End If
		'REFRESH DISPLAY.
		Call frmFluidProps_Refresh()
	End Sub
	
	
	Private Sub cmdCancelOK_Click(ByRef Index As Short)
		Dim i As Short
		Select Case Index
			Case 0 'CANCEL.
				'ROLLBACK TO ORIGINAL VALUES.
				
				'EXIT OUT OF HERE.
				USER_HIT_CANCEL = True
				USER_HIT_OK = False
				Me.Close()
				Exit Sub
			Case 1 'OK.
				'EXIT OUT OF HERE.
				USER_HIT_CANCEL = False
				USER_HIT_OK = True
				Me.Close()
				Exit Sub
		End Select
	End Sub
	
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub
	
	Private Sub frmFluidProps_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		'MISC INITS.
		'   Me.Height = VB6.TwipsToPixelsY(3210)
		'	Me.Width = VB6.TwipsToPixelsX(5235)
		Call CenterOnForm(Me, frmMain)
		If (Bed.Phase = 0) Then
			Me.Text = "Water Properties"
		Else
			Me.Text = "Air Properties"
		End If
		lblUnit(0).Text = "g/cm" & Chr(179)
		lblUnit(1).Text = "g/cm-s"
		
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
		Call frmFluidProps_PopulateUnits()
		'REFRESH DISPLAY.
		Call frmFluidProps_Refresh()
		'DEMO SETTINGS.
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	Private Sub frmFluidProps_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub
	
	
	Private Sub txtWater_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWater.Enter
		Dim Index As Short = txtWater.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtWater(Index)
		Call unitsys_control_txtx_gotfocus(Ctl)
	End Sub
	Private Sub txtWater_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWater.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtWater.GetIndex(eventSender)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtWater_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWater.Leave
		Dim Index As Short = txtWater.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtWater(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
		Select Case Index
			Case 0 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 1 : Val_Low = 1E-20 : Val_High = 1E+20
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
					Case 0 'DENSITY.
						Bed.WaterDensity = NewValue
					Case 1 'VISCOSITY.
						Bed.WaterViscosity = NewValue
				End Select
				'RAISE DIRTY FLAG IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					''THROW DIRTY FLAG.
					'Call frmCompoProp_DirtyStatus_Throw
				End If
				'REFRESH WINDOW.
				Call frmFluidProps_Refresh()
			End If
		End If
	End Sub

	Private Sub _cmdCancelOK_1_ClickEvent(sender As Object, e As EventArgs)
		USER_HIT_CANCEL = True
		USER_HIT_OK = False
		Me.Close()
	End Sub

	Private Sub _cmdCancelOK_0_ClickEvent(sender As Object, e As EventArgs)
		'EXIT OUT OF HERE.
		USER_HIT_CANCEL = False
		USER_HIT_OK = True
		Me.Close()
	End Sub

	Private Sub _chkCorr_0_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)

		State_Check_Water(1) = _chkCorr_0.Checked

		If _chkCorr_0.Checked Then
			'_chkCorr_1.Value = False
			Call Update_FluidDensity(Bed.Temperature, Bed.Pressure, Bed.WaterDensity)
		End If

		'REFRESH DISPLAY.
		Call frmFluidProps_Refresh()
	End Sub

	Private Sub _chkCorr_1_ClickEvent(sender As Object, e As AxThreed.ISSCBCtrlEvents_ClickEvent)
		State_Check_Water(2) = _chkCorr_1.Checked
		If _chkCorr_1.Checked Then
			'_chkCorr_0.Value = False
			Call Update_FluidViscosity(Bed.Temperature, Bed.WaterViscosity)
		End If

		'REFRESH DISPLAY.
		Call frmFluidProps_Refresh()
	End Sub

	Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_1.Click
		USER_HIT_CANCEL = True
		USER_HIT_OK = False
		Me.Close()
	End Sub

	Private Sub OK_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_0.Click
		USER_HIT_CANCEL = False
		USER_HIT_OK = True
		Me.Close()
	End Sub

	Private Sub _chkCorr_0_CheckedChanged(sender As Object, e As EventArgs) Handles _chkCorr_0.CheckedChanged
		State_Check_Water(1) = _chkCorr_0.Checked

		If _chkCorr_0.Checked Then
			'_chkCorr_1.Value = False
			Call Update_FluidDensity(Bed.Temperature, Bed.Pressure, Bed.WaterDensity)
		End If

		'REFRESH DISPLAY.
		Call frmFluidProps_Refresh()
	End Sub

	Private Sub _chkCorr_1_CheckedChanged(sender As Object, e As EventArgs) Handles _chkCorr_1.CheckedChanged
		State_Check_Water(2) = _chkCorr_1.Checked
		If _chkCorr_1.Checked Then
			'_chkCorr_0.Value = False
			Call Update_FluidViscosity(Bed.Temperature, Bed.WaterViscosity)
		End If

		'REFRESH DISPLAY.
		Call frmFluidProps_Refresh()
	End Sub

	Private Sub frmFluidProps_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class