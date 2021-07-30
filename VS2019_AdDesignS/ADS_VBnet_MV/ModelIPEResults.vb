Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmModelIPEResults
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer

	Dim WHICH_MODEL As Short
	
	
	
	
	Const frmModelIPEResults_declarations_end As Boolean = True
	
	
	Sub frmModelIPEResults_Run(ByRef INPUT_WHICH_MODEL As Short)
		WHICH_MODEL = INPUT_WHICH_MODEL
		Me.ShowDialog()
	End Sub
	
	
	Sub frmModelIPEResults_PopulateUnits()
		Dim i As Short
		For i = 0 To 12
			If (i <> 3) Then
				Call unitsys_register(Me, lblDesc(i), txtData(i), Nothing, "", "", "", "", "", 100#, False)
			End If
		Next i
	End Sub
	Sub frmModelIPEResults_Refresh()
		Dim Frm As frmModelIPEResults
		Frm = Me
		Dim i As Short
		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(0), IPES_Data.Input_Renamed.W0)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(1), IPES_Data.Input_Renamed.BB)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(2), IPES_Data.Input_Renamed.GM)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(4), IPES_Data.Output.CSAV)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(5), IPES_Data.Output.QSAV)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(6), IPES_Data.Output.XK1)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(7), IPES_Data.Output.XK2)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(8), IPES_Data.Output.XN)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(9), IPES_Data.Output.CBEG)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(10), IPES_Data.Output.CEND)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(11), IPES_Data.Output.RSQD)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(12), IPES_Data.Output.RMSE)
		'
		' REMOVE 9-12 IF NO DATA AVAILABLE.
		'
		If ((IPES_Data.Output.CBEG = 0#) And (IPES_Data.Output.CEND = 0#) And (IPES_Data.Output.RSQD = 0#) And (IPES_Data.Output.RMSE = 0#)) Then
			For i = 9 To 12
				'UPGRADE_ISSUE: Control lblDesc could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm.lblDesc(i).Visible = False
				'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm.txtData(i).Visible = False
			Next i
		End If
	End Sub
	
	
	Private Sub cmdClose_Click()
		Me.Close()
		Exit Sub
	End Sub
	
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub
	
	Private Sub frmModelIPEResults_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'MISC INITS.
		rs.FindAllControls(Me)

		Call CenterOnForm(Me, frmFreundlich)
		Select Case WHICH_MODEL
			Case MODULECODE_ADLIQ
				txtData(3).Text = "3-Parameter Polanyi Correlation"
			Case MODULECODE_SPEQ
				txtData(3).Text = "D-R Equal Spreading Pressure Calculation"
			Case MODULECODE_HOFMAN
				txtData(3).Text = "Estimated From Gas-Phase D-R Isotherm" & vbCrLf & "Hansen-Fackler model, uniform adsorbate"
			Case Else
		End Select
		'POPULATE UNIT CONTROLS.
		Call frmModelIPEResults_PopulateUnits()
		'REFRESH DISPLAY.
		Call frmModelIPEResults_Refresh()
	End Sub
	Private Sub frmModelIPEResults_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub
	
	
	Private Sub txtData_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Enter
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtData(Index)
		If (Index = 3) Then
			Call Global_GotFocus(Ctl)
			Exit Sub
		End If
		Call unitsys_control_txtx_gotfocus(Ctl)
	End Sub
	'Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
	'  KeyAscii = Global_NumericKeyPress(KeyAscii)
	'End Sub
	Private Sub txtData_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Leave
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtData(Index)
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
			Case 0 : Val_Low = 0.05 : Val_High = 2.5
			Case 1 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 2 : Val_Low = 1E-20 : Val_High = 1E+20
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

	Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
		Call cmdClose_Click()
	End Sub

	Private Sub _lblDesc_4_Click(sender As Object, e As EventArgs) Handles _lblDesc_4.Click

	End Sub

	Private Sub _txtData_4_TextChanged(sender As Object, e As EventArgs) Handles _txtData_4.TextChanged

	End Sub

	Private Sub _lblDesc_5_Click(sender As Object, e As EventArgs) Handles _lblDesc_5.Click

	End Sub

	Private Sub _txtData_5_TextChanged(sender As Object, e As EventArgs) Handles _txtData_5.TextChanged

	End Sub

	Private Sub _lblDesc_6_Click(sender As Object, e As EventArgs) Handles _lblDesc_6.Click

	End Sub

	Private Sub _txtData_6_TextChanged(sender As Object, e As EventArgs) Handles _txtData_6.TextChanged

	End Sub

	Private Sub _lblDesc_7_Click(sender As Object, e As EventArgs) Handles _lblDesc_7.Click

	End Sub

	Private Sub _txtData_7_TextChanged(sender As Object, e As EventArgs) Handles _txtData_7.TextChanged

	End Sub

	Private Sub _lblDesc_8_Click(sender As Object, e As EventArgs) Handles _lblDesc_8.Click

	End Sub

	Private Sub _txtData_8_TextChanged(sender As Object, e As EventArgs) Handles _txtData_8.TextChanged

	End Sub

	Private Sub _lblDesc_9_Click(sender As Object, e As EventArgs) Handles _lblDesc_9.Click

	End Sub

	Private Sub _txtData_9_TextChanged(sender As Object, e As EventArgs) Handles _txtData_9.TextChanged

	End Sub

	Private Sub _lblDesc_10_Click(sender As Object, e As EventArgs) Handles _lblDesc_10.Click

	End Sub

	Private Sub _txtData_10_TextChanged(sender As Object, e As EventArgs) Handles _txtData_10.TextChanged

	End Sub

	Private Sub _lblDesc_11_Click(sender As Object, e As EventArgs) Handles _lblDesc_11.Click

	End Sub

	Private Sub _txtData_11_TextChanged(sender As Object, e As EventArgs) Handles _txtData_11.TextChanged

	End Sub

	Private Sub _lblDesc_12_Click(sender As Object, e As EventArgs) Handles _lblDesc_12.Click

	End Sub

	Private Sub _txtData_12_TextChanged(sender As Object, e As EventArgs) Handles _txtData_12.TextChanged

	End Sub

	Private Sub frmModelIPEResults_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class