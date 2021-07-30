Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmEditAdsorberData
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	'Dim frmEditAdsorber_Cancelled As Integer
	'Dim frmEditAdsorber_RunMode As Integer
	'Const frmEditAdsorber_RunMode_QUERY_DATABASE = 1
	'Const frmEditAdsorber_RunMode_EDIT_DATABASE = 2

	'Dim frmEditAdsorberData_Cancelled As Integer
	Dim frmEditAdsorberData_RunMode As Short
	Const frmEditAdsorberData_RunMode_NEW As Short = 1
	Const frmEditAdsorberData_RunMode_EDIT As Short = 2
	Dim frmEditAdsorberData_UsePhase As Short
	
	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_SAVE As Boolean
	
	
	
	
	Const frmEditAdsorberData_declarations_end As Boolean = True
	
	
	Sub frmEditAdsorberData_AddNew(ByRef INPUT_PHASE As Short, ByRef OUTPUT_USER_HIT_CANCEL As Boolean)
		frmEditAdsorberData_RunMode = frmEditAdsorberData_RunMode_NEW
		frmEditAdsorberData_UsePhase = INPUT_PHASE
		Me.ShowDialog()

		Dim now_phase As Short = frmEditAdsorberData_UsePhase
		If (USER_HIT_CANCEL) Then
			OUTPUT_USER_HIT_CANCEL = True
		Else
			OUTPUT_USER_HIT_CANCEL = False
		End If
	End Sub
	Sub frmEditAdsorberData_Edit(ByRef INPUT_PHASE As Short, ByRef OUTPUT_USER_HIT_CANCEL As Boolean)
		frmEditAdsorberData_RunMode = frmEditAdsorberData_RunMode_EDIT
		frmEditAdsorberData_UsePhase = INPUT_PHASE
		Me.ShowDialog()

		If (USER_HIT_CANCEL) Then
			OUTPUT_USER_HIT_CANCEL = True
		Else
			OUTPUT_USER_HIT_CANCEL = False
		End If
	End Sub
	
	
	Sub frmEditAdsorberData_PopulateUnits()
		Call unitsys_register(Me, lblDesc(0), txtData(0), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(1), txtData(1), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(2), txtData(2), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblDesc(5), txtData(5), Nothing, "", "", "", "", "", 100#, False)
	End Sub
	
	
	Private Sub cmdCancel_Click()
		'frmEditAdsorberData_Cancelled = True
		USER_HIT_CANCEL = True
		USER_HIT_SAVE = False
		frmEditAdsorber.lstManu.SelectedIndex = 0
		frmEditAdsorber.lstName.SelectedIndex = 0
		Me.Dispose()
	End Sub
	Private Sub cmdSave_Click()
		Dim i As Short
		For i = 0 To 7
			If (Trim(txtData(i).Text) = "") Then
				Beep()
				MsgBox("No data item can be set to an empty string; enter a non-empty string or hit Cancel.", MsgBoxStyle.Exclamation, AppName_For_Display_Long)
				Exit Sub
			End If
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_optphase_1.Checked) Then
			frmEditAdsorberData_Record.Phase = 1
		Else
			frmEditAdsorberData_Record.Phase = 2
		End If
		frmEditAdsorberData_Record.InternalArea = txtData(0).Text
		frmEditAdsorberData_Record.MaxCapacity = txtData(1).Text
		frmEditAdsorberData_Record.OutsideDiameter = txtData(2).Text
		frmEditAdsorberData_Record.DesignPressure = txtData(3).Text
		frmEditAdsorberData_Record.DesignFlowRange = txtData(4).Text
		frmEditAdsorberData_Record.DefaultFlowRate = txtData(5).Text
		frmEditAdsorberData_Record.Note = txtData(6).Text
		frmEditAdsorberData_Record.PartNumber = txtData(7).Text
		'frmEditAdsorberData_Cancelled = False
		USER_HIT_CANCEL = False
		USER_HIT_SAVE = True

		frmEditAdsorber.lstManu.SelectedIndex = 0
		frmEditAdsorber.lstName.SelectedIndex = 0
		Me.Dispose()
	End Sub
	
	
	Private Sub Command4_Click()
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub
	
	Private Sub frmEditAdsorberData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

		rs.FindAllControls(Me)

		Dim now_phase As Short
		Dim i As Short
		'MISC INITS.
		'	Me.Height = VB6.TwipsToPixelsY(5505)
		'		Me.Width = VB6.TwipsToPixelsX(5205)
		Call CenterOnForm(Me, frmEditAdsorber)

		now_phase = frmEditAdsorberData_UsePhase
		If (now_phase = 1) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'optPhase(1).Value = True
			_optphase_1.Checked = True

			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'optPhase(2).Value = False
			_optphase_2.Checked = False

		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_1.Checked = False
			'UPGRADE_WARNING: Couldn't resolve default property of object optPhase().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_optphase_2.Checked = True
		End If
		If (frmEditAdsorberData_RunMode = frmEditAdsorberData_RunMode_NEW) Then
			'CREATE NEW RECORD
			frmEditAdsorberData_Record.PartNumber = "New Adsorber Code"
			frmEditAdsorberData_Record.InternalArea = "1"
			frmEditAdsorberData_Record.MaxCapacity = "1000"
			frmEditAdsorberData_Record.OutsideDiameter = "1"
			frmEditAdsorberData_Record.DesignPressure = "not available"
			frmEditAdsorberData_Record.DesignFlowRange = "1-10"
			frmEditAdsorberData_Record.DefaultFlowRate = "10"
			frmEditAdsorberData_Record.Note = "none"
		End If
		If (frmEditAdsorberData_RunMode = frmEditAdsorberData_RunMode_EDIT) Then
			'MODIFY EXISTING RECORD
			'    txtData(0) = Trim$(frmEditAdsorberData_Record.InternalArea)
			'    txtData(1) = Trim$(frmEditAdsorberData_Record.MaxCapacity)
			'    txtData(2) = Trim$(frmEditAdsorberData_Record.OutsideDiameter)
			'    txtData(3) = Trim$(frmEditAdsorberData_Record.DesignPressure)
			'    txtData(4) = Trim$(frmEditAdsorberData_Record.DesignFlowRange)
			'    txtData(5) = Trim$(frmEditAdsorberData_Record.DefaultFlowRate)
			'    txtData(6) = Trim$(frmEditAdsorberData_Record.Note)
			'    txtData(7) = Trim$(frmEditAdsorberData_Record.PartNumber)
		End If
		'POPULATE UNIT CONTROLS.
		Call frmEditAdsorberData_PopulateUnits()
		'REFRESH DISPLAY.
		Call frmEditAdsorberData_Refresh()

	End Sub
	Private Sub frmEditAdsorberData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub


	Private Sub optPhase_Click(ByRef optIndex As Short, ByRef Value As Short)
		'UPGRADE_WARNING: Couldn't resolve default property of object optPhase(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

		If (_optphase_1.Checked And optIndex = 1) Then
			'LIQUID PHASE
			_optphase_2.Checked = False
			lblUnits(4).Text = "(gal/min)"
			lblUnits(5).Text = "(gal/min)"
		ElseIf (_optphase_2.Checked And optIndex = 2) Then
			'GAS PHASE
			_optphase_1.Checked = False
			lblUnits(4).Text = "(ft³/min)"
			lblUnits(5).Text = "(ft³/min)"
		ElseIf (_optphase_1.Checked = False And optIndex = 1) Then
			'GAS PHASE
			_optphase_2.Checked = False
			lblUnits(4).Text = "(ft³/min)"
			lblUnits(5).Text = "(ft³/min)"
		ElseIf (_optphase_2.Checked = False And optIndex = 2) Then
			'Liquid PHASE
			_optphase_1.Checked = True
			lblUnits(4).Text = "(gal/min)"
			lblUnits(5).Text = "(gal/min)"
		End If
	End Sub






	Private Sub txtData_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Enter
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtData(Index)
		If (Index = 7) Or (Index = 3) Or (Index = 4) Or (Index = 6) Then
			Call Global_GotFocus(Ctl)
		Else
			Call unitsys_control_txtx_gotfocus(Ctl)
		End If
	End Sub
	Private Sub txtData_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtData.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtData.GetIndex(eventSender)
		If (Index = 7) Or (Index = 3) Or (Index = 4) Or (Index = 6) Then
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
		If (Index = 7) Or (Index = 3) Or (Index = 4) Or (Index = 6) Then
			Select Case Index
				Case 7 : OldValueStr = Trim(frmEditAdsorberData_Record.PartNumber)
				Case 3 : OldValueStr = Trim(frmEditAdsorberData_Record.DesignPressure)
				Case 4 : OldValueStr = Trim(frmEditAdsorberData_Record.DesignFlowRange)
				Case 6 : OldValueStr = Trim(frmEditAdsorberData_Record.Note)
			End Select
			If (Trim(Ctl.Text) = "") Then
				Ctl.Text = OldValueStr
				'Call Show_Error("You must enter a non-blank string for the carbon name.")
				'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
				'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
			Else
				If (Trim(OldValueStr) <> Trim(Ctl.Text)) Then
					Select Case Index
						Case 7 : frmEditAdsorberData_Record.PartNumber = Trim(Ctl.Text)
						Case 3 : frmEditAdsorberData_Record.DesignPressure = Trim(Ctl.Text)
						Case 4 : frmEditAdsorberData_Record.DesignFlowRange = Trim(Ctl.Text)
						Case 6 : frmEditAdsorberData_Record.Note = Trim(Ctl.Text)
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
			Case 5 : Val_Low = 1E-20 : Val_High = 1E+20
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
					Case 0 'INTERNAL AREA.
						frmEditAdsorberData_Record.InternalArea = Trim(Str(NewValue))
					Case 1 'MAXIMUM CAPACITY.
						frmEditAdsorberData_Record.MaxCapacity = Trim(Str(NewValue))
					Case 2 'OUTSIDE DIAMETER.
						frmEditAdsorberData_Record.OutsideDiameter = Trim(Str(NewValue))
					Case 5 'DEFAULT FLOW RATE.
						frmEditAdsorberData_Record.DefaultFlowRate = Trim(Str(NewValue))
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

	Private Sub _optphase_1_CheckedChanged(sender As Object, e As EventArgs) Handles _optphase_1.CheckedChanged
		Call optPhase_Click(1, 0)
	End Sub

	Private Sub _optphase_2_CheckedChanged(sender As Object, e As EventArgs) Handles _optphase_2.CheckedChanged
		Call optPhase_Click(2, 0)
	End Sub

	Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
		Call cmdCancel_Click()
	End Sub

	Private Sub Save_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
		Call cmdSave_Click()
	End Sub

	Private Sub frmEditAdsorberData_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)
	End Sub
End Class