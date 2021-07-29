Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmKinetic
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_OK As Boolean
	Dim frmKinetic_Is_Dirty As Boolean
	
	'UPGRADE_WARNING: Arrays in structure SaveOldComponent may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim SaveOldComponent As ComponentPropertyType
	
	
	
	
	Const frmKinetic_declarations_end As Boolean = True
	
	
	Sub frmKinetic_Run(ByRef OUTPUT_Raise_Dirty_Flag As Boolean)
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
	
	
	Sub frmKinetic_GenericStatus_Set(ByRef fn_Text As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object Me.sspanel_Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.ToolStripLabelStatus.Text = fn_Text
	End Sub
	Sub frmKinetic_DirtyStatus_Set(ByRef newVal As Boolean)
		If (newVal) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object frmKinetic.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.ToolStripLabelDirty.Text = "Data Changed"
			'UPGRADE_WARNING: Couldn't resolve default property of object frmKinetic.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.ToolStripLabelDirty.ForeColor = Color.FromArgb(QBColor(12))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object frmKinetic.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.ToolStripLabelDirty.Text = "Unchanged"
			'UPGRADE_WARNING: Couldn't resolve default property of object frmKinetic.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.ToolStripLabelDirty.ForeColor = Color.FromArgb(QBColor(0))
		End If
	End Sub
	Sub frmKinetic_DirtyStatus_Set_Current()
		Call frmKinetic_DirtyStatus_Set(frmKinetic_Is_Dirty)
	End Sub
	Sub frmKinetic_DirtyStatus_Throw()
		frmKinetic_Is_Dirty = True
		Call frmKinetic_DirtyStatus_Set_Current()
	End Sub
	Sub frmKinetic_DirtyStatus_Clear()
		frmKinetic_Is_Dirty = False
		Call frmKinetic_DirtyStatus_Set_Current()
	End Sub
	
	
	Sub frmKinetic_PopulateUnits()
		Call unitsys_register(Me, lblSPDFR, txtKF, Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblSPDFR, txtDS, Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblSPDFR, txtDP, Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblSPDFR, txtSPDFR, Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblTort, txtTort, Nothing, "", "", "", "", "", 100#, False)
	End Sub
	
	
	Sub Update_Ds_and_Dp_Editability()
		If (Component(0).Use_Tortuosity_Correlation) Then
			'---- Tortuosity(time) correlation is ON!
			txtTort.Enabled = False
			'---- De-enable Ds user-input
			optDS(0).Enabled = False
			optDS(0).Checked = False
			optDS(1).Checked = True
			Component(0).Corr(2) = True
			'---- De-enable Dp user-input
			optDP(0).Enabled = False
			optDP(0).Checked = False
			optDP(1).Checked = True
			Component(0).Corr(3) = True
		Else
			'---- Tortuosity(time) correlation is OFF!
			txtTort.Enabled = True
			'---- Enable Ds user-input
			optDS(0).Enabled = True
			'---- Enable Dp user-input
			optDP(0).Enabled = True
		End If
	End Sub
	Sub Update_Tortuosity_Display()
		Dim T As Double
		If (Component(0).Use_Tortuosity_Correlation) Then
			'---- Tortuosity(time) correlation is ON!
			T = Tortuosity(0)
			txtTort.Text = Format_It(T, 3)
			''''chkTortuosity_Corr.Value = True
			txtTort.Enabled = False
			'lblTortCorrelation.Caption = "For t<=70 days, tortuosity = 1;" & Chr$(13) & Chr$(10) & "For t>70 days, tortuosity = 0.334 + 0.00000661*(t,minutes)"
			lblTortCorrelation.Text = "For t<=70 days, tortuosity = 1;" & Chr(13) & Chr(10) & "For t>70 days, tortuosity = 0.334 + 0.009518*(t,days)"
			lblTortCorrelation.Visible = True
			lblTortCorrelation.Left = txtTort.Left
			lblTortCorrelation.Top = txtTort.Top
			txtTort.Visible = False
			lblTort.Visible = False
			Call frmKinetic_Refresh()
			'THIS REDISPLAYS txtTort.
		Else
			'---- Tortuosity(time) correlation is OFF!
			T = Component(0).Tortuosity
			Call frmKinetic_Repopulate_Values() 'THIS REDISPLAYS txtTort.
			'txtTort = Format_It(T, 3)
			''''chkTortuosity_Corr.Value = False
			txtTort.Enabled = True
			lblTortCorrelation.Visible = False
			txtTort.Visible = True
			lblTort.Visible = True
		End If
	End Sub
	
	
	Private Sub chkTortuosity_Corr_Click(ByRef Value As Short)
		If (Value = True) Then
			'---- Turn tortuosity(time) correlation ON!
			Component(0).Use_Tortuosity_Correlation = True
			Component(0).Constant_Tortuosity = False
			'frmprint!chkSelect(4).Enabled = True
			'---- Update SPDFR to 1.000e-30!
			Component(0).SPDFR = 1E-30
			'txtSPDFR.Text = "1.000E-30"
			'txtSPDFR = Format_It(Component(0).SPDFR, 3)
			Call frmKinetic_Refresh()
			'THIS REDISPLAYS txtSPDFR AND lblKF,lblDS,lblDP.
			'---- LOCK SPDFR.
			txtSPDFR.ReadOnly = True
		Else
			'---- Turn tortuosity(time) correlation OFF!
			Component(0).Use_Tortuosity_Correlation = False
			Component(0).Constant_Tortuosity = False
			'frmprint!chkSelect(4).Enabled = False
			'---- UNLOCK SPDFR.
			txtSPDFR.ReadOnly = False
		End If
		'lblDS = Format$(Ds(0), "0.00E+00")
		Call Update_Ds_and_Dp_Editability()
		Call Update_Tortuosity_Display()
		'THROW DIRTY FLAG.
		Call frmKinetic_DirtyStatus_Throw()
		'REFRESH WINDOW.
		Call frmKinetic_Refresh()
	End Sub


	Private Sub cmdCancelOK_Click(ByRef Index As Short)
		'Private Sub cmdCancelOK_click(sender As Object, e As EventArgs) Handles _cmdCancelOK_1.ClickEvent, _cmdCancelOK_0.ClickEvent
		'Dim index As Short = Array.IndexOf(cmdCancelOK, sender)
		Select Case Index
			Case 0 'CANCEL.
				'ROLLBACK TO ORIGINAL COMPONENT DATA.
				'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Component(0) = SaveOldComponent
				'EXIT OUT OF HERE.
				USER_HIT_CANCEL = True
				USER_HIT_OK = False
				Me.Close()
				Exit Sub
			Case 1 'OK.
				'SAVE USER/CORRELATION SELECTION OPTIONBOXES.
				Component(0).Corr(1) = optKF(1).Checked
				Component(0).Corr(2) = optDS(1).Checked
				Component(0).Corr(3) = optDP(1).Checked
				'SAVE USER INPUT FOR KF/DS/DP/SPDFR.
				Component(0).KP_User_Input(1) = CDbl(txtKF.Text)
				Component(0).KP_User_Input(2) = CDbl(txtDS.Text)
				Component(0).KP_User_Input(3) = CDbl(txtDP.Text)
				Component(0).SPDFR = CDbl(txtSPDFR.Text)
				'SAVE CURRENT VALUES FOR KF/DS/DP.
				If optKF(0).Checked = True Then
					Component(0).kf = CDbl(txtKF.Text)
				Else
					Component(0).kf = kf(0)
				End If
				If optDS(0).Checked = True Then
					Component(0).Ds = CDbl(txtDS.Text)
				Else
					Component(0).Ds = Ds(0)
				End If
				If optDP(0).Checked = True Then
					Component(0).Dp = CDbl(txtDP.Text)
				Else
					Component(0).Dp = Dp(0)
				End If
				'SAVE CURRENT VALUE FOR TORTUOSITY AND CORRELATION SETTINGS.
				'UPGRADE_WARNING: Couldn't resolve default property of object chkTortuosity_Corr.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (chkTortuosity_Corr.Checked) Then
					Component(0).Use_Tortuosity_Correlation = True
					Component(0).Constant_Tortuosity = False
					Component(0).Tortuosity = 1
				Else
					Component(0).Use_Tortuosity_Correlation = False
					Component(0).Constant_Tortuosity = False
					Component(0).Tortuosity = CDbl(txtTort.Text) 'iffy conversion here!
				End If
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
	
	Private Sub frmKinetic_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Short
		rs.FindAllControls(Me)

		'MISC INITS.
		'	Me.Height = VB6.TwipsToPixelsY(6450)
		'	Me.Width = VB6.TwipsToPixelsX(8535)
		Call CenterOnForm(Me, frmCompoProp)
		Me.Text = "Kinetic Parameters for " & Trim(Component(0).Name)
		lblUnit(1).Text = "cm" & Chr(178) & "/s"
		lblUnit(2).Text = "cm" & Chr(178) & "/s"
		'TORTUOSITY CORRELATION DISPLAY.
		lblTortCorrelation.Visible = False
		If (Component(0).Use_Tortuosity_Correlation) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object chkTortuosity_Corr.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkTortuosity_Corr.Checked = True
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object chkTortuosity_Corr.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkTortuosity_Corr.Checked = False
		End If
		'DISPLAY USER/CORRELATION SELECTION OPTIONBOXES.
		optKF(0).Checked = Not (Component(0).Corr(1))
		optKF(1).Checked = Component(0).Corr(1)
		optDS(0).Checked = Not (Component(0).Corr(2))
		optDS(1).Checked = Component(0).Corr(2)
		optDP(0).Checked = Not (Component(0).Corr(3))
		optDP(1).Checked = Component(0).Corr(3)
		For i = 0 To 1
			optKF(i).Enabled = True
			optDS(i).Enabled = True
			optDP(i).Enabled = True
		Next i
		'SAVE OLD COMPONENT FOR CANCEL ROLLBACK.
		'UPGRADE_WARNING: Couldn't resolve default property of object SaveOldComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SaveOldComponent = Component(0)
		'POPULATE UNIT CONTROLS.
		Call frmKinetic_PopulateUnits()
		'DATA UNCHANGED AS YET.
		Call frmKinetic_DirtyStatus_Clear()
		Call frmKinetic_GenericStatus_Set("")
		'REFRESH DISPLAY.
		Call frmKinetic_Refresh()
		'DEMO SETTINGS.
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	Private Sub frmKinetic_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub
	
	
	
	
	Private Sub UCtl_GotFocus(ByRef Ctl As System.Windows.Forms.Control)
		Dim StatusMessagePanel As String
		Call unitsys_control_txtx_gotfocus(Ctl)
		If (Trim(UCase(Ctl.Name)) = Trim(UCase("txtKF"))) Then
			StatusMessagePanel = "Type in the user-input film diffusion coefficient"
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtDS"))) Then 
			StatusMessagePanel = "Type in the user-input surface diffusion coefficient"
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtDP"))) Then 
			StatusMessagePanel = "Type in the user-input pore diffusion coefficient"
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtSPDFR"))) Then 
			StatusMessagePanel = "Type in the user-input surface-to-pore diffusion flux ratio"
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtTort"))) Then 
			StatusMessagePanel = "Type in the user-input tortuosity"
		Else
			'NOT RECOGNIZED -- DO NOTHING.
		End If
		Call frmKinetic_GenericStatus_Set(StatusMessagePanel)
	End Sub
	Sub UCtl_LostFocus(ByRef Ctl As System.Windows.Forms.Control)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
		If (Trim(UCase(Ctl.Name)) = Trim(UCase("txtKF"))) Then
			Val_Low = 1E-40 : Val_High = 1E+40
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtDS"))) Then 
			Val_Low = 1E-40 : Val_High = 1E+40
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtDP"))) Then 
			Val_Low = 1E-40 : Val_High = 1E+40
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtSPDFR"))) Then 
			Val_Low = 1E-40 : Val_High = 1E+40
		ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtTort"))) Then 
			Val_Low = 1E-40 : Val_High = 1E+40
		Else
			'NOT RECOGNIZED -- DO NOTHING.
		End If
		NewValue_Okay = False
		If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
			NewValue_Okay = True
		End If
		Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
		Call frmKinetic_GenericStatus_Set("")
		If (NewValue_Okay) Then
			If (Raise_Dirty_Flag) Then
				'STORE TO MEMORY.
				If (Trim(UCase(Ctl.Name)) = Trim(UCase("txtKF"))) Then
					Component(0).KP_User_Input(1) = NewValue
				ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtDS"))) Then 
					Component(0).KP_User_Input(2) = NewValue
				ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtDP"))) Then 
					Component(0).KP_User_Input(3) = NewValue
				ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtSPDFR"))) Then 
					Component(0).SPDFR = NewValue
				ElseIf (Trim(UCase(Ctl.Name)) = Trim(UCase("txtTort"))) Then 
					Component(0).Tortuosity = NewValue
				Else
					'NOT RECOGNIZED -- DO NOTHING.
				End If
				'RAISE DIRTY FLAG IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					'THROW DIRTY FLAG.
					Call frmKinetic_DirtyStatus_Throw()
				End If
				'REFRESH WINDOW.
				Call frmKinetic_Refresh()
			End If
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event optDP.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDP_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDP.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optDP.GetIndex(eventSender)
			'THROW DIRTY FLAG.
			Call frmKinetic_DirtyStatus_Throw()
			'REFRESH WINDOW.
			Call frmKinetic_Refresh()
		End If
	End Sub
	'UPGRADE_WARNING: Event optDS.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDS_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDS.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optDS.GetIndex(eventSender)
			'THROW DIRTY FLAG.
			Call frmKinetic_DirtyStatus_Throw()
			'REFRESH WINDOW.
			Call frmKinetic_Refresh()
		End If
	End Sub
	'UPGRADE_WARNING: Event optKF.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optKF_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optKF.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optKF.GetIndex(eventSender)
			'THROW DIRTY FLAG.
			Call frmKinetic_DirtyStatus_Throw()
			'REFRESH WINDOW.
			Call frmKinetic_Refresh()
		End If
	End Sub
	
	
	Private Sub txtDP_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDP.Enter
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtDP : Call UCtl_GotFocus(Ctl)
	End Sub
	Private Sub txtDP_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDP.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtDP_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDP.Leave
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtDP : Call UCtl_LostFocus(Ctl)
	End Sub

	Private Sub txtDS_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDS.Enter
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtDS : Call UCtl_GotFocus(Ctl)
	End Sub
	Private Sub txtDS_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtDS_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDS.Leave
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtDS : Call UCtl_LostFocus(Ctl)
	End Sub

	Private Sub txtKF_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtKF.Enter
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtKF : Call UCtl_GotFocus(Ctl)
	End Sub
	Private Sub txtKF_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtKF_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtKF.Leave
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtKF : Call UCtl_LostFocus(Ctl)
	End Sub

	Private Sub txtSPDFR_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSPDFR.Enter
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtSPDFR : Call UCtl_GotFocus(Ctl)
	End Sub
	Private Sub txtSPDFR_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSPDFR.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtSPDFR_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSPDFR.Leave
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtSPDFR : Call UCtl_LostFocus(Ctl)
	End Sub
	
	Private Sub txtTort_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTort.Enter
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtTort : Call UCtl_GotFocus(Ctl)
	End Sub
	Private Sub txtTort_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTort.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtTort_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTort.Leave
		Dim Ctl As System.Windows.Forms.Control : Ctl = txtTort : Call UCtl_LostFocus(Ctl)
	End Sub

	Private Sub lblCorrelationDS_Click(sender As Object, e As EventArgs) Handles lblCorrelationDS.Click

	End Sub



	Private Sub _cmdCancelOK_1_ClickEvent(sender As Object, e As EventArgs)
		Call cmdCancelOK_Click(1)
	End Sub

	Private Sub _cmdCancelOK_0_ClickEvent(sender As Object, e As EventArgs)
		Call cmdCancelOK_Click(0)
	End Sub

	Private Sub OK_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_1.Click
		Call cmdCancelOK_Click(1)
	End Sub

	Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_0.Click
		Call cmdCancelOK_Click(0)
	End Sub

	Private Sub chkTortuosity_Corr_CheckedChanged(sender As Object, e As EventArgs) Handles chkTortuosity_Corr.CheckedChanged

	End Sub

	Private Sub chkTortuosity_Corr_Click(sender As Object, e As EventArgs) Handles chkTortuosity_Corr.Click
		If (chkTortuosity_Corr.Checked = True) Then
			'---- Turn tortuosity(time) correlation ON!
			Component(0).Use_Tortuosity_Correlation = True
			Component(0).Constant_Tortuosity = False
			'frmprint!chkSelect(4).Enabled = True
			'---- Update SPDFR to 1.000e-30!
			Component(0).SPDFR = 1.0E-30
			'txtSPDFR.Text = "1.000E-30"
			'txtSPDFR = Format_It(Component(0).SPDFR, 3)
			Call frmKinetic_Refresh()
			'THIS REDISPLAYS txtSPDFR AND lblKF,lblDS,lblDP.
			'---- LOCK SPDFR.
			txtSPDFR.ReadOnly = True
		Else
			'---- Turn tortuosity(time) correlation OFF!
			Component(0).Use_Tortuosity_Correlation = False
			Component(0).Constant_Tortuosity = False
			'frmprint!chkSelect(4).Enabled = False
			'---- UNLOCK SPDFR.
			txtSPDFR.ReadOnly = False
		End If
		'lblDS = Format$(Ds(0), "0.00E+00")
		Call Update_Ds_and_Dp_Editability()
		Call Update_Tortuosity_Display()
		'THROW DIRTY FLAG.
		Call frmKinetic_DirtyStatus_Throw()
		'REFRESH WINDOW.
		Call frmKinetic_Refresh()
	End Sub

	Private Sub txtDP_TextChanged(sender As Object, e As EventArgs) Handles txtDP.TextChanged

	End Sub

	Private Sub txtDS_TextChanged(sender As Object, e As EventArgs) Handles txtDS.TextChanged

	End Sub

	Private Sub frmKinetic_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class