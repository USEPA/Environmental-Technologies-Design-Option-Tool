Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmPolanyi
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer

	Dim frmPolanyi_ParentForm As System.Windows.Forms.Form
	
	Dim USER_HIT_OK As Boolean
	Dim USER_HIT_CANCEL As Boolean
	
	
	
	
	
	Const frmPolanyi_declarations_end As Boolean = True
	
	
	Sub frmPolanyi_Edit(ByRef INPUT_ParentForm As System.Windows.Forms.Form, ByRef OUTPUT_Raise_Dirty_Flag As Boolean)
		frmPolanyi_ParentForm = INPUT_ParentForm
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
	
	
	Sub frmPolanyi_PopulateUnits()
		Call unitsys_register(Me, lblInput(0), txtInput(0), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblInput(1), txtInput(1), Nothing, "", "", "", "", "", 100#, False)
		Call unitsys_register(Me, lblInput(2), txtInput(2), Nothing, "", "", "", "", "", 100#, False)
	End Sub
	
	
	Private Sub cmdCancelOK_Click(ByRef Index As Short)
		Dim i As Short
		Select Case Index
			Case 0 'CANCEL.
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
	
	Private Sub frmPolanyi_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		'MISC INITS.
		Me.Height = VB6.TwipsToPixelsY(3150)
		Me.Width = VB6.TwipsToPixelsX(5445)
		Call CenterOnForm(Me, frmPolanyi_ParentForm)
		txtPolanyi.Text = Carbon.Name
		'POPULATE UNIT CONTROLS.
		Call frmPolanyi_PopulateUnits()
		'REFRESH DISPLAY.
		Call frmPolanyi_Refresh()
		'DEMO SETTINGS.
		Call LOCAL___Reset_DemoVersionDisablings()
	End Sub
	Private Sub frmPolanyi_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call unitsys_unregister_all_on_form(Me)
	End Sub
	
	
	Private Sub txtInput_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInput.Enter
		Dim Index As Short = txtInput.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtInput(Index)
		Call unitsys_control_txtx_gotfocus(Ctl)
	End Sub
	Private Sub txtInput_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInput.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtInput.GetIndex(eventSender)
		KeyAscii = Global_NumericKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtInput_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInput.Leave
		Dim Index As Short = txtInput.GetIndex(eventSender)
		Dim NewValue_Okay As Short
		Dim NewValue As Double
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtInput(Index)
		Dim Val_Low As Double
		Dim Val_High As Double
		Dim Raise_Dirty_Flag As Boolean
		Dim Too_Small As Short
		'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
		Select Case Index
			Case 0 : Val_Low = 0.05 : Val_High = 2.5
			Case 1 : Val_Low = 1E-20 : Val_High = 1E+20
			Case 2 : Val_Low = 1E-20 : Val_High = 1E+20
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
					Case 0 'W0.
						Carbon.W0 = NewValue
					Case 1 'BB.
						Carbon.BB = NewValue
					Case 2 'GM.
						Carbon.PolanyiExponent = NewValue
				End Select
				'RAISE DIRTY FLAG IF NECESSARY.
				If (Raise_Dirty_Flag) Then
					''THROW DIRTY FLAG.
					'Call frmCompoProp_DirtyStatus_Throw
				End If
				'REFRESH WINDOW.
				Call frmPolanyi_Refresh()
			End If
		End If
	End Sub

	Private Sub _cmdCancelOK_0_ClickEvent(sender As Object, e As EventArgs)
		Call cmdCancelOK_Click(0)
	End Sub

	Private Sub _cmdCancelOK_1_ClickEvent(sender As Object, e As EventArgs)
		Call cmdCancelOK_Click(1)
	End Sub

	Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_0.Click
		Call cmdCancelOK_Click(0)
	End Sub

	Private Sub OK_Click(sender As Object, e As EventArgs) Handles _cmdCancelOK_1.Click
		Call cmdCancelOK_Click(1)
	End Sub

	Private Sub frmPolanyi_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class