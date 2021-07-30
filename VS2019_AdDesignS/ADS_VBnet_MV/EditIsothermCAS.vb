Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmEditIsothermCAS
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer


	Dim USER_HIT_CANCEL As Boolean
	Dim USER_HIT_OK As Boolean
	
	
	Dim Use_Title As String
	Dim Use_TextLabel_1a As String
	Dim Use_TextLabel_1b As String
	Dim Use_TextLabel_2a As String
	Dim Use_TextLabel_2b As String
	Dim Use_OptionLabel_1 As String
	Dim Use_OptionLabel_2 As String
	Dim Use_OK_Caption As String
	
	Dim VarText_1a As String
	Dim VarText_1b As String
	Dim VarText_2a As String
	Dim VarText_2b As String
	Dim VarBool_1 As Boolean
	Dim VarBool_2 As Boolean
	
	
	Const frmEditIsothermCAS_declarations_end As Boolean = True
	
	
	Public Sub frmEditIsothermCAS_Run(ByRef INPUT_Use_Title As String, ByRef INPUT_Use_TextLabel_1a As String, ByRef INPUT_Use_TextLabel_1b As String, ByRef INPUT_Use_TextLabel_2a As String, ByRef INPUT_Use_TextLabel_2b As String, ByRef INPUT_Use_OptionLabel_1 As String, ByRef INPUT_Use_OptionLabel_2 As String, ByRef INPUT_Use_OK_Caption As String, ByRef OUTPUT_USER_HIT_CANCEL As Boolean, ByRef IO_VarText_1a As String, ByRef IO_VarText_1b As String, ByRef IO_VarText_2a As String, ByRef IO_VarText_2b As String, ByRef IO_VarBool_1 As Boolean, ByRef IO_VarBool_2 As Boolean)
		Use_Title = INPUT_Use_Title
		Use_TextLabel_1a = INPUT_Use_TextLabel_1a
		Use_TextLabel_1b = INPUT_Use_TextLabel_1b
		Use_TextLabel_2a = INPUT_Use_TextLabel_2a
		Use_TextLabel_2b = INPUT_Use_TextLabel_2b
		Use_OptionLabel_1 = INPUT_Use_OptionLabel_1
		Use_OptionLabel_2 = INPUT_Use_OptionLabel_2
		Use_OK_Caption = INPUT_Use_OK_Caption
		VarText_1a = IO_VarText_1a
		VarText_1b = IO_VarText_1b
		VarText_2a = IO_VarText_2a
		VarText_2b = IO_VarText_2b
		VarBool_1 = IO_VarBool_1
		VarBool_2 = IO_VarBool_2
		Me.ShowDialog()
		OUTPUT_USER_HIT_CANCEL = USER_HIT_CANCEL
		If (USER_HIT_OK) Then
			IO_VarText_1a = VarText_1a
			IO_VarText_1b = VarText_1b
			IO_VarText_2a = VarText_2a
			IO_VarText_2b = VarText_2b
			IO_VarBool_1 = VarBool_1
			IO_VarBool_2 = VarBool_2
		End If
	End Sub
	
	
	
	
	
	
	Private Sub cmdSaveCancel_Click(ByRef Index As Short)
		Select Case Index
			Case 0 'OK.
				'TRANSFER DATA FROM CONTROLS TO MEMORY.
				VarText_1a = txtData(0).Text
				VarText_1b = txtData(1).Text
				VarText_2a = txtData(2).Text
				VarText_2b = txtData(3).Text
				'UPGRADE_WARNING: Couldn't resolve default property of object chkData().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				VarBool_1 = _chkData_0.Checked
				'UPGRADE_WARNING: Couldn't resolve default property of object chkData().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				VarBool_2 = _chkData_1.Checked
				'EXIT OUT OF HERE.
				USER_HIT_CANCEL = False
				USER_HIT_OK = True
				Me.Close()
				Exit Sub
			Case 1 'CANCEL.
				'EXIT OUT OF HERE.
				USER_HIT_CANCEL = True
				USER_HIT_OK = False
				Me.Close()
				Exit Sub
		End Select
	End Sub
	
	
	Private Sub frmEditIsothermCAS_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		'MISC INITS.
		Me.Text = Use_Title
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdSaveCancel().Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_cmdSaveCancel_0.Text = Use_OK_Caption
		If (VB.Left(Use_TextLabel_1a, 1) = "^") Then
			Use_TextLabel_1a = VB.Right(Use_TextLabel_1a, Len(Use_TextLabel_1a) - 1)
			txtData(0).ReadOnly = True
			txtData(0).BackColor = System.Drawing.ColorTranslator.FromOle(QBColor(7))
		End If
		If (VB.Left(Use_TextLabel_1b, 1) = "^") Then
			Use_TextLabel_1b = VB.Right(Use_TextLabel_1b, Len(Use_TextLabel_1b) - 1)
			txtData(1).ReadOnly = True
			txtData(1).BackColor = System.Drawing.ColorTranslator.FromOle(QBColor(7))
		End If
		lblDesc(0).Text = Use_TextLabel_1a
		lblDesc(1).Text = Use_TextLabel_1b
		lblDesc(2).Text = Use_TextLabel_2a
		lblDesc(3).Text = Use_TextLabel_2b
		'UPGRADE_WARNING: Couldn't resolve default property of object chkData().Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_chkData_0.Text = Use_OptionLabel_1
		'UPGRADE_WARNING: Couldn't resolve default property of object chkData().Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_chkData_1.Text = Use_OptionLabel_2
		txtData(0).Text = VarText_1a
		txtData(1).Text = VarText_1b
		txtData(2).Text = VarText_2a
		txtData(3).Text = VarText_2b
		'UPGRADE_WARNING: Couldn't resolve default property of object chkData().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_chkData_0.Checked = VarBool_1
		'UPGRADE_WARNING: Couldn't resolve default property of object chkData().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_chkData_1.Checked = VarBool_2
		If (lblDesc(2).Text = "*") Then
			lblDesc(2).Visible = False
			lblDesc(3).Visible = False
			txtData(2).Visible = False
			txtData(3).Visible = False
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object chkData(0).Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (_chkData_0.Text = "*") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object chkData().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_chkData_0.Visible = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkData().Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			_chkData_1.Visible = False
		End If
		'		Me.Height = VB6.TwipsToPixelsY(3390)
		'		Me.Width = VB6.TwipsToPixelsX(7650)
		Call CenterOnForm(Me, frmEditAdsorber)
	End Sub
	
	
	
	
	
	
	Private Sub txtData_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Enter
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim txtCtl As System.Windows.Forms.TextBox
		txtCtl = txtData(Index)
		'UPGRADE_WARNING: Couldn't resolve default property of object txtCtl.Locked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Not txtCtl.Enabled) Then Exit Sub
		Call Global_GotFocus(txtCtl)
	End Sub
	Private Sub txtData_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtData.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtData.GetIndex(eventSender)
		KeyAscii = Global_TextKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtData_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Leave
		Dim Index As Short = txtData.GetIndex(eventSender)
		Dim txtCtl As System.Windows.Forms.Control
		txtCtl = txtData(Index)
		'UPGRADE_WARNING: Couldn't resolve default property of object txtCtl.Locked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Not txtCtl.Enabled) Then Exit Sub
		Call Global_LostFocus(txtCtl)
	End Sub

    Private Sub _txtData_2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _txtData_2.TextChanged

    End Sub

	Private Sub _cmdSaveCancel_0_ClickEvent(sender As Object, e As EventArgs)
		Call cmdSaveCancel_Click(0)
	End Sub

	Private Sub _cmdSaveCancel_1_ClickEvent(sender As Object, e As EventArgs)
		Call cmdSaveCancel_Click(1)
	End Sub

	Private Sub Save_Click(sender As Object, e As EventArgs) Handles _cmdSaveCancel_0.Click
		Call cmdSaveCancel_Click(0)
	End Sub

	Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles _cmdSaveCancel_1.Click
		Call cmdSaveCancel_Click(1)
	End Sub

	Private Sub frmEditIsothermCAS_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class