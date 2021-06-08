Option Strict Off
Option Explicit On
Friend Class frmNewName
	Inherits System.Windows.Forms.Form
	
	Dim return_text As String
	Dim USER_HIT_CANCEL As Boolean
	
	
	
	
	Const frmNewName_declarations_end As Boolean = True
	
	
	Public Function frmNewName_GetName(ByRef Use_Title As String, ByRef use_label As String, ByRef use_default As String, ByRef is_aborted As Boolean) As String
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		'Load(Me)
		Me.Show()
		Me.Text = Use_Title
		lblInstructions.Text = use_label
		txtdata.Text = use_default
		Me.Visible = False
		Me.ShowDialog()
		is_aborted = USER_HIT_CANCEL
		frmNewName_GetName = return_text
	End Function
	
	
	Private Sub Button_Click(ByRef Index As Short)
		Select Case Index
			Case 0 'OK
				return_text = Trim(txtData.Text)
				If (return_text = "") Then
					Call Show_Error("You must enter a non-blank string as a name.")
					Exit Sub
				End If
				USER_HIT_CANCEL = False
				Me.Close()
			Case 1 'Cancel
				USER_HIT_CANCEL = True
				Me.Close()
		End Select
	End Sub
	
	
	Private Sub frmNewName_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Me.Height = VB6.TwipsToPixelsY(1965)
		'Me.Width = VB6.TwipsToPixelsX(5805)
		Call CenterOnForm(Me, frmMain)
	End Sub
	
	
	Private Sub txtData_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Enter
		Dim txtCtl As System.Windows.Forms.Control
		txtCtl = txtData
		'Call DisplayDataEntryError
		Call Global_GotFocus(txtCtl)
	End Sub
	Private Sub txtData_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtData.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If (KeyAscii = 13) Then
			Call Button_Click(0)
		End If
		'keyascii = Global_TextKeyPress(keyascii)
		'  If (KeyAscii = 13) Then SendKeys "{TAB}", True
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtData_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtData.Leave
		Dim txtCtl As System.Windows.Forms.Control
		txtCtl = txtData
		Call Global_LostFocus(txtCtl)
	End Sub


	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		Call Button_Click(0)
	End Sub

	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		Call Button_Click(1)
	End Sub
End Class