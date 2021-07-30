Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmDimensionless
	Inherits System.Windows.Forms.Form


	Dim rs As New Resizer

	Sub populate_cboSelectCompo()
		Dim i As Short
		cboSelectCompo.Items.Clear()
		For i = 1 To Number_Component
			cboSelectCompo.Items.Add(Trim(Component(i).Name))
		Next i
	End Sub
	
	
	Sub Display_Component(ByRef Which As Short)
		Dim N As Short
		If (Which < 1) Or (Which > Number_Component) Then
			Exit Sub
		End If
		N = Which
		Call AssignTextAndTag(txtDimless(0), NumberToMFBString(ST(N)))
		Call AssignTextAndTag(txtDimless(1), NumberToMFBString(Eds(N)))
		Call AssignTextAndTag(txtDimless(2), NumberToMFBString(Edp(N)))
		Call AssignTextAndTag(txtDimless(3), NumberToMFBString(Bip(N)))
		Call AssignTextAndTag(txtDimless(4), NumberToMFBString(Bis(N)))
		Call AssignTextAndTag(txtDimless(5), NumberToMFBString(Dgp(N)))
		Call AssignTextAndTag(txtDimless(6), NumberToMFBString(Dgs(N)))
		If (Which <= cboSelectCompo.Items.Count) Then
			cboSelectCompo.SelectedIndex = Which - 1
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event cboSelectCompo.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboSelectCompo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSelectCompo.SelectedIndexChanged
		Call Display_Component(cboSelectCompo.SelectedIndex + 1)
	End Sub
	
	
	Private Sub cmdDefs_Click()
		'Me.Hide
		'frmDimensionlessDefs.Show
		'Me.Show
	End Sub



	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture2.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture2.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
		
	End Sub
	
	Private Sub frmDimensionless_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'MISC INITS.
		rs.FindAllControls(Me)

		Call CenterOnForm(Me, frmMain)
		Call populate_cboSelectCompo()
		Call Display_Component(frmMain.cboSelectCompo.SelectedIndex + 1)
	End Sub
	
	
	Private Sub txtDimless_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDimless.Enter
		Dim Index As Short = txtDimless.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtDimless(Index)
		Call Global_GotFocus(Ctl)
	End Sub
	Private Sub txtDimless_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDimless.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtDimless.GetIndex(eventSender)
		KeyAscii = Global_ReadOnlyKeyPress(KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtDimless_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDimless.Leave
		Dim Index As Short = txtDimless.GetIndex(eventSender)
		Dim Ctl As System.Windows.Forms.Control
		Ctl = txtDimless(Index)
		Call Global_LostFocus(Ctl)
	End Sub



	Private Sub Close_Click(sender As Object, e As EventArgs) Handles cmdclose.Click
		Me.Close()
	End Sub

	Private Sub frmDimensionless_RegionChanged(sender As Object, e As EventArgs) Handles Me.RegionChanged

	End Sub

	Private Sub frmDimensionless_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)


	End Sub
End Class