Option Strict Off
Option Explicit On
Friend Class frmGoAway
	Inherits System.Windows.Forms.Form
	
	Dim frmGoAway_ParentForm As System.Windows.Forms.Form
	Dim frmGoAway_Caption As String
	Dim frmGoAway_Text As String
	Dim frmGoAway_CheckText As String
	Dim frmGoAway_CheckValue As Short
	
	
	
	
	
	Const frmGoAway_declarations_end As Boolean = True
	
	
	Public Sub frmGoAway_Run(ByRef INPUT_frmGoAway_ParentForm As System.Windows.Forms.Form, ByRef INPUT_frmGoAway_Caption As String, ByRef INPUT_frmGoAway_Text As String, ByRef INPUT_frmGoAway_CheckText As String, ByRef INPUTOUTPUT_frmGoAway_CheckValue As Short)
		frmGoAway_ParentForm = INPUT_frmGoAway_ParentForm
		frmGoAway_Caption = INPUT_frmGoAway_Caption
		frmGoAway_Text = INPUT_frmGoAway_Text
		frmGoAway_CheckText = INPUT_frmGoAway_CheckText
		frmGoAway_CheckValue = INPUTOUTPUT_frmGoAway_CheckValue
		Me.ShowDialog()
		INPUTOUTPUT_frmGoAway_CheckValue = frmGoAway_CheckValue
	End Sub
	
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		If (chkDisplay.CheckState = True) Then
			frmGoAway_CheckValue = 1
		Else
			frmGoAway_CheckValue = 0
		End If
		Me.Close()
	End Sub
	Private Sub frmGoAway_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ht As Double
		Me.Height = VB6.TwipsToPixelsY(3615)
		Me.Width = VB6.TwipsToPixelsX(7410)
		Me.Text = frmGoAway_Caption
		Label1.Text = frmGoAway_Text
		chkDisplay.Text = frmGoAway_CheckText
		If (frmGoAway_CheckValue = 1) Then
			chkDisplay.CheckState = True
		Else
			chkDisplay.CheckState = False
		End If
		Call CenterOnForm(Me, frmGoAway_ParentForm)
		'UPGRADE_ISSUE: PictureBox method Picture1.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'	ht = Picture1.TextHeight(frmGoAway_Text)
		Label1.Height = VB6.TwipsToPixelsY(ht * 1.05)
		Frame1.Height = VB6.TwipsToPixelsY(ht * 1.2)
	End Sub
	Private Sub frmGoAway_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'contam_prop_form.SetFocus
	End Sub
End Class