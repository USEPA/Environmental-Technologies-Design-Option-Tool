Option Strict Off
Option Explicit On
Friend Class frmAbout2
	Inherits System.Windows.Forms.Form
	
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub
	
	Private Sub frmAbout2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim msg As String
		
		Call CenterOnForm(Me, frmMain)
		'Move frmpfpsdm.Left + (frmpfpsdm.Width / 2) - (Me.Width / 2), frmpfpsdm.Top + (frmpfpsdm.Height / 2) - (Me.Height / 2)
		
		msg = "David R. Hokanson" & Chr(13)
		msg = msg & "David W. Hand" & Chr(13)
		msg = msg & "John C. Crittenden" & Chr(13)
		msg = msg & "Tony N. Rogers" & Chr(13)
		msg = msg & "Fr" & Chr(233) & "d" & Chr(233) & "ric Gobin" & Chr(13)
		msg = msg & "Eric J. Oman"
		'UPGRADE_WARNING: Couldn't resolve default property of object pnl_title().Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		_pnl_title_3.Text = msg
		'UPGRADE_WARNING: Couldn't resolve default property of object pnl_title().BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	_pnl_title_3.BackColor = &HC0C0C0
		'UPGRADE_WARNING: Couldn't resolve default property of object pnl_title().ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	_pnl_title_.ForeColor = &H0
		'pnl_title(3).Caption = " Model and Software:      Fr" & Chr$(233) & "d" & Chr$(233) & "ric Gobin     Eric J. Oman" & Chr$(13) & " Development   :      Tony N. Rogers"

		'pnl_title(5).Caption = " Programing Support:     Richard J. Hossli" & Chr$(13) & "                                   Jason E. Mclean"
		'pnl_title(5).BackColor = &HC0C0C0
		'pnl_title(5).ForeColor = &H0&

		''''picmtu(0).Picture = LoadPicture(app.Path & "\mtu_logo.bmp")
	End Sub
End Class