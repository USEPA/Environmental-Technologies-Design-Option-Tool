Option Strict Off
Option Explicit On
Friend Class frmSplash
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer




	Const frmSplash_decl_end As Boolean = True
	
	
	Private Sub cmdButton1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdButton1.Click
		splash_button_pressed = 1
		Me.Close()
	End Sub
	Private Sub cmdButton2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdButton2.Click
		splash_button_pressed = 2
		Me.Close()
	End Sub
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		splash_button_pressed = 3
		Me.Close()
	End Sub
	
	
	Private Sub frmSplash_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ctr_location As Short
		Dim s As String

		rs.FindAllControls(Me)

		'Call debug_output("L1")
		'	Me.Height = VB6.TwipsToPixelsY(6165)
		'		Me.Width = VB6.TwipsToPixelsX(9300)'
		'		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		'		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
		'Call debug_output("L2")
		'
		' CENTER THE LOGOS.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object SSPanelLogos.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	ctr_location = SSPanelLogos.Width / 2 - VB6.PixelsToTwipsX(_lblCompany_0.Width) / 2
		'	_lblCompany_0.Left = VB6.TwipsToPixelsX(ctr_location)
		'  ctr_location = SSPanelLogos.Top + SSPanelLogos.Height / 4 - SSPanel1.Height / 2
		'  SSPanel1.Top = ctr_location
		'UPGRADE_WARNING: Couldn't resolve default property of object SSPanel1.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SSPanelLogos.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object SSPanelLogos.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	ctr_location = SSPanelLogos.Left + 3 * SSPanelLogos.Width / 4 - SSPanel1.Width / 2
		'	_picLogos_1.Visible = True
		'UPGRADE_WARNING: Couldn't resolve default property of object SSPanelLogos.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		ctr_location = SSPanelLogos.Width / 2 - VB6.PixelsToTwipsX(_picLogos_1.Width) / 2
		'		_picLogos_1.Left = VB6.TwipsToPixelsX(ctr_location)
		'UPGRADE_WARNING: Couldn't resolve default property of object SSPanelLogos.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'ctr_location = SSPanelLogos.Width / 2 - VB6.PixelsToTwipsX(_picLogos_2.Width) / 2
		'	_picLogos_2.Left = VB6.TwipsToPixelsX(ctr_location)
		'		_picLogos_1.Top = VB6.TwipsToPixelsY(50)
		'UPGRADE_WARNING: Couldn't resolve default property of object SSPanelLogos.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	_picLogos_2.Top = VB6.TwipsToPixelsY(SSPanelLogos.Height - VB6.PixelsToTwipsY(_picLogos_2.Height) - 50)
		'		ctr_location = VB6.PixelsToTwipsY(_picLogos_1.Top) + VB6.PixelsToTwipsY(_picLogos_1.Height) + 50
		'		_lblCompany_0.Top = VB6.TwipsToPixelsY(ctr_location)
		'
		' MISCELLANEOUS SETTINGS.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object sspNames.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Me.LabelAuthor.Text = "David R. Hokanson" & Environment.NewLine & "David W. Hand" & Environment.NewLine & "John C. Crittenden" & Environment.NewLine & "Tony N. Rogers" & Environment.NewLine & "Eric J. Oman"
		'lblAdditionalNotice.Text = "This program is protected by U.S. and international" & vbCrLf & "copyright laws as described in Help About."
		'_lblCompany_0.Text = "National Center for" & vbCrLf & "Clean Industrial and Treatment Technologies" & vbCrLf & "Michigan Technological University" & vbCrLf & "Houghton, Michigan"
		'picTitle.Visible = True
		'
		' LICENSE-RELATED SETTINGS.
		''''lblVersionInfo(0).Caption = "Version " & get_program_version_with_build_info()
		''''lblVersionInfo(0).Caption = _
		'"Version " & get_program_version_with_build_info_VB4(True) & _
		'" (" & get_program_releasetype() & ")"
		'_lblVersionInfo_0.Text = get_program_version_with_build_info_VB4(True)
		'lblVersionInfo(0).Caption = _
		''    "Version 1.0"
		'    MsgBox "Fix this !!!!  (contact ejoman@mtu.edu)"
		'_lblVersionInfo_1.Text = get_expiration_info(True)
		'_lblVersionInfo_2.Text = "Copyright " & AppCopyrightYears
		'
		' PROGRAM-SPECIFIC SETTINGS.
		'
		If (AppProgramKey = "ADS") Then
			If (Activate_PSDMInRoom = True) Then
				''''sspanel_maintitle.Caption = "Indoor Air Filtration Model"
				''''lblCompany(0).Caption = "MTU"
				'_lblCompany_0.Text = "Michigan Technological University" & Chr(13) & "Houghton, Michigan"
				'			_picLogos_1.Visible = False
				'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_maintitle.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'sssspanel_maintitle.Text = AppName_For_Display_Long & " (" & AppName_For_Display_Short & ")"
				'picTitle.Visible = False
			Else
				' DO NOTHING.
			End If
		End If
		If (AppProgramKey = "ASAP") Then
			' DO NOTHING.
		End If
		If (AppProgramKey = "STEPP") Then
			' DO NOTHING.
		End If
		If (splash_mode = 0) Then
			'
			' SHOW THE CONTINUE/EXIT FRONT WINDOW.
			'
			'Call debug_output("L3")
			cmdButton1.Visible = True
			cmdButton1.Text = "&Continue"
			cmdButton2.Visible = False
			cmdExit.Visible = True
			'
			' ETC.
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_disclaimer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'sspanel_disclaimer.Visible = False
			'cmdButton1.SetFocus
			cmdButton1.TabIndex = 0
			'Call debug_output("L4")
		End If
		''''If (splash_mode = 1) Then
		If (splash_mode = 1) Or (splash_mode = 101) Then
			'Call debug_output("L5")
			'SHOW THE DISCLAIMER WINDOW.
			If (splash_mode = 101) Then
				cmdButton1.Visible = False
				cmdButton2.Visible = False
				cmdExit.Visible = True
				cmdExit.Text = "&Close"
			Else
				cmdButton1.Visible = True
				cmdButton1.Text = "I Agree"
				cmdButton2.Visible = True
				cmdExit.Visible = True
			End If
			'cmdButton1.Visible = True
			'cmdButton1.Caption = "I Agree"
			'cmdButton2.Visible = True
			'cmdExit.Visible = True
			'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_logos.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'sspanel_logos.Visible = False
			'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_disclaimer.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_logos.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'sspanel_disclaimer.Left = sspanel_logos.Left
			'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_disclaimer.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_logos.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'sspanel_disclaimer.Top = sspanel_logos.Top
			'UPGRADE_WARNING: Couldn't resolve default property of object sspanel_disclaimer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'sspanel_disclaimer.Visible = True
			s = "By choosing " & Chr(34) & "I Agree" & Chr(34) & " you acknowledge that "
			s = s & "this software is under development and not guaranteed to be free "
			s = s & "of errors.  Furthermore there may be errors in the software that "
			s = s & "lead to erroneous output.  MTU shall not be liable for any loss, "
			s = s & "damage, injury, or casualty of whatsoever kind, or by whomsoever "
			s = s & "caused to the person or property of anyone arising out of or "
			s = s & "resulting from receipt and use of any aspect of the software.  "
			s = s & "References to specific commercial products, processes, or services "
			s = s & "by trademark, manufacturer, or otherwise does not necessarily "
			s = s & "constitute or imply endorsement/recommendation by the authors or "
			s = s & "the respective organizations under which the software "
			s = s & "was developed."
			'lblDisclaimer.Text = s
			'cmdButton1.SetFocus
			cmdButton1.TabIndex = 0
			'Call debug_output("L6")
		End If
	End Sub
	
	Private Sub frmSplash_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		If (splash_button_pressed = 0) Then
			'If they got here, they must have selected "Close",
			'so perform the exit functionality.
			splash_button_pressed = 3
		End If
	End Sub
	
	Private Sub lblAuthors_Click(ByRef Index As Short)
		
	End Sub

	Private Sub _picLogos_1_Click(sender As Object, e As EventArgs)

	End Sub

	Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

	End Sub

	Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

	End Sub

	Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
		Call ShellExecute_URL("http://github.com/USEPA/Environmental-Technologies-Design-Option-Tool")

	End Sub

	Private Sub frmSplash_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class