Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks

Friend Class frmAbout
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Private Sub cmdLaunchWebSite_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLaunchWebSite.Click
		Call ShellExecute_URL("https://www.epa.gov/water-research/environmental-technologies-design-option-tool-etdot")
	End Sub
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub
	
	
	Private Sub frmAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

		rs.FindAllControls(Me)


		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
		
		'VARIOUS LABELS.
		Me.Text = "About " & AppName_For_Display_Short
		lblProgramName.Text = AppName_For_Display_Long
		''''lblVersionInfo(0).Caption = "Version " & get_program_version_with_build_info()
		'lblVersionInfo(0).Text = get_program_version_with_build_info_VB4(False)
		lblVersionInfo(0).Text = "Version 1.0.50"
		lblVersionInfo(1).Text = get_expiration_info(False)
		lblVersionInfo(2).Text = "Copyright " & AppCopyrightYears
		'lblUserName.Text = Trim(lfd.Z_USERNAME)
		'lblUserCompany.Text = Trim(lfd.Z_USERCOMPANY)
		'lblSerialNumber.Text = Trim(lfd.Z_SERIALNUMBER)
		'lblSerialNumber.Caption = "WWWWWW-WWWWW-WWWWW-WWWWW-WWWWW"
		''''lblVersionInfo(5).Caption = "(Build Code " & get_program_version_with_build_info_VB4(False) & ")"
		lblVersionInfo(5).Text = ""
		
	End Sub

	Private Sub frmAbout_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)
	End Sub

	Private Sub _lblWarning_1_Click(sender As Object, e As EventArgs) Handles _lblWarning_1.Click

	End Sub
End Class