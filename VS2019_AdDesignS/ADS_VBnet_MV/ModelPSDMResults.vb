Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Math
Friend Class frmModelPSDMResults
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	'UPGRADE_WARNING: Lower bound of array Flag_TO was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim Flag_TO(Number_Compo_Max_PFPSDM) As Short

	Dim PopulatingScrollboxes As Short
	Dim HALT_ALL_CONTROLS As Boolean




	Const frmModelPSDMResults_declarations_end As Boolean = True


	Sub Populate_cboYAxisType()
		Dim Ctl As System.Windows.Forms.ComboBox
		Dim newindex As Integer
		Ctl = cboYAxisType
		HALT_ALL_CONTROLS = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ctl.Items.Clear()
		If (Results.is_psdm_in_room_model) Then
			If (Results.AnyCrCloseToZero = False) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				newindex = Ctl.Items.Add("Cr/Cr,ss")
				'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				VB6.SetItemData(Ctl, newindex, CBOYAXISTYPE_C_CO)
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			newindex = Ctl.Items.Add("C/Co")
			'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			VB6.SetItemData(Ctl, newindex, CBOYAXISTYPE_C_CO)
		End If
		newindex = Ctl.Items.Add("µg/L")
		VB6.SetItemData(Ctl, newindex, CBOYAXISTYPE_UG_L)
		'Ctl.Items.Add(New VB6.ListBoxItem("µg/L", CBOYAXISTYPE_UG_L))
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

		newindex = Ctl.Items.Add("mg/L")
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		VB6.SetItemData(Ctl, newindex, CBOYAXISTYPE_MG_L)

		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		newindex = Ctl.Items.Add("g/L")
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		VB6.SetItemData(Ctl, newindex, CBOYAXISTYPE_G_L)

		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

		newindex = Ctl.Items.Add("ppb")
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		VB6.SetItemData(Ctl, newindex, CBOYAXISTYPE_PPB)
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

		newindex = Ctl.Items.Add("ppm")
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		VB6.SetItemData(Ctl, newindex, CBOYAXISTYPE_PPM)

		'nanograms addition MV
		newindex = Ctl.Items.Add("ng/L")
		VB6.SetItemData(Ctl, newindex, CBOYAXISTYPE_NG_L)

		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ctl.SelectedIndex = 0
		HALT_ALL_CONTROLS = False
	End Sub


	'UPGRADE_WARNING: Event cboCompo.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCompo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCompo.SelectedIndexChanged
		Dim f As Double
		Dim index As Integer

		If (PopulatingScrollboxes) Then Exit Sub
		If (Results.is_psdm_in_room_model) Then
			lblSSValue.Text = NumberToMFBString(Results.psdmroom_Crss(cboCompo.SelectedIndex + 1))
		End If
		If (Results.ThroughPut_50(cboCompo.SelectedIndex + 1).C <> -1.0#) And (Results.ThroughPut_50(cboCompo.SelectedIndex + 1).T <> -1.0#) And (Results.ThroughPut_05(cboCompo.SelectedIndex + 1).T <> -1.0#) And (Results.ThroughPut_05(cboCompo.SelectedIndex + 1).C <> -1.0#) And (Results.ThroughPut_95(cboCompo.SelectedIndex + 1).T <> -1.0#) And (Results.ThroughPut_95(cboCompo.SelectedIndex + 1).C <> -1.0#) Then
			f = 100 * Results.Bed.length / Results.ThroughPut_50(cboCompo.SelectedIndex + 1).T 'in cm/days
			lblMTZ.Text = Format_It(f * (Results.ThroughPut_95(cboCompo.SelectedIndex + 1).T - Results.ThroughPut_05(cboCompo.SelectedIndex + 1).T), 3)
		Else
			lblMTZ.Text = "N/A"
		End If
		If (Results.ThroughPut_05(cboCompo.SelectedIndex + 1).T <> -1.0#) And (Results.ThroughPut_05(cboCompo.SelectedIndex + 1).C <> -1.0#) Then
			lblData(0).Text = Format_It(Results.ThroughPut_05(cboCompo.SelectedIndex + 1).T / 24.0# / 60.0#, 2)
			lblData(9).Text = Format_It(Results.ThroughPut_05(cboCompo.SelectedIndex + 1).C, 2)
			lblData(3).Text = Format_It(Results.ThroughPut_05(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
			lblData(6).Text = Format_It(Results.ThroughPut_05(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
		Else
			lblData(0).ResetText()
			lblData(0).Text = "N/A"
			lblData(9).Text = "N/A"
			'lblData(12) = "N/A"
			lblData(3).Text = "N/A"
			lblData(6).Text = "N/A"
		End If
		If (Results.ThroughPut_50(cboCompo.SelectedIndex + 1).T <> -1.0#) And (Results.ThroughPut_50(cboCompo.SelectedIndex + 1).C <> -1.0#) Then
			lblData(1).Text = Format_It(Results.ThroughPut_50(cboCompo.SelectedIndex + 1).T / 24.0# / 60.0#, 2)
			lblData(10).Text = Format_It(Results.ThroughPut_50(cboCompo.SelectedIndex + 1).C, 2)
			lblData(4).Text = Format_It(Results.ThroughPut_50(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
			lblData(7).Text = Format_It(Results.ThroughPut_50(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
		Else
			lblData(1).Text = "N/A"
			lblData(10).Text = "N/A"
			lblData(4).Text = "N/A"
			lblData(7).Text = "N/A"
		End If
		If (Results.ThroughPut_95(cboCompo.SelectedIndex + 1).T <> -1.0#) And (Results.ThroughPut_95(cboCompo.SelectedIndex + 1).C <> -1.0#) Then
			lblData(2).Text = Format_It(Results.ThroughPut_95(cboCompo.SelectedIndex + 1).T / 24.0# / 60.0#, 2)
			lblData(11).Text = Format_It(Results.ThroughPut_95(cboCompo.SelectedIndex + 1).C, 2)
			'lblData(14) = Format_It(Results.ThroughPut_95(cboCompo.ListIndex + 1).Q)
			'lblData(14) = "N/A"
			lblData(5).Text = Format_It(Results.ThroughPut_95(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
			lblData(8).Text = Format_It(Results.ThroughPut_95(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
		Else
			lblData(2).Text = "N/A"
			lblData(11).Text = "N/A"
			'lblData(14) = "N/A"
			lblData(5).Text = "N/A"
			lblData(8).Text = "N/A"
		End If
		'cmdTreat.Caption = Format_It(Treament_Objective(cboCompo.ListIndex + 1).C, 2) & " mg/L"
		If Flag_TO(cboCompo.SelectedIndex + 1) Then
			lblData(12).Text = Format_It(Treatment_Objective(cboCompo.SelectedIndex + 1).T / 60.0# / 24.0#, 2)
			lblData(13).Text = Format_It(Treatment_Objective(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
			lblData(14).Text = Format_It(Treatment_Objective(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
			lblData(15).Text = Format_It(Treatment_Objective(cboCompo.SelectedIndex + 1).C, 2)
		Else
			lblData(12).Text = "N/A"
			lblData(13).Text = "N/A"
			lblData(14).Text = "N/A"
			lblData(15).Text = "N/A"
		End If

	End Sub


	'UPGRADE_WARNING: Event cboGrid.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboGrid_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboGrid.SelectedIndexChanged
		If (Not PopulatingScrollboxes) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GridStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.GridStyle = cboGrid.SelectedIndex
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.DrawMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.DrawMode = 2
			Select Case cboGrid.SelectedIndex
				Case 0
					Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 0
					Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 0
				Case 1
					Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 0
					Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 1
				Case 2
					Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 1
					Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 0
				Case 3
					Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 1
					Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 1
			End Select

		End If
	End Sub


	'UPGRADE_WARNING: Event cboYAxisType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboYAxisType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboYAxisType.SelectedIndexChanged
		If (HALT_ALL_CONTROLS = True) Then Exit Sub
		Call Draw_PFPSDM()
	End Sub

	Private Sub cmdExcel_Click()
		PFPSDM_Excel = True
		CPHSDM_Excel = False
		frmExcelCurves.ShowDialog()
	End Sub


	Private Sub cmdExit_Click()
		Me.Close()
	End Sub


	Private Sub cmdFile_Click()
		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNOverwritePrompt As Object
		Dim f, Error_Code As Short
		Dim temp As String
		Dim J, i, k As Short
		Dim Eq1, Filename_PFPSDM As String

		On Error GoTo File_Error
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.CancelError = True

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FileName = ""
		SaveFileDialog1.FileName = ""

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.DialogTitle = "Print to File"
		SaveFileDialog1.Title = "Print to File"

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
		SaveFileDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FilterIndex = 2
		SaveFileDialog1.FilterIndex = 2


		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNOverwritePrompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Action = 2
		SaveFileDialog1.ShowDialog()

		'f = FileNameIsValid(Filename_PFPSDM, CMDialog1)
		'If Not (f) Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Filename_PFPSDM = CMDialog1.FileName
		Filename_PFPSDM = SaveFileDialog1.FileName

		f = FreeFile()
		FileOpen(f, Filename_PFPSDM, OpenMode.Output)
		PrintLine(f, "Input data for the Plug-Flow Pore And Surface Diffusion Model")
		'-- Print Filename

		PrintLine(f)
		PrintLine(f, "From Data File :", Filename)

		PrintLine(f)
		PrintLine(f, "Component", TAB(30), "K*", TAB(38), "1/n", TAB(47), "C0", TAB(57), "MW", TAB(65), "Vm", TAB(75), "NBP")
		PrintLine(f, TAB(39), "-", TAB(46), "mg/L", TAB(56), "g/mol", TAB(65), "cm" & Chr(179) & "/mol", TAB(76), "C")

		For i = 1 To Number_Component_PFPSDM
			PrintLine(f, Trim(Mid(LTrim(Results.Component(i).Name), 1, 25)), TAB(29), VB6.Format(Results.Component(i).Use_K, "###,##0.000"), TAB(37), VB6.Format(Results.Component(i).Use_OneOverN, "0.000"), TAB(46), Format_It(Results.Component(i).InitialConcentration, 2), TAB(55), Format_It(Results.Component(i).MW, 2), TAB(64), Format_It(Results.Component(i).MolarVolume, 2), TAB(73), Format_It(Results.Component(i).BP, 2))
		Next i
		PrintLine(f)
		PrintLine(f, "* K in (mg/g)*(L/mg)^(1/n) - Vm = Molar Volume at NBP")
		PrintLine(f)

		'-----------------------Bed Data ----------------------
		PrintLine(f, "Bed Data:")

		PrintLine(f, TAB(5), "Bed Length: ", TAB(28), VB6.Format(Results.Bed.length, "0.000E+00") & " m")
		PrintLine(f, TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(Results.Bed.Diameter, "0.000E+00") & " m")
		PrintLine(f, TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(Results.Bed.Weight, "0.000E+00") & " kg")
		PrintLine(f, TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(Results.Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		PrintLine(f, TAB(5), "EBCT: ", TAB(28), VB6.Format(Results.Bed.length * PI * Results.Bed.Diameter * Results.Bed.Diameter / 4.0# / Results.Bed.Flowrate / 60.0#, "0.000E+00") & " mn")
		PrintLine(f)
		PrintLine(f, TAB(5), "Temperature:", TAB(28), VB6.Format(Results.Bed.Temperature, "0.00") & " C")
		If Results.Bed.Phase = 0 Then
			PrintLine(f, TAB(5), "Water Density:", TAB(28), VB6.Format(Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			PrintLine(f, TAB(5), "Water Viscosity:", TAB(28), VB6.Format(Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			PrintLine(f, TAB(5), "Pressure:", TAB(28), VB6.Format(Results.Bed.Pressure, "0.00000") & " atm")
			PrintLine(f, TAB(5), "Air Density:", TAB(28), VB6.Format(Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			PrintLine(f, TAB(5), "Air Viscosity:", TAB(28), VB6.Format(Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		PrintLine(f)

		'-----------------Carbon Properties -------------------------------
		PrintLine(f, "Carbon Properties:")
		PrintLine(f, TAB(5), "Name: ", TAB(28), Trim(Results.Carbon.Name))
		PrintLine(f, TAB(5), "Apparent Density: ", TAB(28), VB6.Format(Results.Carbon.Density, "0.000") & " g/cm" & Chr(179))
		PrintLine(f, TAB(5), "Particle Radius: ", TAB(28), VB6.Format(Results.Carbon.ParticleRadius * 100.0#, "0.000000") & " cm")
		PrintLine(f, TAB(5), "Porosity: ", TAB(28), VB6.Format(Results.Carbon.Porosity, "0.000"))
		PrintLine(f, TAB(5), "Shape Factor: ", TAB(28), VB6.Format(Results.Carbon.ShapeFactor, "0.000"))
		'Print #f, Tab(5); "Tortuosity: "; Tab(28); Format$(Results.Carbon.Tortuosity, "0.000")
		PrintLine(f)

		'---------------Kinetic Parameters -----------------------------------------
		PrintLine(f, "Kinetic parameters:")
		PrintLine(f)
		PrintLine(f, "Component", TAB(24), "kf", TAB(31), "Ds", TAB(40), "Dp", TAB(49), "St", TAB(58), "Eds", TAB(67), "Edp", TAB(75), "SPDFR")
		PrintLine(f, TAB(23), "cm/s", TAB(32), "cm" & Chr(178) & "/s", TAB(41), "cm" & Chr(178) & "/s", TAB(50), "-", TAB(59), "-", TAB(68), "-", TAB(76), "-")
		For i = 1 To Number_Component_PFPSDM
			'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Component(0) = Results.Component(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Edp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Eds(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(f, Trim(Mid(LTrim(Results.Component(i).Name), 1, 20)), TAB(22), Format_It(Results.Component(i).kf, 2), TAB(35), Format_It(Results.Component(i).Ds, 2), TAB(44), Format_It(Results.Component(i).Dp, 2), TAB(54), Format_It(ST(0), 2), TAB(61), Format_It(Eds(0), 2), TAB(68), Format_It(Edp(0), 2), TAB(75), Format_It(Results.Component(i).SPDFR, 2))
		Next i
		PrintLine(f)

		'Fouling-----------------------------------------
		PrintLine(f, "Fouling correlations:")
		PrintLine(f)

		PrintLine(f, " Water type : " & Trim(Results.Bed.Water_Correlation.Name))
		Eq1 = VB6.Format(Results.Bed.Water_Correlation.Coeff(1), "0.00")

		If Results.Bed.Water_Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
		Else
			If Results.Bed.Water_Correlation.Coeff(2) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
			End If
		End If
		If Results.Bed.Water_Correlation.Coeff(3) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
		Else
			If Results.Bed.Water_Correlation.Coeff(3) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
			End If
		End If
		If Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
			If Results.Bed.Water_Correlation.Coeff(4) > 0 Then
				Eq1 = Eq1 & VB6.Format(Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
			Else
				If Results.Bed.Water_Correlation.Coeff(4) < 0 Then
					Eq1 = Eq1 & VB6.Format(Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
				End If
			End If
		End If
		PrintLine(f, "K(t)/K0 = " & Eq1)
		PrintLine(f, "(t in minutes)")
		PrintLine(f)

		For J = 1 To Number_Component_PFPSDM
			Eq1 = ""
			If Results.Component(J).Correlation.Coeff(1) = 1.0# Then
				Eq1 = "(K/K0) "
			Else
				If Results.Component(J).Correlation.Coeff(1) <> 0 Then Eq1 = VB6.Format(Results.Component(J).Correlation.Coeff(1), "0.00") & " * (K/K0) "
			End If
			If Results.Component(J).Correlation.Coeff(2) > 0 Then
				Eq1 = Eq1 & "+ " & VB6.Format(Results.Component(J).Correlation.Coeff(2), "0.00")
			Else
				If Results.Component(J).Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & VB6.Format(System.Math.Abs(Results.Component(J).Correlation.Coeff(2)), "0.00")
			End If
			If Trim(Eq1) = "" Then
				Eq1 = "K/K0"
			End If
			PrintLine(f, Trim(Results.Component(J).Name) & ":")
			PrintLine(f, TAB(10), "Correlation type: " & Trim(Results.Component(J).Correlation.Name))
			PrintLine(f, TAB(10), "K/K0 = " & Eq1)

			If (Results.Component(J).Use_Tortuosity_Correlation) Then
				If (Results.Component(J).Constant_Tortuosity) Then
					PrintLine(f, "Correlation used when SOC competition is important:")
					PrintLine(f, " Tortuosity = 0.782 * EBCT^0.925 ")
				Else
					PrintLine(f, "Correlation used when NOM fouling is important:")
					PrintLine(f, " Tortuosity = 1.0 if t< 70 days")
					PrintLine(f, " Tortuosity = 0.334 + 6.610E-06 * EBCT")
				End If
			End If

			PrintLine(f)
		Next J

		'If Results.Use_Tortuosity_Correlation Then
		'  If Results.Constant_Tortuosity Then
		'    Print #f, "Correlation used when SOC competition is important:"
		'    Print #f, " Tortuosity = 0.782 * EBCT^0.925 "
		'  Else
		'    Print #f, "Correlation used when NOM fouling is important:"
		'    Print #f, " Tortuosity = 1.0 if t< 70 days"
		'    Print #f, " Tortuosity = 0.334 + 6.610E-06 * EBCT"
		'  End If
		'End If
		PrintLine(f)

		'--- Print the results from the table
		PrintLine(f, "Results for the Plug-Flow Pore And Surface Diffusion Model")
		PrintLine(f)
		For i = 1 To Results.NComponent
			PrintLine(f, Results.Component(i).Name)
			PrintLine(f, TAB(30), "Time(days)", TAB(40), "BVT", TAB(50), "TC", TAB(60), "C (mg/L)")
			If (Results.ThroughPut_05(i).T <> -1.0#) And (Results.ThroughPut_05(i).C <> -1.0#) Then
				PrintLine(f, "5% of the influent conc.", TAB(30), Format_It(Results.ThroughPut_05(i).T / 24.0# / 60.0#, 2), TAB(40), Format_It(Results.ThroughPut_05(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(Results.ThroughPut_05(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2), TAB(60), Format_It(Results.ThroughPut_05(i).C, 2))
			Else
				PrintLine(f, "5% of the influent conc.", TAB(30), "N/A", TAB(40), "N/A", TAB(50), "N/A", TAB(60), "N/A")
			End If

			If (Results.ThroughPut_50(i).T <> -1.0#) And (Results.ThroughPut_50(i).C <> -1.0#) Then
				PrintLine(f, "50% of the influent conc.", TAB(30), Format_It(Results.ThroughPut_50(i).T / 24.0# / 60.0#, 2), TAB(40), Format_It(Results.ThroughPut_50(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(Results.ThroughPut_50(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2), TAB(60), Format_It(Results.ThroughPut_50(i).C, 2))
			Else
				PrintLine(f, "50% of the influent conc.", TAB(30), "N/A", TAB(40), "N/A", TAB(50), "N/A", TAB(60), "N/A")
			End If

			If (Results.ThroughPut_95(i).T <> -1.0#) And (Results.ThroughPut_95(i).C <> -1.0#) Then
				PrintLine(f, "95% of the influent conc.", TAB(30), Format_It(Results.ThroughPut_95(i).T / 24.0# / 60.0#, 2), TAB(40), Format_It(Results.ThroughPut_95(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(Results.ThroughPut_95(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2), TAB(60), Format_It(Results.ThroughPut_95(i).C, 2))
			Else
				PrintLine(f, "95% of the influent conc.", TAB(30), "N/A", TAB(40), "N/A", TAB(50), "N/A", TAB(60), "N/A")
			End If
			PrintLine(f)
			If Flag_TO(i) Then
				PrintLine(f, "Treatment Objective: " & Format_It(Treatment_Objective(i).C, 2) & " mg/L")
				PrintLine(f)
				PrintLine(f, TAB(10), "Time (days):", TAB(25), Format_It(Treatment_Objective(i).T / 60.0# / 24.0#, 2))
				PrintLine(f, TAB(10), "BVT:", TAB(25), Format_It(Treatment_Objective(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2))
				PrintLine(f, TAB(10), "Tr. Capacity:", TAB(25), Format_It(Treatment_Objective(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2))
			Else
				PrintLine(f, "The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective(i).C, 2) & "mg/L) could not be calculated.")
			End If
			PrintLine(f)
		Next i
		PrintLine(f, "TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")

		'--- Print PSDM inputs/calculations that were returned from the FORTRAN routine.
		PrintLine(f)
		PrintLine(f, "PSDM Module Input Variables")
		PrintLine(f, "Note: * designates a variable calculated in Visual BASIC")
		PrintLine(f)

		PrintLine(f, "Number of radial collocation points, NC            = " & VB6.Format(PSDM_Inputs.VARS1(1), "0"))
		PrintLine(f, "Number of axial collocation points, MC             = " & VB6.Format(PSDM_Inputs.VARS1(2), "0"))
		PrintLine(f, "Total no. of differential equations, NEQ           = " & VB6.Format(PSDM_Inputs.VARS1(3), "0"))
		PrintLine(f, "Radius of adsorbent particle, RAD (cm)             = " & VB6.Format(PSDM_Inputs.VARS1(4), "0.0000E+00"))
		PrintLine(f, "Apparent particle density, RHOP (g/cm^3)           = " & VB6.Format(PSDM_Inputs.VARS1(5), "0.0000E+00"))
		PrintLine(f, "Void fraction of carbon, EPOR (-)                  = " & VB6.Format(PSDM_Inputs.VARS1(6), "0.0000E+00"))
		PrintLine(f, "Void fraction of bed, EBED (-)                     = " & VB6.Format(PSDM_Inputs.VARS1(7), "0.0000E+00"))
		PrintLine(f, "*Surface loading, SF (gpm/ft^2)                    = " & VB6.Format(PSDM_Inputs.VARS1(8), "0.0000E+00"))
		PrintLine(f, "Packed bed contact time, TAU (sec)                 = " & VB6.Format(PSDM_Inputs.VARS1(9), "0.0000E+00"))
		PrintLine(f, "Empty bed contact time, EBCT (min)                 = " & VB6.Format(PSDM_Inputs.VARS1(10), "0.0000E+00"))
		PrintLine(f, "*Reynolds number, RE (-)                           = " & VB6.Format(PSDM_Inputs.VARS1(11), "0.0000E+00"))
		PrintLine(f, "*Fluid density, DW (g/cm^3)                        = " & VB6.Format(PSDM_Inputs.VARS1(12), "0.0000E+00"))
		PrintLine(f, "*Fluid viscosity, VW (g/cm-s)                      = " & VB6.Format(PSDM_Inputs.VARS1(13), "0.0000E+00"))
		PrintLine(f, "Error flag, NFLAG                                  = " & VB6.Format(PSDM_Inputs.VARS1(15), "0"))
		PrintLine(f)

		For i = 1 To Results.NComponent
			PrintLine(f, Results.Component(i).Name)
			PrintLine(f, "Molal volume at the boiling pt., VB (cm^3/gmol)    = " & VB6.Format(PSDM_Inputs.VARS2(i, 1), "0.0000E+00"))
			PrintLine(f, "Molecular weight of compound, XWT (g/gmol)         = " & VB6.Format(PSDM_Inputs.VARS2(i, 2), "0.0000E+00"))
			PrintLine(f, "Initial bulk liquid-phase conc., CBO (umol/L)      = " & VB6.Format(PSDM_Inputs.VARS2(i, 3), "0.0000E+00"))
			PrintLine(f, "Freundlich iso. cap., XK (umol/g)*(L/umol)^(1/n)   = " & VB6.Format(PSDM_Inputs.VARS2(i, 4), "0.0000E+00"))
			PrintLine(f, "Freundlich isotherm constant, XN (-)               = " & VB6.Format(PSDM_Inputs.VARS2(i, 5), "0.0000E+00"))
			PrintLine(f, "*Liquid diffusivity, DIFL (cm^2/sec)               = " & VB6.Format(PSDM_Inputs.VARS2(i, 6), "0.0000E+00"))
			PrintLine(f, "Film transfer coefficient, KF (cm/sec)             = " & VB6.Format(PSDM_Inputs.VARS2(i, 7), "0.0000E+00"))
			PrintLine(f, "Surface diffusion coefficient, DS (cm^2/s)         = " & VB6.Format(PSDM_Inputs.VARS2(i, 8), "0.0000E+00"))
			PrintLine(f, "Stanton number, ST (-)                             = " & VB6.Format(PSDM_Inputs.VARS2(i, 9), "0.0000E+00"))
			PrintLine(f, "Solute distribution parameter, DGS (-)             = " & VB6.Format(PSDM_Inputs.VARS2(i, 10), "0.0000E+00"))
			PrintLine(f, "Biot number, BIS (-)                               = " & VB6.Format(PSDM_Inputs.VARS2(i, 11), "0.0000E+00"))
			PrintLine(f, "Diffusivity modulus, EDS (-)                       = " & VB6.Format(PSDM_Inputs.VARS2(i, 12), "0.0000E+00"))
			PrintLine(f, "Pore solute dist. parameter, DGP (-)               = " & VB6.Format(PSDM_Inputs.VARS2(i, 13), "0.0000E+00"))
			PrintLine(f, "Pore diffusion coefficient, DP (cm^2/s)            = " & VB6.Format(PSDM_Inputs.VARS2(i, 14), "0.0000E+00"))
			PrintLine(f, "Pore Biot number, BIP (-)                          = " & VB6.Format(PSDM_Inputs.VARS2(i, 15), "0.0000E+00"))
			PrintLine(f, "Pore diffusion modulus, EDP (-)                    = " & VB6.Format(PSDM_Inputs.VARS2(i, 16), "0.0000E+00"))
			PrintLine(f, "Surface to pore diffusivity ratio, D (-)           = " & VB6.Format(PSDM_Inputs.VARS2(i, 17), "0.0000E+00"))
			PrintLine(f, "*Schmidt number, SC (-)                            = " & VB6.Format(PSDM_Inputs.VARS2(i, 18), "0.0000E+00"))
			PrintLine(f, "*SPDFR (-)                                         = " & VB6.Format(PSDM_Inputs.VARS2(i, 19), "0.0000E+00"))
			PrintLine(f)
		Next i

		FileClose((f))

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FileName = ""
		SaveFileDialog1.FileName = ""
		Exit Sub

File_Error:
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = 75) Then
			'DO NOTHING.
		Else
			Call Show_Trapped_Error("cmdFile_Click")
		End If
		Resume Exit_Print_File
Exit_Print_File:
	End Sub


	Private Sub cmdPrint_Click()
		Dim Printer As New Printer

		Dim Error_Code As Short
		Dim temp As String
		Dim i As Short
		Dim H, W As Single
		Dim Eq1, MTZ As String
		Dim J As Short
		Dim f As Double

		On Error GoTo Print_Error

		'---- Print the graph ------------------------
		'	'''    For i = 1 To Number_Component
		'		'''      grpBreak.ThisPoint = i
		'		'''      grpBreak.PatternData = i - 1
		'		'''    Next i
		'		'''
		'		'''    H = grpBreak.Height
		'		'''    W = grpBreak.Width
		'		'''
		'		'''    grpBreak.Visible = False 'Hide it before printing

		'
		' THIS CODE HAD TO BE REPLACED TODAY, 1999-MAY-11, EJOMAN.
		'
		'---- OLD CODE STARTS HERE:
		'If Printer.Width < Printer.Height Then
		'  grpBreak.Height = CSng(Printer.Height / 2#)
		'  grpBreak.Width = Printer.Width
		'Else
		'  grpBreak.Height = Printer.Height
		'  grpBreak.Width = Printer.Width
		'End If
		'---- OLD CODE ENDS.
		'
		'---- NEW CODE STARTS HERE:
		'	'''    If Printer.Width < Printer.Height Then
		'	'''      grpBreak.Height = CDbl(Printer.ScaleHeight) * 0.5
		'	'''      grpBreak.Width = CDbl(Printer.ScaleWidth) * 0.75
		'	'''    Else
		'	'''      grpBreak.Height = CDbl(Printer.ScaleHeight) * 0.75
		'	'''      grpBreak.Width = CDbl(Printer.ScaleWidth) * 0.75
		'	'''    End If
		'---- NEW CODE ENDS.

		'
		' THE PRINTING CODE HAD TO BE REPLACED TODAY, 1999-MAY-11, EJOMAN.
		' REFER TO www.microsoft.com KNOWLEDGE BASE ARTICLE #Q150222.
		'
		'---- OLD CODE STARTS HERE:
		'grpBreak.PrintStyle = 2
		'grpBreak.DrawMode = 5
		'---- OLD CODE ENDS.
		'
		'---- NEW CODE STARTS HERE:
		'	'''    Printer.ScaleLeft = -((Printer.Width - grpBreak.Width) / 2)
		'	'''    Printer.ScaleTop = -((Printer.Height - grpBreak.Height) / 2)
		'	'''    Printer.PaintPicture _
		''''        grpBreak.Picture, _
		''''        0, _
		''''        0, _
		''''        grpBreak.Width, _
		''''        grpBreak.Height
		''''    Printer.Line _
		''''        (0, 0)- _
		''''        (grpBreak.Width, grpBreak.Height), _
		''''        QBColor(0), _
		''''        B
		'---- NEW CODE ENDS.

		'	'''    grpBreak.Height = H
		'	'''    grpBreak.Width = W
		'		'''
		'		'''    grpBreak.Visible = True
		'		'''
		'		'''    grpBreak.PrintStyle = 2
		'		'''    grpBreak.DrawMode = 2

		'
		' A "SKIP TO NEXT PAGE" COMMAND HAD TO BE ADDED TO THE PRINTING
		' CODE TODAY, 1999-MAY-11, EJOMAN.
		'
		'---- NEW CODE STARTS HERE:
		'	'''    Printer.NewPage
		'---- NEW CODE ENDS.

		'Print other results-----------------------------------------------
		Printer.ScaleLeft = -1080 'Set a 3/4-inch margin
		Printer.ScaleTop = -1080
		Printer.CurrentX = 0
		Printer.CurrentY = 0

		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Input data for the Plug-Flow Pore And Surface Diffusion Model")
		Printer.FontSize = 10
		Printer.FontBold = False
		Printer.FontUnderline = False
		'-- Print Filename
		Printer.Print()
		Printer.Print("From Data File: " & Filename)


		Printer.Print()
		Printer.Print("Component", TAB(30), "K*", TAB(38), "1/n", TAB(47), "C0", TAB(57), "MW", TAB(65), "Vm", TAB(75), "NBP")
		Printer.Print(TAB(39), "-", TAB(46), "mg/L", TAB(56), "g/mol", TAB(65), "cm" & Chr(179) & "/mol", TAB(76), "C")

		For i = 1 To Number_Component_PFPSDM
			Printer.Print(Trim(Mid(LTrim(Results.Component(i).Name), 1, 25)), TAB(29), VB6.Format(Results.Component(i).Use_K, "###,##0.000"), TAB(37), VB6.Format(Results.Component(i).Use_OneOverN, "0.000"), TAB(46), Format_It(Results.Component(i).InitialConcentration, 2), TAB(55), Format_It(Results.Component(i).MW, 2), TAB(64), Format_It(Results.Component(i).MolarVolume, 2), TAB(73), Format_It(Results.Component(i).BP, 2))
		Next i
		Printer.Print()
		Printer.Print("* K in (mg/g)*(L/mg)^(1/n) - Vm = Molar Volume at NBP")
		Printer.Print()

		'-----------------------Bed Data ----------------------
		Printer.FontUnderline = True
		Printer.Print("Bed Data:")
		Printer.FontUnderline = False

		Printer.Print(TAB(5), "Bed Length: ", TAB(28), VB6.Format(Results.Bed.length, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(Results.Bed.Diameter, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(Results.Bed.Weight, "0.000E+00") & " kg")
		Printer.Print(TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(Results.Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		Printer.Print(TAB(5), "EBCT: ", TAB(28), VB6.Format(Results.Bed.length * PI * Results.Bed.Diameter * Results.Bed.Diameter / 4.0# / Results.Bed.Flowrate / 60.0#, "0.000E+00") & " mn")
		Printer.Print()
		Printer.Print(TAB(5), "Temperature:", TAB(28), VB6.Format(Results.Bed.Temperature, "0.00") & " C")
		If Results.Bed.Phase = 0 Then
			Printer.Print(TAB(5), "Water Density:", TAB(28), VB6.Format(Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Water Viscosity:", TAB(28), VB6.Format(Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			Printer.Print(TAB(5), "Pressure:", TAB(28), VB6.Format(Results.Bed.Pressure, "0.00000") & " atm")
			Printer.Print(TAB(5), "Air Density:", TAB(28), VB6.Format(Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Air Viscosity:", TAB(28), VB6.Format(Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		Printer.Print()

		'-----------------Carbon Properties -------------------------------
		Printer.FontUnderline = True
		Printer.Print("Carbon Properties:")
		Printer.FontUnderline = False

		Printer.Print(TAB(5), "Name: ", TAB(28), Trim(Results.Carbon.Name))
		Printer.Print(TAB(5), "Apparent Density: ", TAB(28), VB6.Format(Results.Carbon.Density, "0.000") & " g/cm" & Chr(179))
		Printer.Print(TAB(5), "Particle Radius: ", TAB(28), VB6.Format(Results.Carbon.ParticleRadius * 100.0#, "0.000000") & " cm")
		Printer.Print(TAB(5), "Porosity: ", TAB(28), VB6.Format(Results.Carbon.Porosity, "0.000"))
		Printer.Print(TAB(5), "Shape Factor: ", TAB(28), VB6.Format(Results.Carbon.ShapeFactor, "0.000"))

		'Printer.Print Tab(5); "Tortuosity: "; Tab(28); Format$(Results.Carbon.Tortuosity, "0.000")
		Printer.Print()

		'---------------Kinetic Parameters -----------------------------------------
		Printer.FontUnderline = True
		Printer.Print("Kinetic parameters:")
		Printer.FontUnderline = False

		Printer.Print()
		Printer.Print("Component", TAB(15), "kf", TAB(22), "Ds", TAB(29), "Dp", TAB(36), "St", TAB(43), "Eds", TAB(50), "Edp", TAB(57), "SPDFR")
		Printer.Print(TAB(15), "cm/s", TAB(22), "cm" & Chr(178) & "/s", TAB(29), "cm" & Chr(178) & "/s", TAB(36), "-", TAB(43), "-", TAB(50), "-", TAB(57), "-")
		For i = 1 To Number_Component_PFPSDM
			'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Component(0) = Results.Component(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Edp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Eds(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Printer.Print(Trim(Mid(LTrim(Results.Component(i).Name), 1, 20)), TAB(15), Format_It(Results.Component(i).kf, 2), TAB(22), Format_It(Results.Component(i).Ds, 2), TAB(29), Format_It(Results.Component(i).Dp, 2), TAB(36), Format_It(ST(0), 2), TAB(43), Format_It(Eds(0), 2), TAB(50), Format_It(Edp(0), 2), TAB(57), Format_It(Results.Component(i).SPDFR, 2))
		Next i


		Printer.Print()

		'Fouling-----------------------------------------
		Printer.FontUnderline = True
		Printer.Print("Fouling correlations:")
		Printer.FontUnderline = False
		Printer.Print()

		Printer.Print(" Water type : " & Trim(Results.Bed.Water_Correlation.Name))
		Eq1 = VB6.Format(Results.Bed.Water_Correlation.Coeff(1), "0.00")

		If Results.Bed.Water_Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
		Else
			If Results.Bed.Water_Correlation.Coeff(2) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
			End If
		End If
		If Results.Bed.Water_Correlation.Coeff(3) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
		Else
			If Results.Bed.Water_Correlation.Coeff(3) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
			End If
		End If
		If Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
			If Results.Bed.Water_Correlation.Coeff(4) > 0 Then
				Eq1 = Eq1 & VB6.Format(Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
			Else
				If Results.Bed.Water_Correlation.Coeff(4) < 0 Then
					Eq1 = Eq1 & VB6.Format(Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
				End If
			End If
		End If
		Printer.Print("K(t)/K0 = " & Eq1)
		Printer.Print("(t in minutes)")
		Printer.Print()

		For J = 1 To Number_Component_PFPSDM
			Eq1 = ""
			If Results.Component(J).Correlation.Coeff(1) = 1.0# Then
				Eq1 = "(K/K0) "
			Else
				If Results.Component(J).Correlation.Coeff(1) <> 0 Then Eq1 = VB6.Format(Results.Component(J).Correlation.Coeff(1), "0.00") & " * (K/K0) "
			End If
			If Results.Component(J).Correlation.Coeff(2) > 0 Then
				Eq1 = Eq1 & "+ " & VB6.Format(Results.Component(J).Correlation.Coeff(2), "0.00")
			Else
				If Results.Component(J).Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & VB6.Format(System.Math.Abs(Results.Component(J).Correlation.Coeff(2)), "0.00")
			End If
			If Trim(Eq1) = "" Then
				Eq1 = "K/K0"
			End If
			Printer.Print(Trim(Results.Component(J).Name) & ":")
			Printer.Print(TAB(10), "Correlation type: " & Trim(Results.Component(J).Correlation.Name))

			Printer.Print(TAB(10), "K/K0 = " & Eq1)

			If (Results.Component(J).Use_Tortuosity_Correlation) Then
				If (Results.Component(J).Constant_Tortuosity) Then
					Printer.Print("Correlation used when SOC competition is important:")
					Printer.Print(" Tortuosity = 0.782 * EBCT^0.925 ")
				Else
					Printer.Print("Correlation used when NOM fouling is important:")
					Printer.Print(" Tortuosity = 1.0 if t< 70 days")
					Printer.Print(" Tortuosity = 0.334 + 6.610E-06 * EBCT")
				End If
			End If
			Printer.Print()

		Next J

		'If Results.Use_Tortuosity_Correlation Then
		'  If Results.Constant_Tortuosity Then
		'    Printer.Print "Correlation used when SOC competition is important:"
		'    Printer.Print " Tortuosity = 0.782 * EBCT^0.925 "
		'  Else
		'    Printer.Print "Correlation used when NOM fouling is important:"
		'    Printer.Print " Tortuosity = 1.0 if t< 70 days"
		'    Printer.Print " Tortuosity = 0.334 + 6.610E-06 * EBCT"
		'  End If
		'End If
		Printer.Print()

		'Model Parameters
		Printer.FontSize = 10
		Printer.FontBold = False
		Printer.FontUnderline = True
		Printer.Print("Model Parameters")
		Printer.FontUnderline = False
		Printer.FontSize = 10
		Printer.Print(TAB(5), "Total Run Time:", TAB(50), VB6.Format(TimeP.End_Renamed / 24.0# / 60.0#, "0.000") & " days")
		If (TimeP.Init / 60.0# / 24.0#) > 0.001 Then
			Printer.Print(TAB(5), "First Point Displayed:", TAB(50), VB6.Format(TimeP.Init / 24.0# / 60.0#, "0.000") & " days")
		Else
			Printer.Print(TAB(5), "First Point Displayed:", TAB(50), VB6.Format(TimeP.Init / 24.0# / 60.0#, "0.000E+00") & " days")
		End If
		Printer.Print(TAB(5), "Time Step:", TAB(50), VB6.Format(TimeP.Step_Renamed / 24.0# / 60.0#, "0.000") & " days")
		Printer.Print(TAB(5), "Number of Axial Collocation Points:", TAB(50), VB6.Format(MC, "0"))
		Printer.Print(TAB(5), "Number of Radial Collocation Points:", TAB(50), VB6.Format(NC, "0"))
		Printer.Print(TAB(5), "Number of Axial Elements:", TAB(50), VB6.Format(Bed.NumberOfBeds, "0"))

		'
		' PAGE BREAK REMOVED ON 1999-MAY-11, EJOMAN.
		'Printer.NewPage

		'--- Print the results from the table
		Printer.FontSize = 10
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Results for the Plug-Flow Pore And Surface Diffusion Model")
		Printer.FontUnderline = False
		Printer.Print()
		For i = 1 To Results.NComponent
			Printer.FontSize = 12
			Printer.FontBold = True
			Printer.Print(Results.Component(i).Name)
			Printer.FontSize = 10
			Printer.FontBold = False
			Printer.Print(TAB(30), "Time(days)", TAB(40), "BVT", TAB(50), "TC", TAB(60), "C (mg/L)")
			If (Results.ThroughPut_05(i).T <> -1.0#) And (Results.ThroughPut_05(i).C <> -1.0#) Then
				Printer.Print("5% of the influent conc.", TAB(30), Format_It(Results.ThroughPut_05(i).T / 24.0# / 60.0#, 2), TAB(40), Format_It(Results.ThroughPut_05(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(Results.ThroughPut_05(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2), TAB(60), Format_It(Results.ThroughPut_05(i).C, 2))
			Else
				Printer.Print("5% of the influent conc.", TAB(30), "N/A", TAB(40), "N/A", TAB(50), "N/A", TAB(60), "N/A")
			End If

			If (Results.ThroughPut_50(i).T <> -1.0#) And (Results.ThroughPut_50(i).C <> -1.0#) Then
				Printer.Print("50% of the influent conc.", TAB(30), Format_It(Results.ThroughPut_50(i).T / 24.0# / 60.0#, 2), TAB(40), Format_It(Results.ThroughPut_50(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(Results.ThroughPut_50(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2), TAB(60), Format_It(Results.ThroughPut_50(i).C, 2))
			Else
				Printer.Print("50% of the influent conc.", TAB(30), "N/A", TAB(40), "N/A", TAB(50), "N/A", TAB(60), "N/A")
			End If

			If (Results.ThroughPut_95(i).T <> -1.0#) And (Results.ThroughPut_95(i).C <> -1.0#) Then
				Printer.Print("95% of the influent conc.", TAB(30), Format_It(Results.ThroughPut_95(i).T / 24.0# / 60.0#, 2), TAB(40), Format_It(Results.ThroughPut_95(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(Results.ThroughPut_95(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2), TAB(60), Format_It(Results.ThroughPut_95(i).C, 2))
			Else
				Printer.Print("95% of the influent conc.", TAB(30), "N/A", TAB(40), "N/A", TAB(50), "N/A", TAB(60), "N/A")
			End If
			Printer.Print()
			If (Results.ThroughPut_50(i).C <> -1.0#) And (Results.ThroughPut_50(i).T <> -1.0#) And (Results.ThroughPut_05(i).T <> -1.0#) And (Results.ThroughPut_05(i).C <> -1.0#) And (Results.ThroughPut_95(i).T <> -1.0#) And (Results.ThroughPut_95(i).C <> -1.0#) Then
				f = 100.0# * Results.Bed.length / Results.ThroughPut_50(i).T 'in cm/dayss
				MTZ = VB6.Format(f * (Results.ThroughPut_95(i).T - Results.ThroughPut_05(i).T), "0.00E+00")
			Else
				MTZ = "N/A"
			End If
			Printer.Print()
			Printer.Print("MTZ Length 5%-95% (cm) :" & MTZ)
			If Flag_TO(i) Then
				Printer.FontUnderline = True
				Printer.Print("Treatment Objective: " & Format_It(Treatment_Objective(i).C, 2) & " mg/L")
				Printer.FontUnderline = False
				Printer.Print()
				Printer.Print(TAB(10), "Time (days):", TAB(25), Format_It(Treatment_Objective(i).T / 60.0# / 24.0#, 2))
				Printer.Print(TAB(10), "BVT:", TAB(25), Format_It(Treatment_Objective(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2))
				Printer.Print(TAB(10), "Tr. Capacity:", TAB(25), Format_It(Treatment_Objective(i).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2))
			Else
				Printer.Print("The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective(i).C, 2) & "mg/L) could not be calculated.")
			End If
			Printer.Print()
		Next i
		Printer.Print("TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")

		'--- Print PSDM inputs/calculations that were returned from the FORTRAN routine.
		Printer.Print()
		Printer.Print("PSDM Module Input Variables")
		Printer.Print("Note: * designates a variable calculated in Visual BASIC")
		Printer.Print()

		'---- OLD CODE MODIFIED 1999-MAY-11 (EJOMAN) STARTS HERE:
		'Printer.Print "Number of radial collocation points, NC            = " & Format$(PSDM_Inputs.VARS1(1), "0")
		'Printer.Print "Number of axial collocation points, MC             = " & Format$(PSDM_Inputs.VARS1(2), "0")
		'Printer.Print "Total no. of differential equations, NEQ           = " & Format$(PSDM_Inputs.VARS1(3), "0")
		'Printer.Print "Radius of adsorbent particle, RAD (cm)             = " & Format$(PSDM_Inputs.VARS1(4), "0.0000E+00")
		'Printer.Print "Apparent particle density, RHOP (g/cm^3)           = " & Format$(PSDM_Inputs.VARS1(5), "0.0000E+00")
		'Printer.Print "Void fraction of carbon, EPOR (-)                  = " & Format$(PSDM_Inputs.VARS1(6), "0.0000E+00")
		'Printer.Print "Void fraction of bed, EBED (-)                     = " & Format$(PSDM_Inputs.VARS1(7), "0.0000E+00")
		'Printer.Print "*Surface loading, SF (gpm/ft^2)                    = " & Format$(PSDM_Inputs.VARS1(8), "0.0000E+00")
		'Printer.Print "Packed bed contact time, TAU (sec)                 = " & Format$(PSDM_Inputs.VARS1(9), "0.0000E+00")
		'Printer.Print "Empty bed contact time, EBCT (min)                 = " & Format$(PSDM_Inputs.VARS1(10), "0.0000E+00")
		'Printer.Print "*Reynolds number, RE (-)                           = " & Format$(PSDM_Inputs.VARS1(11), "0.0000E+00")
		'Printer.Print "*Fluid density, DW (g/cm^3)                        = " & Format$(PSDM_Inputs.VARS1(12), "0.0000E+00")
		'Printer.Print "*Fluid viscosity, VW (g/cm-s)                      = " & Format$(PSDM_Inputs.VARS1(13), "0.0000E+00")
		'Printer.Print "Error flag, NFLAG                                  = " & Format$(PSDM_Inputs.VARS1(15), "0")
		'Printer.Print
		'For i = 1 To Results.NComponent
		'  Printer.Print Results.Component(i).Name
		'  Printer.Print "Molal volume at the boiling pt., VB (cm^3/gmol)    = " & Format$(PSDM_Inputs.VARS2(i, 1), "0.0000E+00")
		'  Printer.Print "Molecular weight of compound, XWT (g/gmol)         = " & Format$(PSDM_Inputs.VARS2(i, 2), "0.0000E+00")
		'  Printer.Print "Initial bulk liquid-phase conc., CBO (umol/L)      = " & Format$(PSDM_Inputs.VARS2(i, 3), "0.0000E+00")
		'  Printer.Print "Freundlich iso. cap., XK (umol/g)*(L/umol)^(1/n)   = " & Format$(PSDM_Inputs.VARS2(i, 4), "0.0000E+00")
		'  Printer.Print "Freundlich isotherm constant, XN (-)               = " & Format$(PSDM_Inputs.VARS2(i, 5), "0.0000E+00")
		'  Printer.Print "*Liquid diffusivity, DIFL (cm^2/sec)               = " & Format$(PSDM_Inputs.VARS2(i, 6), "0.0000E+00")
		'  Printer.Print "Film transfer coefficient, KF (cm/sec)             = " & Format$(PSDM_Inputs.VARS2(i, 7), "0.0000E+00")
		'  Printer.Print "Surface diffusion coefficient, DS (cm^2/s)         = " & Format$(PSDM_Inputs.VARS2(i, 8), "0.0000E+00")
		'  Printer.Print "Stanton number, ST (-)                             = " & Format$(PSDM_Inputs.VARS2(i, 9), "0.0000E+00")
		'  Printer.Print "Solute distribution parameter, DGS (-)             = " & Format$(PSDM_Inputs.VARS2(i, 10), "0.0000E+00")
		'  Printer.Print "Biot number, BIS (-)                               = " & Format$(PSDM_Inputs.VARS2(i, 11), "0.0000E+00")
		'  Printer.Print "Diffusivity modulus, EDS (-)                       = " & Format$(PSDM_Inputs.VARS2(i, 12), "0.0000E+00")
		'  Printer.Print "Pore solute dist. parameter, DGP (-)               = " & Format$(PSDM_Inputs.VARS2(i, 13), "0.0000E+00")
		'  Printer.Print "Pore diffusion coefficient, DP (cm^2/s)            = " & Format$(PSDM_Inputs.VARS2(i, 14), "0.0000E+00")
		'  Printer.Print "Pore Biot number, BIP (-)                          = " & Format$(PSDM_Inputs.VARS2(i, 15), "0.0000E+00")
		'  Printer.Print "Pore diffusion modulus, EDP (-)                    = " & Format$(PSDM_Inputs.VARS2(i, 16), "0.0000E+00")
		'  Printer.Print "Surface to pore diffusivity ratio, D (-)           = " & Format$(PSDM_Inputs.VARS2(i, 17), "0.0000E+00")
		'  Printer.Print "*Schmidt number, SC (-)                            = " & Format$(PSDM_Inputs.VARS2(i, 18), "0.0000E+00")
		'  Printer.Print "*SPDFR (-)                                         = " & Format$(PSDM_Inputs.VARS2(i, 19), "0.0000E+00")
		'  Printer.Print
		'Next i
		'---- OLD CODE ENDS.
		'
		'---- NEW CODE MODIFIED 1999-MAY-11 (EJOMAN) STARTS HERE:
		Printer.Print("Number of radial collocation points, NC = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(1), "0"))
		Printer.Print("Number of axial collocation points, MC = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(2), "0"))
		Printer.Print("Total no. of differential equations, NEQ = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(3), "0"))
		Printer.Print("Radius of adsorbent particle, RAD (cm) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(4), "0.0000E+00"))
		Printer.Print("Apparent particle density, RHOP (g/cm^3) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(5), "0.0000E+00"))
		Printer.Print("Void fraction of carbon, EPOR (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(6), "0.0000E+00"))
		Printer.Print("Void fraction of bed, EBED (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(7), "0.0000E+00"))
		Printer.Print("*Surface loading, SF (gpm/ft^2) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(8), "0.0000E+00"))
		Printer.Print("Packed bed contact time, TAU (sec) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(9), "0.0000E+00"))
		Printer.Print("Empty bed contact time, EBCT (min) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(10), "0.0000E+00"))
		Printer.Print("*Reynolds number, RE (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(11), "0.0000E+00"))
		Printer.Print("*Fluid density, DW (g/cm^3) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(12), "0.0000E+00"))
		Printer.Print("*Fluid viscosity, VW (g/cm-s) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(13), "0.0000E+00"))
		Printer.Print("Error flag, NFLAG = ", TAB(70), VB6.Format(PSDM_Inputs.VARS1(15), "0"))
		Printer.Print()
		For i = 1 To Results.NComponent
			Printer.Print(Results.Component(i).Name)
			Printer.Print("Molal volume at the boiling pt., VB (cm^3/gmol) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 1), "0.0000E+00"))
			Printer.Print("Molecular weight of compound, XWT (g/gmol) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 2), "0.0000E+00"))
			Printer.Print("Initial bulk liquid-phase conc., CBO (umol/L) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 3), "0.0000E+00"))
			Printer.Print("Freundlich iso. cap., XK (umol/g)*(L/umol)^(1/n) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 4), "0.0000E+00"))
			Printer.Print("Freundlich isotherm constant, XN (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 5), "0.0000E+00"))
			Printer.Print("*Liquid diffusivity, DIFL (cm^2/sec) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 6), "0.0000E+00"))
			Printer.Print("Film transfer coefficient, KF (cm/sec) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 7), "0.0000E+00"))
			Printer.Print("Surface diffusion coefficient, DS (cm^2/s) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 8), "0.0000E+00"))
			Printer.Print("Stanton number, ST (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 9), "0.0000E+00"))
			Printer.Print("Solute distribution parameter, DGS (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 10), "0.0000E+00"))
			Printer.Print("Biot number, BIS (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 11), "0.0000E+00"))
			Printer.Print("Diffusivity modulus, EDS (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 12), "0.0000E+00"))
			Printer.Print("Pore solute dist. parameter, DGP (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 13), "0.0000E+00"))
			Printer.Print("Pore diffusion coefficient, DP (cm^2/s) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 14), "0.0000E+00"))
			Printer.Print("Pore Biot number, BIP (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 15), "0.0000E+00"))
			Printer.Print("Pore diffusion modulus, EDP (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 16), "0.0000E+00"))
			Printer.Print("Surface to pore diffusivity ratio, D (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 17), "0.0000E+00"))
			Printer.Print("*Schmidt number, SC (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 18), "0.0000E+00"))
			Printer.Print("*SPDFR (-) = ", TAB(70), VB6.Format(PSDM_Inputs.VARS2(i, 19), "0.0000E+00"))
			Printer.Print()
		Next i
		'---- NEW CODE ENDS.

		Printer.EndDoc()
		Exit Sub

Print_Error:
		Call Show_Trapped_Error("cmdPrint_Click")
		Resume Exit_Print
Exit_Print:

	End Sub


	Private Sub cmdSave_Click()
		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNOverwritePrompt As Object
		Dim i, f, J As Short
		Dim temp As String
		Dim Filename_PFS As String

		On Error GoTo Save_Results_PF_Error

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.CancelError = True

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FileName = ""
		SaveFileDialog1.FileName = ""

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
		SaveFileDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel File(*.csv)|*.csv"

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FilterIndex = 2
		SaveFileDialog1.FilterIndex = 3

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.DialogTitle = "Save curves from PSDM"
		SaveFileDialog1.Title = "Save curves from PSDM"

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNOverwritePrompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist

		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Action = 2
		SaveFileDialog1.ShowDialog()

		'f = FileNameIsValid(Filename_PFS, CMDialog1)
		'If Not (f) Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Filename_PFS = CMDialog1.FileName
		Filename_PFS = SaveFileDialog1.FileName


		'Save, T, BVF, Usage rate, C/C0
		f = FreeFile()
		FileOpen(f, Filename_PFS, OpenMode.Output)
		WriteLine(f, "Results file for PSDM - Windows - Version " & VB6.Format(NVersion, "0.00"))
		temp = "Time(min)    BVT(-)   VTM(m^3/kg)   "
		For i = 1 To Results.NComponent
			temp = temp & Trim(Results.Component(i).Name) & "          "
		Next i
		WriteLine(f, temp)
		WriteLine(f)

		temp = ""
		For i = 1 To Results.npoints
			temp = VB6.Format(Results.T(i), "0.00")
			temp = temp & "       " & VB6.Format(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, "0.00")
			temp = temp & "       " & VB6.Format(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.Weight, "0.00")
			For J = 1 To Results.NComponent
				temp = temp & "          " & VB6.Format(Results.CP(J, i), "0.000")
			Next J
			PrintLine(f, temp)
			temp = ""
		Next i
		FileClose(f)
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FileName = ""
		SaveFileDialog1.FileName = ""
		Exit Sub

Save_Results_PF_Error:
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = 75) Then
			'DO NOTHING.
		Else
			Call Show_Trapped_Error("cmdSave_Click")
		End If
		Resume Exit_Save_Results_PF
Exit_Save_Results_PF:
	End Sub


	Private Sub cmdSelect_Click()
		Dim Error_Code As Short
		Dim temp As String
		On Error GoTo Select_Print_Error
		'CMDialog1.flags = PD_PRINTSETUP
		'CMDialog1.Action = 5
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.CancelError = False
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.ShowPrinter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.ShowPrinter()

		Exit Sub
Select_Print_Error:
		Call Show_Trapped_Error("cmdSelect_Click")
		Resume Exit_Select_Print
Exit_Select_Print:
	End Sub


	Private Sub cmdTreat_Click()
		Dim Objective As String
		Dim temp, Tr_Obj As Double
		Dim J, i As Short

		Objective = InputBox("Enter your treatment objective in mg/L for " & Trim(Results.Component(cboCompo.SelectedIndex + 1).Name) & ":", AppName_For_Display_Long, lblData(15).Text)
		On Error GoTo Bad_Treament_Objective
		temp = CDbl(Objective)
		i = cboCompo.SelectedIndex + 1
		Tr_Obj = temp / Results.Component(cboCompo.SelectedIndex + 1).InitialConcentration
		For J = 1 To Number_Points_Max
			If J > 2 Then
				If (Results.CP(i, J) >= Tr_Obj) And (Results.CP(i, J - 1) < Tr_Obj) Then
					Treatment_Objective(cboCompo.SelectedIndex + 1).T = (Results.T(J) - Results.T(J - 1)) / (Results.CP(i, J) - Results.CP(i, J - 1)) * (Tr_Obj - Results.CP(i, J - 1)) + Results.T(J - 1)
					Treatment_Objective(cboCompo.SelectedIndex + 1).C = ((Results.CP(i, J) - Results.CP(i, J - 1)) / (Results.T(J) - Results.T(J - 1)) * (Treatment_Objective(cboCompo.SelectedIndex + 1).T - Results.T(J - 1)) + Results.CP(i, J - 1)) * Results.Component(i).InitialConcentration
					GoTo Exit_Loop
				End If
			End If
		Next J
		Flag_TO(cboCompo.SelectedIndex + 1) = False
		lblData(12).Text = "N/A"
		lblData(13).Text = "N/A"
		lblData(14).Text = "N/A"
		lblData(15).Text = "N/A"
		Exit Sub
Exit_Loop:
		Flag_TO(cboCompo.SelectedIndex + 1) = True
		lblData(12).Text = Format_It(Treatment_Objective(cboCompo.SelectedIndex + 1).T / 60.0# / 24.0#, 2)
		lblData(13).Text = Format_It(Treatment_Objective(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
		lblData(14).Text = Format_It(Treatment_Objective(cboCompo.SelectedIndex + 1).T * 60.0# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
		lblData(15).Text = Format_It(Treatment_Objective(cboCompo.SelectedIndex + 1).C, 2)
		Exit Sub

Bad_Treament_Objective:
		Resume Exit_lblLegend_Click
Exit_lblLegend_Click:


	End Sub

	Private Sub Draw_PFPSDM()
		Dim i, J As Short
		Dim Data_Max, factor As Double
		Dim Bottom_Title As String
		'UPGRADE_WARNING: Lower bound of array X_Values was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim X_Values(Number_Points_Max) As Double
		Dim biggest_numpoints As Short
		Dim index_with_biggest_numpoints As Short
		Dim LastPointI As Short
		Dim SameX As Double
		Dim SameY As Double

		Dim most_recent_x As Double
		Dim most_recent_y As Double
		Dim end_the_plot As Short


		Chart1.Series.Clear()
		Chart1.ChartAreas(0).RecalculateAxesScale()

		'Copy the results
		'UPGRADE_WARNING: Couldn't resolve default property of object optType(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool(_optType_0.Checked) Then 'Time
			factor = 1.0# / 60.0# / 24.0# 'mn > days
			Bottom_Title = "Time(days)"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object optType(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CBool(_optType_1.Checked) Then 'BVF         mn * (mn/s) * (m3/s) / m / (m2) -> dimensionless
				factor = (60.0# * Results.Bed.Flowrate / (Results.Bed.length * PI * (Results.Bed.Diameter / 2.0#) ^ 2)) / 1000
				Bottom_Title = "Bed Volumes Treated (x1000)"
			Else 'Treatment Capacity
				factor = 60.0# * Results.Bed.Flowrate / Results.Bed.Weight 'mn * (s/mn) * (m3/s) / (kg) -> m3/kg
				'factor = 60# * Results.Bed.Flowrate / Results.Bed.Length / Pi / (Results.Bed.Diameter / 2#) ^ 2
				'factor = factor / (Bed.density * 1000)
				Bottom_Title = "m" & Chr(179) & " treated per kg of adsorbent"
			End If
		End If
		'Results.T(I,1) time is in mn
		'Results.T(I,2) is BVF
		For i = 1 To Number_Points_Max
			X_Values(i) = Results.T(i) * factor
		Next i

		' The following code is a rather
		' unfortunate kludge, in my opinion.  I could find no other way to
		' convince/force Visual Basic's graphical interface to accept two sets
		' of data that were of two different sizes, so I determined which one
		' was the smaller set and then filled the remainer of the smaller set
		' with copies of the last data point in it (X,Y) (note, the default
		' is for the data to hook back to the point (0,0) at the end of its
		' plotting due to the fact that, by default, the (X,Y) data points
		' that are unspecified are filled with 0's).
		' -- If possible, it would be nice to replace this with something
		' more elegant, but hey, it works. -- Eric J. Oman, 7/31/96

		'Define Graph
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.NumSets = Results.NComponent
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.GraphType = 6
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.GraphStyle = 4

		''Determine the set with the largest number of data points
		'biggest_numpoints = -1
		'index_with_biggest_numpoints = -1
		'For j = 1 To grpBreak.NumSets
		'  'If (biggest_numpoints < Results.NumPoints_Before_ThroughPut_100(j)) Then
		'  If (biggest_numpoints < Results.NPoints) Then
		'    index_with_biggest_numpoints = j
		'    'biggest_numpoints = Results.NumPoints_Before_ThroughPut_100(j)
		'    biggest_numpoints = Results.NPoints
		'  End If
		'Next j
		biggest_numpoints = Results.npoints



		'add to end of list
		'Dim arr As Integer() = {2, 3, 7}
		'Dim newItem As Integer = 4

		'Dim arr As Double() = {0}


		'Array.Resize(arr, arr.Length + 1)
		'arr(arr.Length - 1) = newItem

		'add points to series
		'For i = 0 To (arr.Length - 1)
		's.Points.AddXY(i, arr(i))
		'Next i

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'For J = 1 To grpBreak.NumSets
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisSet = J
		'grpBreak.NumPoints = Results.NPoints
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.NumPoints = biggest_numpoints
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.PatternData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.PatternData = 1
		'Next J

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.AutoInc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.AutoInc = 0

		Dim dbl_CPConversionFactor As Double
		Dim dblConvertedCP As Double
		Dim OUT_strYAxisTitle As String
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For J = 1 To Results.NComponent

			Dim s As New Series
			s.ChartType = SeriesChartType.Line


			dbl_CPConversionFactor = CBOYAXISTYPE_GetUnitConversion(CShort(VB6.GetItemData(cboYAxisType, cboYAxisType.SelectedIndex)), Results.is_psdm_in_room_model, Results.AnyCrCloseToZero, J, Results.Bed.Phase, OUT_strYAxisTitle)
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ThisSet = J
			end_the_plot = False
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For i = 1 To Results.npoints
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.ThisPoint = i
				If (end_the_plot = False) Then
					dblConvertedCP = Results.CP(J, i) * dbl_CPConversionFactor
					If (Results.CP(J, i) < 0) Then
						If (Results.CP(J, i) = -10000.0#) Then
							end_the_plot = True
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'grpBreak.GraphData = 0#
						End If
					Else 'Results.CP(1,1)
						''''grpBreak.GraphData = Results.CP(j, i)
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpBreak.GraphData = dblConvertedCP

						'add converted to s

						s.Points.AddXY(X_Values(i), dblConvertedCP)

					End If
					If (end_the_plot = False) Then
						'grpBreak.ThisPoint = i
						'grpBreak.LabelText = ""
						'grpBreak.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'	grpBreak.XPosData = X_Values(i)
						most_recent_x = X_Values(i)
						''''most_recent_y = Results.CP(j, i)
						most_recent_y = dblConvertedCP
					End If
				End If
				If (end_the_plot = True) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.GraphData = most_recent_y
					'grpBreak.ThisPoint = i
					'grpBreak.LabelText = ""
					'grpBreak.ThisPoint = i
					'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpBreak.XPosData = most_recent_x
				End If
			Next i
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ThisPoint = J
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.LegendText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.LegendText = Trim(Results.Component(J).Name)

			s.LegendText = Trim(Results.Component(J).Name)
			Chart1.Series.Add(s)

		Next J

		Chart1.Legends(0).Title = "Component:"

		Chart1.ChartAreas(0).AxisX.Minimum = 0

		'rouding for max
		Dim roundmax As Integer
		Dim logscaler As Double 'used for sclaing log
		logscaler = Math.Log10(most_recent_x)
		logscaler = Math.Floor(logscaler)
		logscaler = logscaler - 1
		logscaler = 10 ^ logscaler

		roundmax = Math.Ceiling(most_recent_x / logscaler) * logscaler

		'Chart1.ChartAreas(0).AxisX.Maximum = roundmax
		'Dont need this method anymore!

		Chart1.ChartAreas(0).AxisX.Title = Bottom_Title
		Chart1.ChartAreas(0).AxisX.LabelStyle.Font = New System.Drawing.Font("Times New Roman", 10.25F)

		Chart1.ChartAreas(0).AxisY.Title = OUT_strYAxisTitle
		Chart1.ChartAreas(0).AxisY.LabelStyle.Font = New System.Drawing.Font("Times New Roman", 10.25F)

		''Next, set values for remaining sets with # points < biggest_numpoints
		'For j = 1 To grpBreak.NumSets
		'  If (j <> index_with_biggest_numpoints) Then
		'    grpBreak.ThisSet = j
		'    LastPointI = Results.NumPoints_Before_ThroughPut_100(j)
		'    SameX = X_Values(LastPointI)
		'    SameY = Results.CP(j, LastPointI)
		'    For i = LastPointI + 1 To biggest_numpoints
		'      grpBreak.ThisPoint = i
		'      grpBreak.GraphData = SameY
		'      grpBreak.ThisPoint = i
		'      grpBreak.XPosData = SameX
		'    Next i
		'  End If
		'Next j

		'Other formatting
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.PatternedLines. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.PatternedLines = 0
		Data_Max = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		''For J = 1 To grpBreak.NumSets
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisSet = J
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'For i = 1 To grpBreak.NumPoints
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisPoint = i
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'If grpBreak.GraphData > Data_Max Then
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Data_Max = grpBreak.GraphData
		'End If
		'Next i
		'Next J
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisMax = (Int(Data_Max * 10.0# + 1)) / 10.0#
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisTicks. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisTicks = 4
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisStyle = 2
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisMin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisMin = 0#
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.BottomTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.BottomTitle = Bottom_Title

		''''If (Results.is_psdm_in_room_model) Then
		''''  If (Results.AnyCrCloseToZero = True) Then
		''''    grpBreak.LeftTitle = "Cr, " & Chr$(181) & "g/L!!"
		''''  Else
		''''    grpBreak.LeftTitle = "Cr/Cr,ss!!"
		''''  End If
		''''Else
		''''  grpBreak.LeftTitle = "C/Co!!"
		''''End If
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.LeftTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.LeftTitle = OUT_strYAxisTitle

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.DrawMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.DrawMode = 2

	End Sub

	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click


		Dim Filename_PFPSDM As String

		'Dim f As Short

		Picture1.Image = CaptureActiveWindow()

		SaveFileDialog1.FileName = ""

		SaveFileDialog1.Title = "Print to File"

		'SaveFileDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"

		'SaveFileDialog1.FilterIndex = 2

		If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
			'SaveFileDialog1.OpenFile()
			Picture1.Image.Save(SaveFileDialog1.FileName, Imaging.ImageFormat.Jpeg)
		End If

		'f = FreeFile()
		Filename_PFPSDM = SaveFileDialog1.FileName
		'FileOpen(f, Filename_PFPSDM, OpenMode.Output)
		' Set focus back to form.
		'Me.Activate()



		'PrintPictureToFitPage(f, (Picture1.Image))
		'FileClose((f))

	End Sub

	Private Sub frmModelPSDMResults_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

		rs.FindAllControls(Me)

		Dim J, i As Short
		'Set Window
		'
		' MISC INITS.
		'
		Call Populate_cboYAxisType()
		'is_psdm_in_room_model As Integer
		'int_Which_PSDMR_Model As Integer
		'Global Const PSDMR_MODE_INROOM = 1
		'Global Const PSDMR_MODE_ALONE = 2
		If (Results.is_psdm_in_room_model) Then
			Select Case Results.int_Which_PSDMR_Model
				Case PSDMR_MODE_INROOM
					''''Me.Caption = "Results for the PSDM in Room Model"
					Me.Text = "Results for the PSDMR-in-Room Model (Reactions Present)"
					lblLegend(4).Text = "5% of Cr,ss"
					lblLegend(5).Text = "50% of Cr,ss"
					lblLegend(6).Text = "95% of Cr,ss"
					'UPGRADE_WARNING: Couldn't resolve default property of object ssframe_SSConc.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'ssframe_SSConc.Visible = True
				Case PSDMR_MODE_ALONE
					''''Me.Caption = "Results for the PSDM in Room Model"
					Me.Text = "Results for the PSDMR-Alone Model (Reactions Present)"
					lblLegend(4).Text = "5% of influent conc."
					lblLegend(5).Text = "50% of influent conc."
					lblLegend(6).Text = "95% of influent conc."
					'UPGRADE_WARNING: Couldn't resolve default property of object ssframe_SSConc.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'ssframe_SSConc.Visible = False
			End Select
		Else
			Me.Text = "Results for the PSDM (No Reactions Present)"
			lblLegend(4).Text = "5% of influent conc."
			lblLegend(5).Text = "50% of influent conc."
			lblLegend(6).Text = "95% of influent conc."
			'UPGRADE_WARNING: Couldn't resolve default property of object ssframe_SSConc.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'ssframe_SSConc.Visible = False
		End If
		lblSSValueUnits.Text = "µg/L"

		Call CenterOnForm(Me, frmMain)

		PopulatingScrollboxes = False
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		''''Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmbreak.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmbreak.Height / 2)
		''''Me.HelpContextID = Hlp_Results_for
		''''CMDialog1.CancelError = True
		lblLegend(2).Text = "BVT(m" & Chr(179) & "/m" & Chr(179) & ")"
		lblLegend(3).Text = "VTM(m" & Chr(179) & "/kg)"
		Call Populate_Scrollboxes()

		Call cboCompo_SelectedIndexChanged(cboCompo, New System.EventArgs())
		Call cboGrid_SelectedIndexChanged(cboGrid, New System.EventArgs())
		'cboCompo.SelectedIndex = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object optType().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call optType_click(1, New System.EventArgs())

		Me.Refresh()
		'    optType(1) = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		'grpBreak.GridStyle = 0
		Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 0
		Chart1.ChartAreas(0).AxisY.MajorGrid.LineWidth = 0
	End Sub
	Private Sub frmModelPSDMResults_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call UserPrefs_Save()
	End Sub


	'	Private Sub optType_Click(ByRef Index As Short, ByRef Value As Short)
	Private Sub optType_click(sender As Object, e As EventArgs) Handles _optType_2.Click, _optType_1.Click, _optType_0.Click
		If (Not PopulatingScrollboxes) Then
			Call Draw_PFPSDM()
		End If
	End Sub


	Private Sub Populate_Scrollboxes()
		Dim i As Short

		PopulatingScrollboxes = True

		cboGrid.Items.Add("None")
		cboGrid.Items.Add("Horizontal")
		cboGrid.Items.Add("Vertical")
		cboGrid.Items.Add("Both")

		cboCompo.Items.Clear()   'Shang otherwise duplicates in cboCompo, for example componnet1, component1, component2, component2

		For i = 1 To Results.NComponent
			cboCompo.Items.Add(Trim(Results.Component(i).Name))
			'UPGRADE_WARNING: Couldn't resolve default property of object Treatment_Objective(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Treatment_Objective(i) = Results.ThroughPut_05(i)
			If Treatment_Objective(i).C <> -1 Then
				Flag_TO(i) = True
			Else
				Flag_TO(i) = False
			End If
		Next i

		'-- Read in INI settings
		cboGrid.SelectedIndex = 0
		cboCompo.SelectedIndex = 0
		Call UserPrefs_Load()

		PopulatingScrollboxes = False

	End Sub

	Private Sub UserPrefs_Load()
		Dim X As Integer

		On Error GoTo err_FRMBREAK_UserPrefs_Load

		X = CInt(INI_Getsetting("FRMBREAK_cboGrid"))
		If ((X >= 0) And (X <= cboGrid.Items.Count - 1)) Then
			cboGrid.SelectedIndex = X
		End If
		X = CInt(INI_Getsetting("FRMBREAK_optType"))
		If ((X >= 0) And (X <= 2)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object optType().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			optType(X).Checked = True
		End If

		Exit Sub

resume_err_FRMBREAK_UserPrefs_Load:
		Call UserPrefs_Save()
		Exit Sub

err_FRMBREAK_UserPrefs_Load:
		Resume resume_err_FRMBREAK_UserPrefs_Load

	End Sub

	Private Sub UserPrefs_Save()
		Dim X As Integer

		X = cboGrid.SelectedIndex
		Call INI_PutSetting("FRMBREAK_cboGrid", Trim(CStr(X)))
		'UPGRADE_WARNING: Couldn't resolve default property of object optType(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool((_optType_0.Checked)) Then X = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object optType(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool((_optType_1.Checked)) Then X = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object optType(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool((_optType_2.Checked)) Then X = 2
		Call INI_PutSetting("FRMBREAK_optType", Trim(CStr(X)))

	End Sub

	Private Sub cmdExit_ClickEvent(sender As Object, e As EventArgs)
		Me.Dispose()   'Shang from Close to Dispose so that results are updated correctly with regard to component selection

	End Sub


	Private Sub _optType_1_Click(sender As Object, e As EventArgs)

	End Sub

	Private Sub cmdExcel_ClickEvent(sender As Object, e As EventArgs)
		Call cmdExcel_Click()
	End Sub

	Private Sub cmdSelect_ClickEvent(sender As Object, e As EventArgs)
		Call cmdSelect_Click()
	End Sub

	Private Sub cmdPrint_ClickEvent(sender As Object, e As EventArgs)
		Call cmdPrint_Click()
	End Sub


	Private Sub cmdTreatA_ClickEvent(sender As Object, e As EventArgs)
		Call cmdTreat_Click()
	End Sub

	Private Sub SaveCurves_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
		Call cmdSave_Click()
	End Sub

	Private Sub PrinttoFile_Click(sender As Object, e As EventArgs) Handles cmdFile.Click
		Call cmdFile_Click()
	End Sub

	Private Sub Excel_Click(sender As Object, e As EventArgs) Handles cmdExcel.Click
		Call cmdExcel_Click()
	End Sub

	Private Sub Select_Printer_Click(sender As Object, e As EventArgs)
		Call cmdSelect_Click()
	End Sub

	Private Sub Print_Click(sender As Object, e As EventArgs)
		Call cmdPrint_Click()
	End Sub

	Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
		Me.Dispose() 'dispose instead of close
	End Sub

	Private Sub cmdTreatA_Click(sender As Object, e As EventArgs) Handles cmdTreatA.Click
		Call cmdTreat_Click()
	End Sub

	Private Sub cmdTreat_Click(sender As Object, e As EventArgs) Handles cmdTreat.Click
		Call cmdTreat_Click()
	End Sub

	Private Sub frmModelPSDMResults_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class