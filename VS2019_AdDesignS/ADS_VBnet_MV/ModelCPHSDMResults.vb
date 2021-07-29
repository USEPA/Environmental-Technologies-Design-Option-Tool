Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Math

Friend Class frmModelCPHSDMResults
	Inherits System.Windows.Forms.Form

	Dim rs As New Resizer

	Dim Treatment_Objective As Throughput
	Dim Flag_TO As Short
	
	Dim PopulatingScrollboxes As Short
	
	
	
	Const frmModelCPHSDMResults_declarations_end As Boolean = True
	
	
	Private Sub Draw_CPM()
		Dim i, J, f As Short
		Dim FileNamebis As String
		Dim Data_Max, factor As Double
		Dim Bottom_Title As String
		'UPGRADE_WARNING: Lower bound of array X_Values was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim X_Values(CPM_Max_Points) As Double


		Chart1.Series.Clear()
		Chart1.ChartAreas(0).RecalculateAxesScale()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

		'UPGRADE_WARNING: Couldn't resolve default property of object optType(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool(_optType_0.Checked) Then 'Time   'changed from opttype(0) Shang
			factor = 1.0#
			Bottom_Title = "Time(days)"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object optType(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CBool(_optType_1.Checked) Then 'BVF
				factor = (24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2.0#) ^ 2) / 1000
				Bottom_Title = "Bed Volumes Treated (Thousands)"
			Else 'Treatment Capacity
				factor = 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight
				Bottom_Title = "m" & Chr(179) & " treated per kg of adsorbent"
			End If
		End If
		For i = 1 To CPM_Max_Points
			X_Values(i) = CPM_Results.T(i) * factor
		Next i

		'Define Graph
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.NumSets = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.GraphType = 6 'SCATTER
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.GraphStyle = 4
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.NumPoints = 100


		'TEMP BEGINS.
		'grpBreak.ThisSet = 1
		'grpBreak.NumPoints = 100
		'grpBreak.AutoInc = 0
		'For i = 1 To 100
		'  grpBreak.ThisPoint = i
		'  grpBreak.GraphData = CDbl(i)
		'  grpBreak.XPosData = CDbl(i)
		'Next i
		'TEMP ENDS.

		Dim s As New Series
		s.ChartType = SeriesChartType.Line

		Chart1.ChartAreas(0).AxisX.Minimum = 0

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.AutoInc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.AutoInc = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GridStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.GridStyle = cboGrid.SelectedIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisSet = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 1 To 100 'numpoints = 100
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.ThisPoint = i
			If CPM_Results.C_Over_C0(i) < 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.GraphData = 0#
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpBreak.GraphData = CPM_Results.C_Over_C0(i)
				s.Points.AddXY(X_Values(i), CPM_Results.C_Over_C0(i))
			End If
			''''grpBreak.ThisPoint = i
			''''grpBreak.LabelText = ""
			''''grpBreak.ThisPoint = i
			'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.XPosData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpBreak.XPosData = X_Values(i)
		Next i

		Dim X_max As Double
		X_max = X_Values(100)


		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.ThisPoint = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.PatternData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.PatternData = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.PatternedLines. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.PatternedLines = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisStyle = 2
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisMin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisMin = 0#
		Data_Max = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 1 To 100 'numpoints
			If CPM_Results.C_Over_C0(i) > Data_Max Then
				Data_Max = CPM_Results.C_Over_C0(i)
			End If
		Next i

		'rouding for max
		Dim roundmax As Integer
		Dim logscaler As Double 'used for sclaing log
		logscaler = Math.Log10(X_max)
		logscaler = Math.Floor(logscaler)
		logscaler = logscaler - 1
		logscaler = 10 ^ logscaler

		roundmax = Math.Ceiling(X_max / logscaler) * logscaler

		'Chart1.ChartAreas(0).AxisX.Maximum = roundmax

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisMax = (Int(Data_Max * 10# + 1)) / 10#
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.YAxisTicks. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.YAxisTicks = 4
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.BottomTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.BottomTitle = Bottom_Title
		''''grpBreak.BottomTitle = "Testing"
		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.LeftTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.LeftTitle = "C/Co"

		Chart1.Legends(0).Title = "Component:"
		s.LegendText = Trim(CPM_Results.Component.Name)

		Chart1.ChartAreas(0).AxisX.Title = Bottom_Title
		Chart1.ChartAreas(0).AxisX.LabelStyle.Font = New System.Drawing.Font("Times New Roman", 10.25F)

		Chart1.ChartAreas(0).AxisY.Title = "C/Co"
		Chart1.ChartAreas(0).AxisY.LabelStyle.Font = New System.Drawing.Font("Times New Roman", 10.25F)

		Chart1.Series.Add(s)

		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

		'UPGRADE_WARNING: Couldn't resolve default property of object grpBreak.DrawMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpBreak.DrawMode = 2

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
	
	
	Private Sub cmdExcel_Click()
		PFPSDM_Excel = False
		CPHSDM_Excel = True
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
		Dim Eq1 As String
		Dim Filename_Input As String
		
		On Error GoTo File_Error
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
		'CMDialog1.flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Action = 2

		'f = FileNameIsValid(Filename_Input, CMDialog1)
		'If Not (f) Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Filename_Input = CMDialog1.FileName
		Filename_Input = SaveFileDialog1.FileName

		f = FreeFile
		FileOpen(f, Filename_Input, OpenMode.Output)
		
		PrintLine(f, "Input data for the Constant Pattern Model")
		'-- Print Filename
		
		PrintLine(f)
		PrintLine(f, "From Data File :", Filename)
		
		
		PrintLine(f)
		PrintLine(f, "Chemical:", TAB(10), Trim(CPM_Results.Component.Name))
		PrintLine(f, TAB(5), "Molecular weight: ", TAB(28), VB6.Format(CPM_Results.Component.MW, "0.00") & " g/mol")
		PrintLine(f, TAB(5), "Normal Boiling Point: ", TAB(28), VB6.Format(CPM_Results.Component.BP, "0.00") & " C")
		PrintLine(f, TAB(5), "Molar Volume @ NBP: ", TAB(28), Format_It(CPM_Results.Component.MolarVolume, 2) & " cm" & Chr(179) & "/mol")
		PrintLine(f, TAB(5), "Initial Concentration: ", TAB(28), Format_It(CPM_Results.Component.InitialConcentration, 2) & " mg/L")
		PrintLine(f, TAB(5), "K: ", TAB(28), VB6.Format(CPM_Results.Component.Use_K, "0.000") & " (mg/g)(L/mg)^(1/n)")
		PrintLine(f, TAB(5), "1/n: ", TAB(28), VB6.Format(CPM_Results.Component.Use_OneOverN, "0.000"))
		PrintLine(f)
		
		'-----------------------Bed Data ----------------------
		PrintLine(f, "Bed Data:")
		
		PrintLine(f, TAB(5), "Bed Length: ", TAB(28), VB6.Format(CPM_Results.Bed.length, "0.000E+00") & " m")
		PrintLine(f, TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(CPM_Results.Bed.Diameter, "0.000E+00") & " m")
		PrintLine(f, TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(CPM_Results.Bed.Weight, "0.000E+00") & " kg")
		PrintLine(f, TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(CPM_Results.Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		PrintLine(f, TAB(5), "EBCT: ", TAB(28), VB6.Format(CPM_Results.Bed.length * PI * CPM_Results.Bed.Diameter * CPM_Results.Bed.Diameter / 4# / CPM_Results.Bed.Flowrate / 60#, "0.000E+00") & " mn")
		PrintLine(f)
		PrintLine(f, TAB(5), "Temperature:", TAB(28), VB6.Format(CPM_Results.Bed.Temperature, "0.00") & " C")
		If CPM_Results.Bed.Phase = 0 Then
			PrintLine(f, TAB(5), "Water Density:", TAB(28), VB6.Format(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			PrintLine(f, TAB(5), "Water Viscosity:", TAB(28), VB6.Format(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			PrintLine(f, TAB(5), "Pressure:", TAB(28), VB6.Format(CPM_Results.Bed.Pressure, "0.00000") & " atm")
			PrintLine(f, TAB(5), "Air Density:", TAB(28), VB6.Format(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			PrintLine(f, TAB(5), "Air Viscosity:", TAB(28), VB6.Format(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		PrintLine(f)
		
		'-----------------Carbon Properties -------------------------------
		PrintLine(f, "Carbon Properties:")
		
		PrintLine(f, TAB(5), "Name: ", TAB(28), Trim(CPM_Results.Carbon.Name))
		PrintLine(f, TAB(5), "Apparent Density: ", TAB(28), VB6.Format(CPM_Results.Carbon.Density, "0.000") & " g/cm" & Chr(179))
		PrintLine(f, TAB(5), "Particle Radius: ", TAB(28), VB6.Format(CPM_Results.Carbon.ParticleRadius * 100#, "0.000000") & " cm")
		PrintLine(f, TAB(5), "Porosity: ", TAB(28), VB6.Format(CPM_Results.Carbon.Porosity, "0.000"))
		PrintLine(f, TAB(5), "Shape Factor: ", TAB(28), VB6.Format(CPM_Results.Carbon.ShapeFactor, "0.000"))
		'Print #f, Tab(5); "Tortuosity: "; Tab(28); Format$(CPM_Results.Carbon.Tortuosity, "0.000")
		PrintLine(f)
		
		'---------------Kinetic Parameters -----------------------------------------
		PrintLine(f, "Kinetic parameters:")
		PrintLine(f, TAB(5), "kf", TAB(28), Format_It(CPM_Results.Component.kf, 2) & " cm/s")
		PrintLine(f, TAB(5), "Ds", TAB(28), Format_It(CPM_Results.Component.Ds, 2) & " cm" & Chr(178) & "/s")
		PrintLine(f, TAB(5), "SPDFR", TAB(28), Format_It(CPM_Results.Component.SPDFR, 2))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Component(0) = CPM_Results.Component
		PrintLine(f, TAB(5), "St", TAB(28), Format_It(ST(0), 2))
		'UPGRADE_WARNING: Couldn't resolve default property of object Eds(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(f, TAB(5), "Eds", TAB(28), Format_It(Eds(0), 2))
		
		PrintLine(f)
		
		'Fouling-----------------------------------------
		PrintLine(f, "Fouling correlations:")
		PrintLine(f)
		PrintLine(f)
		PrintLine(f, " Water type : " & Trim(CPM_Results.Bed.Water_Correlation.Name))
		Eq1 = VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(1), "0.00")
		
		If CPM_Results.Bed.Water_Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
		Else
			If CPM_Results.Bed.Water_Correlation.Coeff(2) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(CPM_Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
			End If
		End If
		If CPM_Results.Bed.Water_Correlation.Coeff(3) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
		Else
			If CPM_Results.Bed.Water_Correlation.Coeff(3) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(CPM_Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
			End If
		End If
		If CPM_Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
			If CPM_Results.Bed.Water_Correlation.Coeff(4) > 0 Then
				Eq1 = Eq1 & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
			Else
				If CPM_Results.Bed.Water_Correlation.Coeff(4) < 0 Then
					Eq1 = Eq1 & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
				End If
			End If
		End If
		PrintLine(f, "K(t)/K0 = " & Eq1)
		PrintLine(f, "(t in minutes)")
		PrintLine(f)
		
		Eq1 = ""
		If CPM_Results.Component.Correlation.Coeff(1) = 1# Then
			Eq1 = "(K/K0) "
		Else
			If CPM_Results.Component.Correlation.Coeff(1) <> 0 Then Eq1 = VB6.Format(CPM_Results.Component.Correlation.Coeff(1), "0.00") & " * (K/K0) "
		End If
		If CPM_Results.Component.Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & "+ " & VB6.Format(CPM_Results.Component.Correlation.Coeff(2), "0.00")
		Else
			If CPM_Results.Component.Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & VB6.Format(System.Math.Abs(CPM_Results.Component.Correlation.Coeff(2)), "0.00")
		End If
		If Trim(Eq1) = "" Then
			Eq1 = "K/K0"
		End If
		PrintLine(f, Trim(CPM_Results.Component.Name) & ":")
		PrintLine(f, TAB(10), "Correlation type: " & Trim(CPM_Results.Component.Correlation.Name))
		PrintLine(f, TAB(10), "K/K0 = " & Eq1)
		PrintLine(f)
		
		If (CPM_Results.Component.Use_Tortuosity_Correlation) Then
			If (CPM_Results.Component.Constant_Tortuosity) Then
				PrintLine(f, "Correlation used when SOC competition is important:")
				PrintLine(f, " Tortuosity = 0.782 * EBCT^0.925 ")
			Else
				PrintLine(f, "Correlation used when NOM fouling is important:")
				PrintLine(f, " Tortuosity = 1.0 if t< 70 days")
				PrintLine(f, " Tortuosity = 0.334 + 6.610E-06 * EBCT")
			End If
		End If
		PrintLine(f)
		
		'--------- CPM Results ----------------------------------
		PrintLine(f, "Constant Pattern Model Results for " & Trim(CPM_Results.Component.Name) & ":")
		PrintLine(f)
		PrintLine(f, "Minimum Stanton number:", TAB(30), Format_It(CPM_Results.Par(1), 2))
		PrintLine(f, "Minimum EBCT:", TAB(30), Format_It(CPM_Results.Par(2), 2) & " min")
		PrintLine(f, "Minimum Column Length:", TAB(30), Format_It(CPM_Results.Par(3), 2) & " cm")
		PrintLine(f, "Throughput at 95% of the MTZ:", TAB(30), Format_It(CPM_Results.Par(4), 2))
		PrintLine(f, "Throughput at 5% of the MTZ:", TAB(30), Format_It(CPM_Results.Par(5), 2))
		PrintLine(f, "EBCT of the MTZ:", TAB(30), Format_It(CPM_Results.Par(6), 2) & " min")
		PrintLine(f, "Length of the MTZ:", TAB(30), Format_It(CPM_Results.Par(7), 2) & " cm")
		
		PrintLine(f)
		PrintLine(f, TAB(30), "Time(days)", TAB(40), "BVT", TAB(50), "TC", TAB(60), "C (mg/L)")
		PrintLine(f, "5% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_05.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_05.C, 2))
		PrintLine(f, "50% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_50.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_50.C, 2))
		PrintLine(f, "95% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_95.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_95.C, 2))
		PrintLine(f)
		PrintLine(f, "TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")
		PrintLine(f)
		
		If Flag_TO Then
			PrintLine(f, "Treatment Objective: " & Format_It(Treatment_Objective.C, 2) & " mg/L")
			PrintLine(f)
			PrintLine(f, "Time (days):", TAB(20), Format_It(Treatment_Objective.T, 2))
			PrintLine(f, "BVT:", TAB(20), Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2))
			PrintLine(f, "Tr. Capacity:", TAB(20), Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2))
		Else
			PrintLine(f, "The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective.C, 2) & "mg/L) could not be calculated.")
		End If
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
		Dim H, W As Single
		Dim Eq1 As String
		Dim i As Short
		
		On Error GoTo Print_Error

		'---Print Graph ---------------------------------------------------
		'	'''    H = grpBreak.Height
		'	'''    W = grpBreak.Width
		'	'''
		'	'''    grpBreak.Visible = False 'Hide it before printing

		'
		' THIS CODE HAD TO BE REPLACED TODAY, 1999-MAY-11, EJOMAN.
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

		'MsgBox _
		'"Printer.Height = " & Trim$(Str$(Printer.Height)) & ", " & _
		'"Printer.Width = " & Trim$(Str$(Printer.Width)) & ", " & _
		'"Printer.ScaleHeight = " & Trim$(Str$(Printer.ScaleHeight)) & ", " & _
		'"Printer.ScaleWidth = " & Trim$(Str$(Printer.ScaleWidth)) & ", " & _
		'"Printer.ScaleLeft = " & Trim$(Str$(Printer.ScaleLeft)) & ", " & _
		'"Printer.ScaleTop = " & Trim$(Str$(Printer.ScaleTop)) & ", "

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
		'	'''    Printer.Line _
		''''        (0, 0)- _
		''''        (grpBreak.Width, grpBreak.Height), _
		''''        QBColor(0), _
		''''        B
		'---- NEW CODE ENDS.

		'	'''    grpBreak.Height = H
		'	'''    grpBreak.Width = W
		'	'''
		'	'''    grpBreak.Visible = True
		'	'''
		'	'''    grpBreak.PrintStyle = 2
		'	'''    grpBreak.DrawMode = 2

		'
		' A "SKIP TO NEXT PAGE" COMMAND HAD TO BE ADDED TO THE PRINTING
		' CODE TODAY, 1999-MAY-11, EJOMAN.
		'
		'---- NEW CODE STARTS HERE:
		'	'''    Printer.NewPage
		'---- NEW CODE ENDS.

		'---Print other results------------------------------------------
		Printer.ScaleLeft = -1080 'Set a 3/4-inch margin
		Printer.ScaleTop = -1080
		Printer.CurrentX = 0
		Printer.CurrentY = 0
		
		'-- Print Filename
		
		Printer.FontSize = 10
		Printer.Print("From Data File: " & Filename)
		Printer.Print()
		
		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Input data for the Constant Pattern Model")
		Printer.FontSize = 10
		Printer.FontBold = False
		Printer.FontUnderline = True
		Printer.Print()
		Printer.Print("Chemical:", TAB(10), Trim(CPM_Results.Component.Name))
		Printer.FontUnderline = False
		Printer.Print(TAB(5), "Molecular weight: ", TAB(28), VB6.Format(CPM_Results.Component.MW, "0.00") & " g/mol")
		Printer.Print(TAB(5), "Normal Boiling Point: ", TAB(28), VB6.Format(CPM_Results.Component.BP, "0.00") & " C")
		Printer.Print(TAB(5), "Molar Volume @ NBP: ", TAB(28), Format_It(CPM_Results.Component.MolarVolume, 2) & " cm" & Chr(179) & "/mol")
		Printer.Print(TAB(5), "Initial Concentration: ", TAB(28), Format_It(CPM_Results.Component.InitialConcentration, 2) & " mg/L")
		Printer.Print(TAB(5), "K: ", TAB(28), VB6.Format(CPM_Results.Component.Use_K, "0.000") & " (mg/g)(L/mg)^(1/n)")
		Printer.Print(TAB(5), "1/n: ", TAB(28), VB6.Format(CPM_Results.Component.Use_OneOverN, "0.000"))
		Printer.Print()
		
		'-----------------------Bed Data ----------------------
		Printer.FontUnderline = True
		Printer.Print("Bed Data:")
		Printer.FontUnderline = False
		
		Printer.Print(TAB(5), "Bed Length: ", TAB(28), VB6.Format(CPM_Results.Bed.length, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(CPM_Results.Bed.Diameter, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(CPM_Results.Bed.Weight, "0.000E+00") & " kg")
		Printer.Print(TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(CPM_Results.Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		Printer.Print(TAB(5), "EBCT: ", TAB(28), VB6.Format(CPM_Results.Bed.length * PI * CPM_Results.Bed.Diameter * CPM_Results.Bed.Diameter / 4# / CPM_Results.Bed.Flowrate / 60#, "0.000E+00") & " mn")
		Printer.Print()
		Printer.Print(TAB(5), "Temperature:", TAB(28), VB6.Format(CPM_Results.Bed.Temperature, "0.00") & " C")
		If CPM_Results.Bed.Phase = 0 Then
			Printer.Print(TAB(5), "Water Density:", TAB(28), VB6.Format(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Water Viscosity:", TAB(28), VB6.Format(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			Printer.Print(TAB(5), "Pressure:", TAB(28), VB6.Format(CPM_Results.Bed.Pressure, "0.00000") & " atm")
			Printer.Print(TAB(5), "Air Density:", TAB(28), VB6.Format(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Air Viscosity:", TAB(28), VB6.Format(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		Printer.Print()
		
		'-----------------Carbon Properties -------------------------------
		Printer.FontUnderline = True
		Printer.Print("Carbon Properties:")
		Printer.FontUnderline = False
		
		Printer.Print(TAB(5), "Name: ", TAB(28), Trim(CPM_Results.Carbon.Name))
		Printer.Print(TAB(5), "Apparent Density: ", TAB(28), VB6.Format(CPM_Results.Carbon.Density, "0.000") & " g/cm" & Chr(179))
		Printer.Print(TAB(5), "Particle Radius: ", TAB(28), VB6.Format(CPM_Results.Carbon.ParticleRadius * 100#, "0.000000") & " cm")
		Printer.Print(TAB(5), "Porosity: ", TAB(28), VB6.Format(CPM_Results.Carbon.Porosity, "0.000"))
		Printer.Print(TAB(5), "Shape Factor: ", TAB(28), VB6.Format(CPM_Results.Carbon.ShapeFactor, "0.000"))
		'Printer.Print Tab(5); "Tortuosity: "; Tab(28); Format$(CPM_Results.Carbon.Tortuosity, "0.000")
		Printer.Print()
		
		'---------------Kinetic Parameters -----------------------------------------
		Printer.FontUnderline = True
		Printer.Print("Kinetic parameters:")
		Printer.FontUnderline = False
		Printer.Print(TAB(5), "kf", TAB(28), Format_It(CPM_Results.Component.kf, 2) & " cm/s")
		Printer.Print(TAB(5), "Ds", TAB(28), Format_It(CPM_Results.Component.Ds, 2) & " cm" & Chr(178) & "/s")
		Printer.Print(TAB(5), "SPDFR", TAB(28), Format_It(CPM_Results.Component.SPDFR, 2))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Component(0) = CPM_Results.Component
		Printer.Print(TAB(5), "St", TAB(28), Format_It(ST(0), 2))
		'UPGRADE_WARNING: Couldn't resolve default property of object Eds(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Printer.Print(TAB(5), "Eds", TAB(28), Format_It(Eds(0), 2))
		
		Printer.Print()
		
		'Fouling-----------------------------------------
		Printer.FontUnderline = True
		Printer.Print("Fouling correlations:")
		Printer.FontUnderline = False
		Printer.Print()
		
		Printer.Print()
		Printer.Print(" Water type : " & Trim(CPM_Results.Bed.Water_Correlation.Name))
		Eq1 = VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(1), "0.00")
		
		If CPM_Results.Bed.Water_Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
		Else
			If CPM_Results.Bed.Water_Correlation.Coeff(2) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(CPM_Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
			End If
		End If
		If CPM_Results.Bed.Water_Correlation.Coeff(3) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
		Else
			If CPM_Results.Bed.Water_Correlation.Coeff(3) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(CPM_Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
			End If
		End If
		If CPM_Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
			If CPM_Results.Bed.Water_Correlation.Coeff(4) > 0 Then
				Eq1 = Eq1 & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
			Else
				If CPM_Results.Bed.Water_Correlation.Coeff(4) < 0 Then
					Eq1 = Eq1 & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
				End If
			End If
		End If
		Printer.Print("K(t)/K0 = " & Eq1)
		Printer.Print("(t in minutes)")
		Printer.Print()
		
		
		Eq1 = ""
		If CPM_Results.Component.Correlation.Coeff(1) = 1# Then
			Eq1 = "(K/K0) "
		Else
			If CPM_Results.Component.Correlation.Coeff(1) <> 0 Then Eq1 = VB6.Format(CPM_Results.Component.Correlation.Coeff(1), "0.00") & " * (K/K0) "
		End If
		If CPM_Results.Component.Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & "+ " & VB6.Format(CPM_Results.Component.Correlation.Coeff(2), "0.00")
		Else
			If CPM_Results.Component.Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & VB6.Format(System.Math.Abs(CPM_Results.Component.Correlation.Coeff(2)), "0.00")
		End If
		If Trim(Eq1) = "" Then
			Eq1 = "K/K0"
		End If
		Printer.Print(Trim(CPM_Results.Component.Name) & ":")
		Printer.Print(TAB(10), "Correlation type: " & Trim(CPM_Results.Component.Correlation.Name))
		
		Printer.Print(TAB(10), "K/K0 = " & Eq1)
		
		Printer.Print()
		
		If (CPM_Results.Component.Use_Tortuosity_Correlation) Then
			If (CPM_Results.Component.Constant_Tortuosity) Then
				Printer.Print("Correlation used when SOC competition is important:")
				Printer.Print(" Tortuosity = 0.782 * EBCT^0.925 ")
			Else
				Printer.Print("Correlation used when NOM fouling is important:")
				Printer.Print(" Tortuosity = 1.0 if t< 70 days")
				Printer.Print(" Tortuosity = 0.334 + 6.610E-06 * EBCT")
			End If
		End If
		Printer.Print()
		
		'--------- CPM Results ----------------------------------
		Printer.Print()
		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Constant Pattern Model Results for " & Trim(CPM_Results.Component.Name) & ":")
		Printer.Print()
		Printer.FontSize = 10
		Printer.FontBold = False
		Printer.FontUnderline = False
		Printer.Print("Minimum Stanton number:", TAB(30), Format_It(CPM_Results.Par(1), 2))
		Printer.Print("Minimum EBCT:", TAB(30), Format_It(CPM_Results.Par(2), 2) & " min")
		Printer.Print("Minimum Column Length:", TAB(30), Format_It(CPM_Results.Par(3), 2) & " cm")
		Printer.Print("MTZ Length:", TAB(30), Format_It(CPM_Results.Par(7), 2) & " cm")
		
		Printer.Print()
		Printer.Print(TAB(30), "Time(days)", TAB(40), "BVT", TAB(50), "TC", TAB(60), "C (mg/L)")
		Printer.Print("5% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_05.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_05.C, 2))
		Printer.Print("50% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_50.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_50.C, 2))
		Printer.Print("95% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_95.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_95.C, 2))
		Printer.Print()
		Printer.Print("TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")
		Printer.Print()
		
		If Flag_TO Then
			Printer.Print("Treatment Objective: " & Format_It(Treatment_Objective.C, 2) & " mg/L")
			Printer.Print()
			Printer.Print("Time (days):", TAB(20), Format_It(Treatment_Objective.T, 2))
			Printer.Print("BVT:", TAB(20), Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2))
			Printer.Print("Tr. Capacity:", TAB(20), Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2))
		Else
			Printer.Print("The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective.C, 2) & "mg/L) could not be calculated.")
		End If
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
		Dim f, i As Short
		Dim temp As String
		Dim Filename_CPM As String
		On Error GoTo Save_Results_CPM_Error
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.CancelError = True
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FileName = ""
		SaveFileDialog1.FileName = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
		SaveFileDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel File (*.csv)|*.csv"
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FilterIndex = 2
		SaveFileDialog1.FilterIndex = 3
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.DialogTitle = "Save curve from Constant Pattern Model"
		SaveFileDialog1.Title = "Save curve from Constant Pattern Model"
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNOverwritePrompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Action = 2
		SaveFileDialog1.ShowDialog()
		'f = FileNameIsValid(Filename_CPM, CMDialog1)
		'If Not (f) Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Filename_CPM = CMDialog1.FileName
		Filename_CPM = SaveFileDialog1.FileName


		f = FreeFile
		FileOpen(f, Filename_CPM, OpenMode.Output)
		WriteLine(f, "Results file for Constant Pattern Model")
		temp = "Time(days)       "
		temp = temp & "BVT" & "        " & "Usage Rate " & "     " & Trim(CPM_Results.Component.Name)
		PrintLine(f, temp)
		PrintLine(f, " days             -         m" & Chr(179) & "/kg GAC  ")
		WriteLine(f)
		temp = ""
		For i = 1 To 100
			temp = VB6.Format(CPM_Results.T(i), "0.00")
			temp = temp & "       " & VB6.Format(CPM_Results.T(i) * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2#) ^ 2, "0.00")
			temp = temp & "       " & VB6.Format(CPM_Results.T(i) * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, "0.00")
			temp = temp & "       " & VB6.Format(CPM_Results.C_Over_C0(i), "0.000")
			PrintLine(f, temp)
			temp = ""
		Next i
		FileClose(f)
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.FileName = ""
		SaveFileDialog1.FileName = ""
		Exit Sub
Save_Results_CPM_Error:
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = 75) Then
			'DO NOTHING.
		Else
			Call Show_Trapped_Error("cmdSave_Click")
		End If
		Resume Exit_Save_Results_CPM
Exit_Save_Results_CPM: 
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
		Dim J As Short
		Objective = InputBox("Enter your treatment objective in mg/L:", AppName_For_Display_Long, lblData(9).Text)
		On Error GoTo Bad_Treament_Objective
		temp = CDbl(Objective)
		Tr_Obj = temp / CPM_Results.Component.InitialConcentration
		For J = 1 To CPM_Max_Points
			If J > 2 Then
				If (CPM_Results.C_Over_C0(J) >= Tr_Obj) And (CPM_Results.C_Over_C0(J - 1) < Tr_Obj) Then
					Treatment_Objective.T = (CPM_Results.T(J) - CPM_Results.T(J - 1)) / (CPM_Results.C_Over_C0(J) - CPM_Results.C_Over_C0(J - 1)) * (Tr_Obj - CPM_Results.C_Over_C0(J - 1)) + CPM_Results.T(J - 1)
					Treatment_Objective.C = ((CPM_Results.C_Over_C0(J) - CPM_Results.C_Over_C0(J - 1)) / (CPM_Results.T(J) - CPM_Results.T(J - 1)) * (Treatment_Objective.T - CPM_Results.T(J - 1)) + CPM_Results.C_Over_C0(J - 1)) * CPM_Results.Component.InitialConcentration
					GoTo Exit_Loop
				End If
			End If
		Next J
		Flag_TO = False
		lblData(15).Text = "N/A"
		lblData(16).Text = "N/A"
		lblData(17).Text = "N/A"
		lblData(18).Text = "N/A"
		Exit Sub
Exit_Loop: 
		Flag_TO = True
		lblData(15).Text = Format_It(Treatment_Objective.T, 2)
		lblData(16).Text = Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
		lblData(17).Text = Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
		lblData(18).Text = Format_It(Treatment_Objective.C, 2)
		Exit Sub
Bad_Treament_Objective: 
		Resume Exit_lblLegend_Click
Exit_lblLegend_Click: 
	End Sub
	
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) 
		Dim Printer As New Printer
		Picture1.Image = CaptureActiveWindow()
		PrintPictureToFitPage(Printer, (Picture1.Image))
		Printer.EndDoc()
		' Set focus back to form.
		Me.Activate()
	End Sub
	
	Private Sub frmModelCPHSDMResults_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Short
		rs.FindAllControls(Me)

		PopulatingScrollboxes = False
		_optType_0.Checked = True  'default load
		_optType_1.Checked = False
		_optType_2.Checked = False

		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		'Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmCPM.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmCPM.Height / 2)
		Call CenterOnForm(Me, frmMain)

		'UPGRADE_WARNING: Couldn't resolve default property of object Frame3D1.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GroupBox2.Text = "Results for " & Trim(CPM_Results.Component.Name) & ":"
		lblParaValue(0).Text = Format_It(CPM_Results.Par(1), 2) 'Minimum Stanton
		lblParaValue(2).Text = Format_It(CPM_Results.Par(7), 2) 'MTZ Length
		lblParaValue(5).Text = Format_It(CPM_Results.Par(2), 2) 'Minimum EBCT
		lblParaValue(6).Text = Format_It(CPM_Results.Par(3), 2) 'Minimum Column Length
		
		lblLegend(2).Text = "BVT(m" & Chr(179) & "/m" & Chr(179) & ")"
		lblLegend(3).Text = "VTM(m" & Chr(179) & "/kg)"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Treatment_Objective. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Treatment_Objective = CPM_Results.ThroughPut_05
		
		Call Populate_Scrollboxes()
		Call cboGrid_SelectedIndexChanged(cboGrid, New System.EventArgs())
		'UPGRADE_WARNING: Couldn't resolve default property of object optType().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call optType_Click(_optType_0, New System.EventArgs())

		'cboGrid.AddItem "None"
		'cboGrid.AddItem "Horizontal"
		'cboGrid.AddItem "Vertical"
		'cboGrid.AddItem "Both"
		'cboGrid.ListIndex = 0

		'    optType(0) = True

		Flag_TO = True
		lblData(0).Text = Format_It(CPM_Results.ThroughPut_05.T, 2)
		lblData(1).Text = Format_It(CPM_Results.ThroughPut_50.T, 2)
		lblData(2).Text = Format_It(CPM_Results.ThroughPut_95.T, 2)
		'------ C -------
		lblData(9).Text = Format_It(CPM_Results.ThroughPut_05.C, 2)
		lblData(10).Text = Format_It(CPM_Results.ThroughPut_50.C, 2)
		lblData(11).Text = Format_It(CPM_Results.ThroughPut_95.C, 2)
		
		'----- BVF ---------
		lblData(3).Text = Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
		lblData(4).Text = Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
		lblData(5).Text = Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
		
		'-----Carbon Us. rate --------- m3 of water/kg of GAC
		lblData(6).Text = Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
		lblData(7).Text = Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
		lblData(8).Text = Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)

		'-----Treatment Objective------
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdTreat.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdTreat.Text = "Treat. Objective"
		lblData(15).Text = lblData(0).Text
		lblData(16).Text = lblData(3).Text
		lblData(17).Text = lblData(6).Text
		lblData(18).Text = lblData(9).Text
		
		'     grpBreak.GridStyle = 0
		
	End Sub
	Private Sub frmModelCPHSDMResults_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call UserPrefs_Save()
	End Sub


	'Private Sub optType_Click(ByRef Index As Short, ByRef Value As Short)
	Private Sub optType_Click(sender As Object, e As EventArgs) Handles _optType_2.Click, _optType_1.Click, _optType_0.Click
		If (Not PopulatingScrollboxes) Then
			Call Draw_CPM()
		End If
	End Sub


	Private Sub Populate_Scrollboxes()
		Dim i As Short
		PopulatingScrollboxes = True
		cboGrid.Items.Add("None")
		cboGrid.Items.Add("Horizontal")
		cboGrid.Items.Add("Vertical")
		cboGrid.Items.Add("Both")
		'-- Read in INI settings
		cboGrid.SelectedIndex = 0
		Call UserPrefs_Load()
		PopulatingScrollboxes = False
	End Sub
	
	
	Private Sub UserPrefs_Load()
		Dim X As Integer
		On Error GoTo err_FRMCPM_UserPrefs_Load
		X = CInt(INI_Getsetting("FRMCPM_cboGrid"))
		If ((X >= 0) And (X <= cboGrid.Items.Count - 1)) Then
			cboGrid.SelectedIndex = X
		End If
		X = CInt(INI_Getsetting("FRMCPM_optType"))
		If ((X >= 0) And (X <= 2)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object optType().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If X = 0 Then
				_optType_0.Checked = True
			ElseIf X = 1 Then
				_optType_1.Checked = True
			ElseIf X = 2 Then
				_optType_2.Checked = True
			End If

			' optType(X).Value = True  'original out by Shang
		End If
		Exit Sub
resume_err_FRMCPM_UserPrefs_Load: 
		Call UserPrefs_Save()
		Exit Sub
err_FRMCPM_UserPrefs_Load: 
		Resume resume_err_FRMCPM_UserPrefs_Load
	End Sub
	Private Sub UserPrefs_Save()
		Dim X As Integer
		X = cboGrid.SelectedIndex
		Call INI_PutSetting("FRMCPM_cboGrid", Trim(CStr(X)))
		'UPGRADE_WARNING: Couldn't resolve default property of object optType(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool((_optType_0.Checked)) Then X = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object optType(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool((_optType_1.Checked)) Then X = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object optType(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CBool((_optType_2.Checked)) Then X = 2
		Call INI_PutSetting("FRMCPM_optType", Trim(CStr(X)))
	End Sub

	Private Sub cmdExit_ClickEvent(sender As Object, e As EventArgs)
		Me.Dispose()   'Shang
	End Sub

	Private Sub cmdTreat_ClickEvent(sender As Object, e As EventArgs)
		Dim Objective As String
		Dim temp, Tr_Obj As Double
		Dim J As Short
		Objective = InputBox("Enter your treatment objective in mg/L:", AppName_For_Display_Long, lblData(9).Text)
		On Error GoTo Bad_Treament_Objective
		temp = CDbl(Objective)
		Tr_Obj = temp / CPM_Results.Component.InitialConcentration
		For J = 1 To CPM_Max_Points
			If J > 2 Then
				If (CPM_Results.C_Over_C0(J) >= Tr_Obj) And (CPM_Results.C_Over_C0(J - 1) < Tr_Obj) Then
					Treatment_Objective.T = (CPM_Results.T(J) - CPM_Results.T(J - 1)) / (CPM_Results.C_Over_C0(J) - CPM_Results.C_Over_C0(J - 1)) * (Tr_Obj - CPM_Results.C_Over_C0(J - 1)) + CPM_Results.T(J - 1)
					Treatment_Objective.C = ((CPM_Results.C_Over_C0(J) - CPM_Results.C_Over_C0(J - 1)) / (CPM_Results.T(J) - CPM_Results.T(J - 1)) * (Treatment_Objective.T - CPM_Results.T(J - 1)) + CPM_Results.C_Over_C0(J - 1)) * CPM_Results.Component.InitialConcentration
					GoTo Exit_Loop
				End If
			End If
		Next J
		Flag_TO = False
		lblData(15).Text = "N/A"
		lblData(16).Text = "N/A"
		lblData(17).Text = "N/A"
		lblData(18).Text = "N/A"
		Exit Sub
Exit_Loop:
		Flag_TO = True
		lblData(15).Text = Format_It(Treatment_Objective.T, 2)
		lblData(16).Text = Format_It(Treatment_Objective.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
		lblData(17).Text = Format_It(Treatment_Objective.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
		lblData(18).Text = Format_It(Treatment_Objective.C, 2)
		Exit Sub
Bad_Treament_Objective:
		Resume Exit_lblLegend_Click
Exit_lblLegend_Click:
	End Sub

	Private Sub cmdFile_ClickEvent(sender As Object, e As EventArgs)

	End Sub




	Private Sub cmdExcel_Click(sender As Object, e As EventArgs) Handles cmdExcel.Click
		PFPSDM_Excel = False
		CPHSDM_Excel = True
		frmExcelCurves.ShowDialog()
	End Sub

	Private Sub cmdSelect_Click(sender As Object, e As EventArgs)
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

	Private Sub cmdPrint_ClickEvent(sender As Object, e As EventArgs)
		Dim Printer As New Printer
		Dim Error_Code As Short
		Dim temp As String
		Dim H, W As Single
		Dim Eq1 As String
		Dim i As Short

		On Error GoTo Print_Error

		'---Print Graph ---------------------------------------------------
		'	'''    H = grpBreak.Height
		'	'''    W = grpBreak.Width
		'	'''
		'	'''    grpBreak.Visible = False 'Hide it before printing

		'
		' THIS CODE HAD TO BE REPLACED TODAY, 1999-MAY-11, EJOMAN.
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

		'MsgBox _
		'"Printer.Height = " & Trim$(Str$(Printer.Height)) & ", " & _
		'"Printer.Width = " & Trim$(Str$(Printer.Width)) & ", " & _
		'"Printer.ScaleHeight = " & Trim$(Str$(Printer.ScaleHeight)) & ", " & _
		'"Printer.ScaleWidth = " & Trim$(Str$(Printer.ScaleWidth)) & ", " & _
		'"Printer.ScaleLeft = " & Trim$(Str$(Printer.ScaleLeft)) & ", " & _
		'"Printer.ScaleTop = " & Trim$(Str$(Printer.ScaleTop)) & ", "

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
		'	'''    Printer.Line _
		''''        (0, 0)- _
		''''        (grpBreak.Width, grpBreak.Height), _
		''''        QBColor(0), _
		''''        B
		'---- NEW CODE ENDS.

		'	'''    grpBreak.Height = H
		'	'''    grpBreak.Width = W
		'	'''
		'	'''    grpBreak.Visible = True
		'	'''
		'	'''    grpBreak.PrintStyle = 2
		'	'''    grpBreak.DrawMode = 2

		'
		' A "SKIP TO NEXT PAGE" COMMAND HAD TO BE ADDED TO THE PRINTING
		' CODE TODAY, 1999-MAY-11, EJOMAN.
		'
		'---- NEW CODE STARTS HERE:
		'	'''    Printer.NewPage
		'---- NEW CODE ENDS.

		'---Print other results------------------------------------------
		Printer.ScaleLeft = -1080 'Set a 3/4-inch margin
		Printer.ScaleTop = -1080
		Printer.CurrentX = 0
		Printer.CurrentY = 0

		'-- Print Filename

		Printer.FontSize = 10
		Printer.Print("From Data File: " & Filename)
		Printer.Print()

		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Input data for the Constant Pattern Model")
		Printer.FontSize = 10
		Printer.FontBold = False
		Printer.FontUnderline = True
		Printer.Print()
		Printer.Print("Chemical:", TAB(10), Trim(CPM_Results.Component.Name))
		Printer.FontUnderline = False
		Printer.Print(TAB(5), "Molecular weight: ", TAB(28), VB6.Format(CPM_Results.Component.MW, "0.00") & " g/mol")
		Printer.Print(TAB(5), "Normal Boiling Point: ", TAB(28), VB6.Format(CPM_Results.Component.BP, "0.00") & " C")
		Printer.Print(TAB(5), "Molar Volume @ NBP: ", TAB(28), Format_It(CPM_Results.Component.MolarVolume, 2) & " cm" & Chr(179) & "/mol")
		Printer.Print(TAB(5), "Initial Concentration: ", TAB(28), Format_It(CPM_Results.Component.InitialConcentration, 2) & " mg/L")
		Printer.Print(TAB(5), "K: ", TAB(28), VB6.Format(CPM_Results.Component.Use_K, "0.000") & " (mg/g)(L/mg)^(1/n)")
		Printer.Print(TAB(5), "1/n: ", TAB(28), VB6.Format(CPM_Results.Component.Use_OneOverN, "0.000"))
		Printer.Print()

		'-----------------------Bed Data ----------------------
		Printer.FontUnderline = True
		Printer.Print("Bed Data:")
		Printer.FontUnderline = False

		Printer.Print(TAB(5), "Bed Length: ", TAB(28), VB6.Format(CPM_Results.Bed.length, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(CPM_Results.Bed.Diameter, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(CPM_Results.Bed.Weight, "0.000E+00") & " kg")
		Printer.Print(TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(CPM_Results.Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		Printer.Print(TAB(5), "EBCT: ", TAB(28), VB6.Format(CPM_Results.Bed.length * PI * CPM_Results.Bed.Diameter * CPM_Results.Bed.Diameter / 4.0# / CPM_Results.Bed.Flowrate / 60.0#, "0.000E+00") & " mn")
		Printer.Print()
		Printer.Print(TAB(5), "Temperature:", TAB(28), VB6.Format(CPM_Results.Bed.Temperature, "0.00") & " C")
		If CPM_Results.Bed.Phase = 0 Then
			Printer.Print(TAB(5), "Water Density:", TAB(28), VB6.Format(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Water Viscosity:", TAB(28), VB6.Format(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			Printer.Print(TAB(5), "Pressure:", TAB(28), VB6.Format(CPM_Results.Bed.Pressure, "0.00000") & " atm")
			Printer.Print(TAB(5), "Air Density:", TAB(28), VB6.Format(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Air Viscosity:", TAB(28), VB6.Format(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		Printer.Print()

		'-----------------Carbon Properties -------------------------------
		Printer.FontUnderline = True
		Printer.Print("Carbon Properties:")
		Printer.FontUnderline = False

		Printer.Print(TAB(5), "Name: ", TAB(28), Trim(CPM_Results.Carbon.Name))
		Printer.Print(TAB(5), "Apparent Density: ", TAB(28), VB6.Format(CPM_Results.Carbon.Density, "0.000") & " g/cm" & Chr(179))
		Printer.Print(TAB(5), "Particle Radius: ", TAB(28), VB6.Format(CPM_Results.Carbon.ParticleRadius * 100.0#, "0.000000") & " cm")
		Printer.Print(TAB(5), "Porosity: ", TAB(28), VB6.Format(CPM_Results.Carbon.Porosity, "0.000"))
		Printer.Print(TAB(5), "Shape Factor: ", TAB(28), VB6.Format(CPM_Results.Carbon.ShapeFactor, "0.000"))
		'Printer.Print Tab(5); "Tortuosity: "; Tab(28); Format$(CPM_Results.Carbon.Tortuosity, "0.000")
		Printer.Print()

		'---------------Kinetic Parameters -----------------------------------------
		Printer.FontUnderline = True
		Printer.Print("Kinetic parameters:")
		Printer.FontUnderline = False
		Printer.Print(TAB(5), "kf", TAB(28), Format_It(CPM_Results.Component.kf, 2) & " cm/s")
		Printer.Print(TAB(5), "Ds", TAB(28), Format_It(CPM_Results.Component.Ds, 2) & " cm" & Chr(178) & "/s")
		Printer.Print(TAB(5), "SPDFR", TAB(28), Format_It(CPM_Results.Component.SPDFR, 2))

		'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Component(0) = CPM_Results.Component
		Printer.Print(TAB(5), "St", TAB(28), Format_It(ST(0), 2))
		'UPGRADE_WARNING: Couldn't resolve default property of object Eds(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Printer.Print(TAB(5), "Eds", TAB(28), Format_It(Eds(0), 2))

		Printer.Print()

		'Fouling-----------------------------------------
		Printer.FontUnderline = True
		Printer.Print("Fouling correlations:")
		Printer.FontUnderline = False
		Printer.Print()

		Printer.Print()
		Printer.Print(" Water type : " & Trim(CPM_Results.Bed.Water_Correlation.Name))
		Eq1 = VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(1), "0.00")

		If CPM_Results.Bed.Water_Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
		Else
			If CPM_Results.Bed.Water_Correlation.Coeff(2) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(CPM_Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
			End If
		End If
		If CPM_Results.Bed.Water_Correlation.Coeff(3) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
		Else
			If CPM_Results.Bed.Water_Correlation.Coeff(3) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(CPM_Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
			End If
		End If
		If CPM_Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
			If CPM_Results.Bed.Water_Correlation.Coeff(4) > 0 Then
				Eq1 = Eq1 & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
			Else
				If CPM_Results.Bed.Water_Correlation.Coeff(4) < 0 Then
					Eq1 = Eq1 & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
				End If
			End If
		End If
		Printer.Print("K(t)/K0 = " & Eq1)
		Printer.Print("(t in minutes)")
		Printer.Print()


		Eq1 = ""
		If CPM_Results.Component.Correlation.Coeff(1) = 1.0# Then
			Eq1 = "(K/K0) "
		Else
			If CPM_Results.Component.Correlation.Coeff(1) <> 0 Then Eq1 = VB6.Format(CPM_Results.Component.Correlation.Coeff(1), "0.00") & " * (K/K0) "
		End If
		If CPM_Results.Component.Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & "+ " & VB6.Format(CPM_Results.Component.Correlation.Coeff(2), "0.00")
		Else
			If CPM_Results.Component.Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & VB6.Format(System.Math.Abs(CPM_Results.Component.Correlation.Coeff(2)), "0.00")
		End If
		If Trim(Eq1) = "" Then
			Eq1 = "K/K0"
		End If
		Printer.Print(Trim(CPM_Results.Component.Name) & ":")
		Printer.Print(TAB(10), "Correlation type: " & Trim(CPM_Results.Component.Correlation.Name))

		Printer.Print(TAB(10), "K/K0 = " & Eq1)

		Printer.Print()

		If (CPM_Results.Component.Use_Tortuosity_Correlation) Then
			If (CPM_Results.Component.Constant_Tortuosity) Then
				Printer.Print("Correlation used when SOC competition is important:")
				Printer.Print(" Tortuosity = 0.782 * EBCT^0.925 ")
			Else
				Printer.Print("Correlation used when NOM fouling is important:")
				Printer.Print(" Tortuosity = 1.0 if t< 70 days")
				Printer.Print(" Tortuosity = 0.334 + 6.610E-06 * EBCT")
			End If
		End If
		Printer.Print()

		'--------- CPM Results ----------------------------------
		Printer.Print()
		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Constant Pattern Model Results for " & Trim(CPM_Results.Component.Name) & ":")
		Printer.Print()
		Printer.FontSize = 10
		Printer.FontBold = False
		Printer.FontUnderline = False
		Printer.Print("Minimum Stanton number:", TAB(30), Format_It(CPM_Results.Par(1), 2))
		Printer.Print("Minimum EBCT:", TAB(30), Format_It(CPM_Results.Par(2), 2) & " min")
		Printer.Print("Minimum Column Length:", TAB(30), Format_It(CPM_Results.Par(3), 2) & " cm")
		Printer.Print("MTZ Length:", TAB(30), Format_It(CPM_Results.Par(7), 2) & " cm")

		Printer.Print()
		Printer.Print(TAB(30), "Time(days)", TAB(40), "BVT", TAB(50), "TC", TAB(60), "C (mg/L)")
		Printer.Print("5% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_05.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_05.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_05.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_05.C, 2))
		Printer.Print("50% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_50.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_50.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_50.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_50.C, 2))
		Printer.Print("95% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_95.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_95.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_95.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_95.C, 2))
		Printer.Print()
		Printer.Print("TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")
		Printer.Print()

		If Flag_TO Then
			Printer.Print("Treatment Objective: " & Format_It(Treatment_Objective.C, 2) & " mg/L")
			Printer.Print()
			Printer.Print("Time (days):", TAB(20), Format_It(Treatment_Objective.T, 2))
			Printer.Print("BVT:", TAB(20), Format_It(Treatment_Objective.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2))
			Printer.Print("Tr. Capacity:", TAB(20), Format_It(Treatment_Objective.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2))
		Else
			Printer.Print("The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective.C, 2) & "mg/L) could not be calculated.")
		End If
		Printer.EndDoc()
		Exit Sub

Print_Error:
		Call Show_Trapped_Error("cmdPrint_Click")
		Resume Exit_Print
Exit_Print:

	End Sub

	Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
		Call cmdSave_Click()
	End Sub


	Private Sub cmdPrint_Click(sender As Object, e As EventArgs)
		cmdPrint_Click()
	End Sub

	Private Sub cmdFile_Click(sender As Object, e As EventArgs) Handles cmdFile.Click
		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNOverwritePrompt As Object
		Dim f, Error_Code As Short
		Dim temp As String
		Dim J, i, k As Short
		Dim Eq1 As String
		Dim Filename_Input As String

		On Error GoTo File_Error
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

		'f = FileNameIsValid(Filename_Input, CMDialog1)
		'If Not (f) Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'Filename_Input = CMDialog1.FileName
		Filename_Input = SaveFileDialog1.FileName

		f = FreeFile()
		FileOpen(f, Filename_Input, OpenMode.Output)

		PrintLine(f, "Input data for the Constant Pattern Model")
		'-- Print Filename

		PrintLine(f)
		PrintLine(f, "From Data File :", Filename)


		PrintLine(f)
		PrintLine(f, "Chemical:", TAB(10), Trim(CPM_Results.Component.Name))
		PrintLine(f, TAB(5), "Molecular weight: ", TAB(28), VB6.Format(CPM_Results.Component.MW, "0.00") & " g/mol")
		PrintLine(f, TAB(5), "Normal Boiling Point: ", TAB(28), VB6.Format(CPM_Results.Component.BP, "0.00") & " C")
		PrintLine(f, TAB(5), "Molar Volume @ NBP: ", TAB(28), Format_It(CPM_Results.Component.MolarVolume, 2) & " cm" & Chr(179) & "/mol")
		PrintLine(f, TAB(5), "Initial Concentration: ", TAB(28), Format_It(CPM_Results.Component.InitialConcentration, 2) & " mg/L")
		PrintLine(f, TAB(5), "K: ", TAB(28), VB6.Format(CPM_Results.Component.Use_K, "0.000") & " (mg/g)(L/mg)^(1/n)")
		PrintLine(f, TAB(5), "1/n: ", TAB(28), VB6.Format(CPM_Results.Component.Use_OneOverN, "0.000"))
		PrintLine(f)

		'-----------------------Bed Data ----------------------
		PrintLine(f, "Bed Data:")

		PrintLine(f, TAB(5), "Bed Length: ", TAB(28), VB6.Format(CPM_Results.Bed.length, "0.000E+00") & " m")
		PrintLine(f, TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(CPM_Results.Bed.Diameter, "0.000E+00") & " m")
		PrintLine(f, TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(CPM_Results.Bed.Weight, "0.000E+00") & " kg")
		PrintLine(f, TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(CPM_Results.Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		PrintLine(f, TAB(5), "EBCT: ", TAB(28), VB6.Format(CPM_Results.Bed.length * PI * CPM_Results.Bed.Diameter * CPM_Results.Bed.Diameter / 4.0# / CPM_Results.Bed.Flowrate / 60.0#, "0.000E+00") & " mn")
		PrintLine(f)
		PrintLine(f, TAB(5), "Temperature:", TAB(28), VB6.Format(CPM_Results.Bed.Temperature, "0.00") & " C")
		If CPM_Results.Bed.Phase = 0 Then
			PrintLine(f, TAB(5), "Water Density:", TAB(28), VB6.Format(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			PrintLine(f, TAB(5), "Water Viscosity:", TAB(28), VB6.Format(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			PrintLine(f, TAB(5), "Pressure:", TAB(28), VB6.Format(CPM_Results.Bed.Pressure, "0.00000") & " atm")
			PrintLine(f, TAB(5), "Air Density:", TAB(28), VB6.Format(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			PrintLine(f, TAB(5), "Air Viscosity:", TAB(28), VB6.Format(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		PrintLine(f)

		'-----------------Carbon Properties -------------------------------
		PrintLine(f, "Carbon Properties:")

		PrintLine(f, TAB(5), "Name: ", TAB(28), Trim(CPM_Results.Carbon.Name))
		PrintLine(f, TAB(5), "Apparent Density: ", TAB(28), VB6.Format(CPM_Results.Carbon.Density, "0.000") & " g/cm" & Chr(179))
		PrintLine(f, TAB(5), "Particle Radius: ", TAB(28), VB6.Format(CPM_Results.Carbon.ParticleRadius * 100.0#, "0.000000") & " cm")
		PrintLine(f, TAB(5), "Porosity: ", TAB(28), VB6.Format(CPM_Results.Carbon.Porosity, "0.000"))
		PrintLine(f, TAB(5), "Shape Factor: ", TAB(28), VB6.Format(CPM_Results.Carbon.ShapeFactor, "0.000"))
		'Print #f, Tab(5); "Tortuosity: "; Tab(28); Format$(CPM_Results.Carbon.Tortuosity, "0.000")
		PrintLine(f)

		'---------------Kinetic Parameters -----------------------------------------
		PrintLine(f, "Kinetic parameters:")
		PrintLine(f, TAB(5), "kf", TAB(28), Format_It(CPM_Results.Component.kf, 2) & " cm/s")
		PrintLine(f, TAB(5), "Ds", TAB(28), Format_It(CPM_Results.Component.Ds, 2) & " cm" & Chr(178) & "/s")
		PrintLine(f, TAB(5), "SPDFR", TAB(28), Format_It(CPM_Results.Component.SPDFR, 2))

		'UPGRADE_WARNING: Couldn't resolve default property of object Component(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Component(0) = CPM_Results.Component
		PrintLine(f, TAB(5), "St", TAB(28), Format_It(ST(0), 2))
		'UPGRADE_WARNING: Couldn't resolve default property of object Eds(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(f, TAB(5), "Eds", TAB(28), Format_It(Eds(0), 2))

		PrintLine(f)

		'Fouling-----------------------------------------
		PrintLine(f, "Fouling correlations:")
		PrintLine(f)
		PrintLine(f)
		PrintLine(f, " Water type : " & Trim(CPM_Results.Bed.Water_Correlation.Name))
		Eq1 = VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(1), "0.00")

		If CPM_Results.Bed.Water_Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
		Else
			If CPM_Results.Bed.Water_Correlation.Coeff(2) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(CPM_Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
			End If
		End If
		If CPM_Results.Bed.Water_Correlation.Coeff(3) > 0 Then
			Eq1 = Eq1 & " + " & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
		Else
			If CPM_Results.Bed.Water_Correlation.Coeff(3) < 0 Then
				Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(CPM_Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
			End If
		End If
		If CPM_Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
			If CPM_Results.Bed.Water_Correlation.Coeff(4) > 0 Then
				Eq1 = Eq1 & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
			Else
				If CPM_Results.Bed.Water_Correlation.Coeff(4) < 0 Then
					Eq1 = Eq1 & VB6.Format(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
				End If
			End If
		End If
		PrintLine(f, "K(t)/K0 = " & Eq1)
		PrintLine(f, "(t in minutes)")
		PrintLine(f)

		Eq1 = ""
		If CPM_Results.Component.Correlation.Coeff(1) = 1.0# Then
			Eq1 = "(K/K0) "
		Else
			If CPM_Results.Component.Correlation.Coeff(1) <> 0 Then Eq1 = VB6.Format(CPM_Results.Component.Correlation.Coeff(1), "0.00") & " * (K/K0) "
		End If
		If CPM_Results.Component.Correlation.Coeff(2) > 0 Then
			Eq1 = Eq1 & "+ " & VB6.Format(CPM_Results.Component.Correlation.Coeff(2), "0.00")
		Else
			If CPM_Results.Component.Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & VB6.Format(System.Math.Abs(CPM_Results.Component.Correlation.Coeff(2)), "0.00")
		End If
		If Trim(Eq1) = "" Then
			Eq1 = "K/K0"
		End If
		PrintLine(f, Trim(CPM_Results.Component.Name) & ":")
		PrintLine(f, TAB(10), "Correlation type: " & Trim(CPM_Results.Component.Correlation.Name))
		PrintLine(f, TAB(10), "K/K0 = " & Eq1)
		PrintLine(f)

		If (CPM_Results.Component.Use_Tortuosity_Correlation) Then
			If (CPM_Results.Component.Constant_Tortuosity) Then
				PrintLine(f, "Correlation used when SOC competition is important:")
				PrintLine(f, " Tortuosity = 0.782 * EBCT^0.925 ")
			Else
				PrintLine(f, "Correlation used when NOM fouling is important:")
				PrintLine(f, " Tortuosity = 1.0 if t< 70 days")
				PrintLine(f, " Tortuosity = 0.334 + 6.610E-06 * EBCT")
			End If
		End If
		PrintLine(f)

		'--------- CPM Results ----------------------------------
		PrintLine(f, "Constant Pattern Model Results for " & Trim(CPM_Results.Component.Name) & ":")
		PrintLine(f)
		PrintLine(f, "Minimum Stanton number:", TAB(30), Format_It(CPM_Results.Par(1), 2))
		PrintLine(f, "Minimum EBCT:", TAB(30), Format_It(CPM_Results.Par(2), 2) & " min")
		PrintLine(f, "Minimum Column Length:", TAB(30), Format_It(CPM_Results.Par(3), 2) & " cm")
		PrintLine(f, "Throughput at 95% of the MTZ:", TAB(30), Format_It(CPM_Results.Par(4), 2))
		PrintLine(f, "Throughput at 5% of the MTZ:", TAB(30), Format_It(CPM_Results.Par(5), 2))
		PrintLine(f, "EBCT of the MTZ:", TAB(30), Format_It(CPM_Results.Par(6), 2) & " min")
		PrintLine(f, "Length of the MTZ:", TAB(30), Format_It(CPM_Results.Par(7), 2) & " cm")

		PrintLine(f)
		PrintLine(f, TAB(30), "Time(days)", TAB(40), "BVT", TAB(50), "TC", TAB(60), "C (mg/L)")
		PrintLine(f, "5% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_05.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_05.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_05.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_05.C, 2))
		PrintLine(f, "50% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_50.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_50.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_50.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_50.C, 2))
		PrintLine(f, "95% of the influent conc.", TAB(30), Format_It(CPM_Results.ThroughPut_95.T, 2), TAB(40), Format_It(CPM_Results.ThroughPut_95.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2), TAB(50), Format_It(CPM_Results.ThroughPut_95.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2), TAB(60), Format_It(CPM_Results.ThroughPut_95.C, 2))
		PrintLine(f)
		PrintLine(f, "TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")
		PrintLine(f)

		If Flag_TO Then
			PrintLine(f, "Treatment Objective: " & Format_It(Treatment_Objective.C, 2) & " mg/L")
			PrintLine(f)
			PrintLine(f, "Time (days):", TAB(20), Format_It(Treatment_Objective.T, 2))
			PrintLine(f, "BVT:", TAB(20), Format_It(Treatment_Objective.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2))
			PrintLine(f, "Tr. Capacity:", TAB(20), Format_It(Treatment_Objective.T * 24.0# * 3600.0# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2))
		Else
			PrintLine(f, "The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective.C, 2) & "mg/L) could not be calculated.")
		End If
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

	Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
		Me.Close()
	End Sub

	Private Sub cmdTreat_Click(sender As Object, e As EventArgs) Handles cmdTreat.Click
		Call cmdTreat_Click()
	End Sub

	Private Sub frmModelCPHSDMResults_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class