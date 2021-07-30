Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports System.Windows.Forms.DataVisualization.Charting

Friend Class frmModelECMResults
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer


	Dim length As Double
	Dim NumW As Short
	'UPGRADE_WARNING: Lower bound of array Solid_ConcW was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim Solid_ConcW(Number_Compo_Max, Number_Compo_Max) As Object
	'UPGRADE_WARNING: Lower bound of array Liquid_ConcW was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim Liquid_ConcW(Number_Compo_Max, Number_Compo_Max) As Object
	'UPGRADE_WARNING: Lower bound of array CoCW was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim CoCW(Number_Compo_Max, Number_Compo_Max) As Double
	'UPGRADE_WARNING: Lower bound of array IndexW was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim IndexW(Number_Compo_Max) As Short
	'UPGRADE_ISSUE: Declaration type not supported: Array of fixed-length strings. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"'
	Dim Name_CompW(Number_Compo_Max) As String
	Dim Time_Break() As Double 'Breakthroug time for each chhemical
	Dim Time_Min As Double
	Dim Time_Unit As String
	
	Dim PopulatingScrollboxes As Short
	
	
	
	
	Const frmModelECMResults_declarations_end As Boolean = True
	
	
	'UPGRADE_WARNING: Event cboGlob.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboGlob_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboGlob.SelectedIndexChanged
		Dim i As Short
		
		If (Not PopulatingScrollboxes) Then
			If cboGlob.Text = "C/Co" Then i = 1
			If cboGlob.Text = "Q" Then i = 2
			If cboGlob.Text = "C (Liquid Conc.)" Then i = 3
			Call Draw(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.DrawMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'grpGlob.DrawMode = 2
		End If
		
	End Sub
	
	Private Sub cmdClose_Click()
		Me.Close()
	End Sub



	Private Sub cmdPrint_Click()
		Dim Printer As New Printer
		Dim Error_Code As Short
		Dim temp As String
		Dim k, i, J As Short
		
		On Error GoTo Print_Error
		'---Print other results-----------------------------------------------
		Printer.ScaleLeft = -1080 'Set a 3/4-inch margin
		Printer.ScaleTop = -1080
		Printer.CurrentX = 0
		Printer.CurrentY = 0
		
		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Input data for the Equilibrium Colum Model")
		Printer.FontSize = 10
		Printer.FontBold = False
		Printer.FontUnderline = False
		'-- Print Filename
		Printer.Print()
		Printer.Print("From Data File : " & Filename)
		Printer.Print("Date/time stamp:" & DateString & " " & TimeString)
		
		Printer.Print()
		Printer.Print("Component", TAB(30), "K*", TAB(38), "1/n", TAB(45), "Init. Conc.", TAB(59), "MW")
		Printer.Print(TAB(39), "-", TAB(48), "mg/L", TAB(58), "g/mol")
		
		For i = 1 To Number_Component_ECM
			'K = Component_Index_ECM(i)
			k = IndexW(i)
			'      Printer.Print Trim$(Mid$(LTrim$(Component(K).Name), 1, 25)); Tab(29); Format$(Component(K).Use_K, "###,##0.000"); Tab(37); Format$(Component(K).Use_OneOverN, "0.000"); Tab(48); Format_It(Component(K).InitialConcentration, 2); Tab(58); Format$(Component(K).MW, "0.00")
			Printer.Print(Trim(Mid(LTrim(Component(k).Name), 1, 25)), TAB(29), VB6.Format(Component(k).Use_K, "###,##0.000"), TAB(37), VB6.Format(Component(k).Use_OneOverN, "0.000"), TAB(48), Format_It(Component(k).InitialConcentration, 2), TAB(58), VB6.Format(Component(k).MW, "0.00"))
		Next i
		Printer.Print()
		Printer.Print("* K in (mg/g)*(L/mg)^(1/n)")
		
		Printer.Print()
		
		'-----------------------Bed Data ----------------------
		Printer.FontUnderline = True
		Printer.Print("Bed Data:")
		Printer.FontUnderline = False
		
		Printer.Print(TAB(5), "Bed Length: ", TAB(28), VB6.Format(Bed.length, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(Bed.Diameter, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(Bed.Weight, "0.000E+00") & " kg")
		Printer.Print(TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		Printer.Print(TAB(5), "EBCT: ", TAB(28), VB6.Format(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#, "0.000E+00") & " mn")
		Printer.Print()
		Printer.Print(TAB(5), "Temperature:", TAB(28), VB6.Format(Bed.Temperature, "0.00") & " C")
		If Bed.Phase = 0 Then
			Printer.Print(TAB(5), "Water Density:", TAB(28), VB6.Format(Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Water Viscosity:", TAB(28), VB6.Format(Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			Printer.Print(TAB(5), "Pressure:", TAB(28), VB6.Format(Bed.Pressure, "0.00000") & " atm")
			Printer.Print(TAB(5), "Air Density:", TAB(28), VB6.Format(Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Air Viscosity:", TAB(28), VB6.Format(Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		Printer.Print()
		
		'-----------------Carbon Properties -------------------------------
		Printer.FontUnderline = True
		Printer.Print("Carbon Properties:")
		Printer.FontUnderline = False
		
		Printer.Print(TAB(5), "Name: ", TAB(28), Trim(Carbon.Name))
		Printer.Print(TAB(5), "Apparent Density: ", TAB(28), VB6.Format(Carbon.Density, "0.000") & " g/cm" & Chr(179))
		Printer.Print(TAB(5), "Particle Radius: ", TAB(28), VB6.Format(Carbon.ParticleRadius * 100#, "0.000000") & " cm")
		Printer.Print(TAB(5), "Porosity: ", TAB(28), VB6.Format(Carbon.Porosity, "0.000"))
		Printer.Print(TAB(5), "Shape Factor: ", TAB(28), VB6.Format(Carbon.ShapeFactor, "0.000"))
		'Printer.Print Tab(5); "Tortuosity: "; Tab(28); Format$(Carbon.Tortuosity, "0.000")
		Printer.Print()
		
		
		Printer.Print()
		'--- Print the results from the table
		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Results for the Equilibrium Column Model")
		Printer.FontUnderline = False
		Printer.FontSize = 10
		Printer.FontBold = False
		
		Printer.Print()
		Printer.Print("Zone", TAB(9), "Component", TAB(35), "BVF", TAB(44), "Wave Vel.", TAB(54), "TC", TAB(63), "Breakthrough")
		Printer.Print(TAB(45), "cm/s", TAB(54), "m3/kg", TAB(63), Time_Unit)
		For i = 1 To Number_Component_ECM
			Printer.Print("Zone " & VB6.Format(i, "0"), TAB(9), Mid(Trim(Component(IndexW(i)).Name), 1, 25), TAB(35), Format_It(Output_ECM(i).Bed_Volume_Fed, 2), TAB(45), Format_It(Output_ECM(i).Wave_Velocity, 2), TAB(54), Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2), TAB(63), Format_It(Time_Break(i), 2))
			
			'Change made: (ejo, 3/1/96)
			'==========================
			'was: Format_It(Output_ECM(i).Carbon_Usage_Rate, 2)
			'is now: Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2)
			
		Next i
		Printer.Print()
		Printer.Print("TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")
		Printer.Print()
		
		For i = 1 To Number_Component_ECM
			Printer.FontBold = True
			Printer.Print(Mid(Trim(Component(IndexW(i)).Name), 1, 25))
			Printer.FontBold = False
			Printer.Print("Zone ", TAB(9), "C/Co", TAB(19), "C (mg/L)", TAB(29), "Q (mg/L)")
			For J = 1 To Number_Component_ECM
				'UPGRADE_WARNING: Couldn't resolve default property of object Solid_ConcW(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Liquid_ConcW(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Printer.Print("Zone " & VB6.Format(J, "0"), TAB(9), Format_It(CoCW(i, J), 2), TAB(19), Format_It(Liquid_ConcW(i, J) / 1000#, 2), TAB(29), Format_It(Solid_ConcW(i, J) / 1000#, 2))
			Next J
			Printer.Print()
		Next i
		
		Printer.Print()
		'--- Print the mass balance results
		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Mass Balance Results")
		Printer.FontUnderline = False
		Printer.FontSize = 10
		Printer.FontBold = False
		
		Printer.Print()
		'Printer.Print "Component"; Tab(30); "Left-Hand"; Tab(45); "Right-Hand"; Tab(60); "Percent Err."
		'Printer.Print ""; Tab(30); "(ug/cm2/s)"; Tab(45); "(ug/cm2/s)"; Tab(60); "(%)"
		Printer.Print("Component", TAB(30), "Percent Err.")
		Printer.Print("", TAB(30), "(%)")
		For i = 1 To Number_Component_ECM
			'      Printer.Print Mid$(Trim$(Component(i).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(i), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(i), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(i), 3)
			'      Printer.Print Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(IndexW(i)), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(IndexW(i)), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3)
			Printer.Print(Mid(Trim(Component(IndexW(i)).Name), 1, 25), TAB(30), Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3))
		Next i
		Printer.Print()
		
		
		Printer.EndDoc()
		Exit Sub
Print_Error: 
		Call Show_Trapped_Error("cmdPrint_Click")
		Resume Exit_Print
Exit_Print: 
		
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
	
	
	Private Sub Draw(ByRef GFlag As Short)
		Dim Num_Compo As Short
		Num_Compo = Number_Component_ECM
		Dim J, i, k, Numpoints As Short

		Dim s(NumW) As Series

		Chart1.Titles.Clear()
		Chart1.Series.Clear()
		Chart1.ChartAreas(0).RecalculateAxesScale()

		Select Case GFlag
			Case 1
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.NumSets = NumW

				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GraphType = 4
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GraphStyle = 0

				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For J = 1 To NumW
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisSet = J
					s(J) = New Series
					s(J).ChartType = SeriesChartType.Column
					If Num_Compo >= 2 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Numpoints = NumW
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.NumPoints = 2
						Numpoints = 2
					End If
				Next J
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GraphTitle = " C/Co for All Components"
				Chart1.Titles.Add("C/Co for All Components")
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GridStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GridStyle = 3

				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For J = 1 To NumW
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisSet = J
					'K = IndexW(NumW - J + 1)
					'k = J
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisPoint = J
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.LegendText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.LegendText = Name_CompW(J)
					s(J).LegendText = Name_CompW(J)
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For i = 1 To Numpoints
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.LabelText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

						'grpGlob.LabelText = "Zone " & VB6.Format(i, "0")


						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.GraphData = CoCW(k, i)
						If CoCW(J, i) = 0 Then
							s(J).Points.Add(0.0001)
						Else
							s(J).Points.Add(CoCW(J, i))
						End If


					Next i
				Next J
			Case 2
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.NumSets = NumW
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GraphType = 3
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GraphStyle = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GraphTitle = " Q (" & Chr(181) & "g/g) for All Components"
				Chart1.Titles.Add(" Q (" & Chr(181) & "g/g) for All Components")

				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GridStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GridStyle = 3

				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For J = 1 To NumW 'numsets
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisSet = J
					If NumW > 2 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.NumPoints = NumW
						Numpoints = NumW
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.NumPoints = 2
						Numpoints = 2
					End If
				Next J

				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For J = 1 To NumW 'grpGlob.NumSets
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisSet = J
					'K = IndexW(NumW - J + 1)
					'k = J
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisPoint = J
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.LegendText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					s(J) = New Series
					s(J).LegendText = Name_CompW(J)

					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For i = 1 To NumW 'grpGlob.NumPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.ThisPoint = i

						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.LabelText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

						'grpGlob.LabelText = "Zone " & VB6.Format(i, "0")

						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Solid_ConcW(k, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Solid_ConcW(J, i) = 0 Then
							s(J).Points.Add(0.0001)
						Else
							s(J).Points.Add(Solid_ConcW(J, i))
						End If
					Next i
				Next J

			Case 3
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.NumSets = NumW
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GraphType = 3
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GraphStyle = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Chart1.Titles.Add(" C (" & Chr(181) & "g/L) for All Components")
				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GridStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'grpGlob.GridStyle = 3

				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For J = 1 To NumW 'grpGlob.NumSets
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisSet = J
					If NumW > 2 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Numpoints = NumW
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.NumPoints = 2
						Numpoints = 2
					End If
				Next J

				'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For J = 1 To NumW 'grpGlob.NumSets
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisSet = J
					'K = IndexW(NumW - J + 1)
					'k = J
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.ThisPoint = J
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.LegendText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'grpGlob.LegendText = Name_CompW(k)
					s(J) = New Series
					s(J).LegendText = Name_CompW(J)
					'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.NumPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For i = 1 To NumW ' grpGlob.NumPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.LabelText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

						'grpGlob.LabelText = "Zone " & VB6.Format(i, "0")

						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'grpGlob.ThisPoint = i
						'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.GraphData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Liquid_ConcW(k, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

						If Liquid_ConcW(J, i) = 0 Then
							s(J).Points.Add(0.0001)
						Else
							s(J).Points.Add(Liquid_ConcW(J, i))
						End If

					Next i
				Next J
		End Select
		'If Number_Component_ECM = 1 Then
		'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.ThisPoint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpGlob.ThisPoint = 2
		'UPGRADE_WARNING: Couldn't resolve default property of object grpGlob.LabelText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'grpGlob.LabelText = ""
		'End If
		For J = 1 To NumW
			Chart1.Series.Add(s(J))
		Next J

		Chart1.ChartAreas(0).AxisX.Title = "Zone:"
		Chart1.ChartAreas(0).AxisX.LabelStyle.Font = New System.Drawing.Font("Times New Roman", 10.25F)
		Chart1.ChartAreas(0).AxisY.LabelStyle.Font = New System.Drawing.Font("Times New Roman", 10.25F)

	End Sub
	
	Private Sub frmModelECMResults_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i, J As Short
		rs.FindAllControls(Me)

		'Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmGlobal.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmGlobal.Height / 2)
		Call CenterOnForm(Me, frmMain)
		
		Label1.Text = "VTM" & Chr(13) & "(m" & Chr(179) & "/kg)"
		'Me.HelpContextID = Hlp_Global_Results
		''''Caption = "Results for the Equilibrium Column Model"
		NumW = Number_Component_ECM
		For i = 1 To NumW
			IndexW(i) = Output_ECM(i).Index
			Name_CompW(i) = Component(IndexW(i)).Name
			For J = 1 To NumW
				'UPGRADE_WARNING: Couldn't resolve default property of object Solid_ConcW(i, J). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Solid_ConcW(i, J) = Output_ECM(i).Solid_Concentration(J)
				'UPGRADE_WARNING: Couldn't resolve default property of object Liquid_ConcW(i, J). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Liquid_ConcW(i, J) = Output_ECM(i).Liquid_Concentration(J)
				CoCW(i, J) = Output_ECM(i).C_Over_C0(J)
			Next J
		Next i
		
		''''fraGlob = "Results"
		lblZone.Text = ""
		lblData1.Text = ""
		lblData2.Text = ""
		lblData3.Text = ""
		lblData4.Text = ""
		lblData5.Text = ""
		lblCompo.Text = ""
		'UPGRADE_WARNING: Lower bound of array Time_Break was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim Time_Break(NumW)
		
		'Time_Min = 1E+100
		For i = 1 To NumW
			Time_Break(i) = Bed.length * 100# / Output_ECM(i).Wave_Velocity / 3600# / 24#
			'  If Time_Break(I) < Time_Min Then
			'   Time_Min = Time_Break(I)
			'  End If
			'Next I
			'If Time_Min < 1# Then
			' For I = 1 To NumW
			'   Time_Break(I) = Time_Break(I) * 24# * 60#
		Next i
		' Time_Unit = "mn"
		'Else Time_Unit = "days"
		'End If
		Time_Unit = "days"
		
		Label6.Text = "Breakthrough time(" & Time_Unit & ")"
		For i = 1 To NumW
			lblZone.Text = lblZone.Text & VB6.Format(i, "0") & Chr(10)
			lblCompo.Text = lblCompo.Text & LCase(Trim(Component(IndexW(i)).Name) & Chr(10))
			lblData1.Text = lblData1.Text & Format_It(Output_ECM(i).Bed_Volume_Fed, 2) & Chr(10)
			lblData2.Text = lblData2.Text & Format_It(Output_ECM(i).Wave_Velocity, 2) & Chr(10)
			lblData3.Text = lblData3.Text & Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2) & Chr(10)
			lblData4.Text = lblData4.Text & Format_It(Time_Break(i), 3) & Chr(10)
			lblData5.Text = lblData5.Text & Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(i), 2) & Chr(10)
		Next i
		
		Call Populate_Scrollboxes()
		Call cboGlob_SelectedIndexChanged(cboGlob, New System.EventArgs())

		'Chart1.ChartAreas(0).
		'Call Draw(1)
		'grpGlob.DrawMode = 2


	End Sub
	
	Private Sub frmModelECMResults_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		Call UserPrefs_Save()
		
	End Sub
	
	
	Private Sub Populate_Scrollboxes()
		
		PopulatingScrollboxes = True
		
		cboGlob.Items.Clear()
		cboGlob.Items.Add("C/Co")
		cboGlob.Items.Add("Q")
		cboGlob.Items.Add("C (Liquid Conc.)")
		
		'-- Read in INI settings
		cboGlob.SelectedIndex = 0
		Call UserPrefs_Load()
		
		PopulatingScrollboxes = False
		
	End Sub
	
	Private Sub UserPrefs_Load()
		Dim X As Integer
		
		On Error GoTo err_FRMGLOBAL_UserPrefs_Load
		
		X = CInt(INI_Getsetting("FRMGLOBAL_cboGlob"))
		If ((X >= 0) And (X <= cboGlob.Items.Count - 1)) Then
			cboGlob.SelectedIndex = X
		End If
		Exit Sub
		
resume_err_FRMGLOBAL_UserPrefs_Load: 
		Call UserPrefs_Save()
		Exit Sub
		
err_FRMGLOBAL_UserPrefs_Load: 
		Resume resume_err_FRMGLOBAL_UserPrefs_Load
		
	End Sub
	
	Private Sub UserPrefs_Save()
		Dim X As Integer
		
		X = cboGlob.SelectedIndex
		Call INI_PutSetting("FRMGLOBAL_cboGlob", Trim(CStr(X)))
		
	End Sub

	Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
		Me.Dispose()
	End Sub

	Private Sub cmdFile_Click(sender As Object, e As EventArgs) Handles cmdFile.Click
		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNOverwritePrompt As Object
		Dim f, Error_Code As Short
		Dim temp As String
		Dim J, i, k As Short
		Dim Filename_Input As String

		On Error GoTo File_Error
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.CancelError = True
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.DialogTitle = "Print to File"
		SaveFileDialog1.Title = "Print to File"
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SaveFileDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Excel File (*.csv)|*.csv"
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SaveFileDialog1.FilterIndex = 3
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
		PrintLine(f, "Input data for the Equilibrium Colum Model")
		'-- Print Filename

		PrintLine(f)
		PrintLine(f, "From Data File : " & Filename)
		PrintLine(f, "Date/time stamp: " & DateString & " " & TimeString)

		PrintLine(f)
		PrintLine(f, "Component", TAB(30), "K*", TAB(38), "1/n", TAB(45), "Init. Conc.", TAB(59), "MW")
		PrintLine(f, TAB(39), "-", TAB(48), "mg/L", TAB(58), "g/mol")

		For i = 1 To Number_Component_ECM
			'K = Component_Index_ECM(i)
			k = IndexW(i)
			'        Print #f, Trim$(Mid$(LTrim$(Component(K).Name), 1, 25)); Tab(29); Format$(Component(K).Use_K, "###,##0.000"); Tab(37); Format$(Component(K).Use_OneOverN, "0.000"); Tab(48); Format_It(Component(K).InitialConcentration, 2); Tab(58); Format$(Component(K).MW, "0.00")
			PrintLine(f, Trim(Mid(LTrim(Component(k).Name), 1, 25)), TAB(29), VB6.Format(Component(k).Use_K, "###,##0.000"), TAB(37), VB6.Format(Component(k).Use_OneOverN, "0.000"), TAB(48), Format_It(Component(k).InitialConcentration, 2), TAB(58), VB6.Format(Component(k).MW, "0.00"))
		Next i
		PrintLine(f)
		PrintLine(f, "* K in (mg/g)*(L/mg)^(1/n)")

		PrintLine(f)

		'-----------------------Bed Data ----------------------
		PrintLine(f, "Bed Data:")

		PrintLine(f, TAB(5), "Bed Length: ", TAB(28), VB6.Format(Bed.length, "0.000E+00") & " m")
		PrintLine(f, TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(Bed.Diameter, "0.000E+00") & " m")
		PrintLine(f, TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(Bed.Weight, "0.000E+00") & " kg")
		PrintLine(f, TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		PrintLine(f, TAB(5), "EBCT: ", TAB(28), VB6.Format(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4.0# / Bed.Flowrate / 60.0#, "0.000E+00") & " mn")
		PrintLine(f)
		PrintLine(f, TAB(5), "Temperature:", TAB(28), VB6.Format(Bed.Temperature, "0.00") & " C")
		If Bed.Phase = 0 Then
			PrintLine(f, TAB(5), "Water Density:", TAB(28), VB6.Format(Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			PrintLine(f, TAB(5), "Water Viscosity:", TAB(28), VB6.Format(Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			PrintLine(f, TAB(5), "Pressure:", TAB(28), VB6.Format(Bed.Pressure, "0.00000") & " atm")
			PrintLine(f, TAB(5), "Air Density:", TAB(28), VB6.Format(Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			PrintLine(f, TAB(5), "Air Viscosity:", TAB(28), VB6.Format(Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		PrintLine(f)

		'-----------------Carbon Properties -------------------------------
		PrintLine(f, "Carbon Properties:")

		PrintLine(f, TAB(5), "Name: ", TAB(28), Trim(Carbon.Name))
		PrintLine(f, TAB(5), "Apparent Density: ", TAB(28), VB6.Format(Carbon.Density, "0.000") & " g/cm" & Chr(179))
		PrintLine(f, TAB(5), "Particle Radius: ", TAB(28), VB6.Format(Carbon.ParticleRadius * 100.0#, "0.000000") & " cm")
		PrintLine(f, TAB(5), "Porosity: ", TAB(28), VB6.Format(Carbon.Porosity, "0.000"))
		PrintLine(f, TAB(5), "Shape Factor: ", TAB(28), VB6.Format(Carbon.ShapeFactor, "0.000"))
		'Print #f, Tab(5); "Tortuosity: "; Tab(28); Format$(Carbon.Tortuosity, "0.000")
		PrintLine(f)

		PrintLine(f)
		'--- Print the results from the table
		PrintLine(f, "Results for the Equilibrium Column Model")

		PrintLine(f)
		PrintLine(f, "Zone", TAB(9), "Component", TAB(35), "BVF", TAB(44), "Wave Vel.", TAB(54), "TC", TAB(63), "Breakthrough")
		PrintLine(f, TAB(45), "cm/s", TAB(54), "m3/kg", TAB(63), Time_Unit)
		For i = 1 To Number_Component_ECM
			PrintLine(f, "Zone " & VB6.Format(i, "0"), TAB(9), Mid(Trim(Component(IndexW(i)).Name), 1, 25), TAB(35), Format_It(Output_ECM(i).Bed_Volume_Fed, 2), TAB(45), Format_It(Output_ECM(i).Wave_Velocity, 2), TAB(54), Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2), TAB(63), Format_It(Time_Break(i), 2))

			'Change made: (ejo, 3/1/96)
			'==========================
			'was: Format_It(Output_ECM(i).Carbon_Usage_Rate, 2)
			'is now: Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2)

		Next i
		PrintLine(f)
		PrintLine(f, "TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")
		PrintLine(f)

		For i = 1 To Number_Component_ECM
			PrintLine(f, Mid(Trim(Component(IndexW(i)).Name), 1, 25))
			PrintLine(f, "Zone ", TAB(9), "C/Co", TAB(19), "C (mg/L)", TAB(29), "Q (mg/L)")
			For J = 1 To Number_Component_ECM
				'UPGRADE_WARNING: Couldn't resolve default property of object Solid_ConcW(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Liquid_ConcW(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(f, "Zone " & VB6.Format(J, "0"), TAB(9), Format_It(CoCW(i, J), 2), TAB(19), Format_It(Liquid_ConcW(i, J) / 1000.0#, 2), TAB(29), Format_It(Solid_ConcW(i, J) / 1000.0#, 2))
			Next J
			PrintLine(f)
		Next i

		PrintLine(f)
		PrintLine(f)
		'--- Print the mass balance results
		PrintLine(f, "Mass Balance Results")
		PrintLine(f)
		PrintLine(f, "Component", TAB(30), "Percent Err.")
		'Print #f, ""; Tab(30); "(ug/cm2/s)"; Tab(45); "(ug/cm2/s)"; Tab(60); "(%)"
		PrintLine(f, "", TAB(30), "(%)")
		For i = 1 To Number_Component_ECM
			'        Print #f, Mid$(Trim$(Component(i).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(IndexW(i)), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(IndexW(i)), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3)
			'        Print #f, Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(i), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(i), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(i), 3)
			PrintLine(f, Mid(Trim(Component(IndexW(i)).Name), 1, 25), TAB(30), Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3))
		Next i
		PrintLine(f)


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

	Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
		Dim Printer As New Printer
		Dim Error_Code As Short
		Dim temp As String
		Dim k, i, J As Short

		On Error GoTo Print_Error
		'---Print other results-----------------------------------------------
		Printer.ScaleLeft = -1080 'Set a 3/4-inch margin
		Printer.ScaleTop = -1080
		Printer.CurrentX = 0
		Printer.CurrentY = 0

		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Input data for the Equilibrium Colum Model")
		Printer.FontSize = 10
		Printer.FontBold = False
		Printer.FontUnderline = False
		'-- Print Filename
		Printer.Print()
		Printer.Print("From Data File : " & Filename)
		Printer.Print("Date/time stamp:" & DateString & " " & TimeString)

		Printer.Print()
		Printer.Print("Component", TAB(30), "K*", TAB(38), "1/n", TAB(45), "Init. Conc.", TAB(59), "MW")
		Printer.Print(TAB(39), "-", TAB(48), "mg/L", TAB(58), "g/mol")

		For i = 1 To Number_Component_ECM
			'K = Component_Index_ECM(i)
			k = IndexW(i)
			'      Printer.Print Trim$(Mid$(LTrim$(Component(K).Name), 1, 25)); Tab(29); Format$(Component(K).Use_K, "###,##0.000"); Tab(37); Format$(Component(K).Use_OneOverN, "0.000"); Tab(48); Format_It(Component(K).InitialConcentration, 2); Tab(58); Format$(Component(K).MW, "0.00")
			Printer.Print(Trim(Mid(LTrim(Component(k).Name), 1, 25)), TAB(29), VB6.Format(Component(k).Use_K, "###,##0.000"), TAB(37), VB6.Format(Component(k).Use_OneOverN, "0.000"), TAB(48), Format_It(Component(k).InitialConcentration, 2), TAB(58), VB6.Format(Component(k).MW, "0.00"))
		Next i
		Printer.Print()
		Printer.Print("* K in (mg/g)*(L/mg)^(1/n)")

		Printer.Print()

		'-----------------------Bed Data ----------------------
		Printer.FontUnderline = True
		Printer.Print("Bed Data:")
		Printer.FontUnderline = False

		Printer.Print(TAB(5), "Bed Length: ", TAB(28), VB6.Format(Bed.length, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Bed Diameter: ", TAB(28), VB6.Format(Bed.Diameter, "0.000E+00") & " m")
		Printer.Print(TAB(5), "Weight of GAC: ", TAB(28), VB6.Format(Bed.Weight, "0.000E+00") & " kg")
		Printer.Print(TAB(5), "Inlet Flowrate: ", TAB(28), VB6.Format(Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
		Printer.Print(TAB(5), "EBCT: ", TAB(28), VB6.Format(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4.0# / Bed.Flowrate / 60.0#, "0.000E+00") & " mn")
		Printer.Print()
		Printer.Print(TAB(5), "Temperature:", TAB(28), VB6.Format(Bed.Temperature, "0.00") & " C")
		If Bed.Phase = 0 Then
			Printer.Print(TAB(5), "Water Density:", TAB(28), VB6.Format(Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Water Viscosity:", TAB(28), VB6.Format(Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		Else
			Printer.Print(TAB(5), "Pressure:", TAB(28), VB6.Format(Bed.Pressure, "0.00000") & " atm")
			Printer.Print(TAB(5), "Air Density:", TAB(28), VB6.Format(Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
			Printer.Print(TAB(5), "Air Viscosity:", TAB(28), VB6.Format(Bed.WaterViscosity, "0.00E+00") & " g/cm.s")
		End If
		Printer.Print()

		'-----------------Carbon Properties -------------------------------
		Printer.FontUnderline = True
		Printer.Print("Carbon Properties:")
		Printer.FontUnderline = False

		Printer.Print(TAB(5), "Name: ", TAB(28), Trim(Carbon.Name))
		Printer.Print(TAB(5), "Apparent Density: ", TAB(28), VB6.Format(Carbon.Density, "0.000") & " g/cm" & Chr(179))
		Printer.Print(TAB(5), "Particle Radius: ", TAB(28), VB6.Format(Carbon.ParticleRadius * 100.0#, "0.000000") & " cm")
		Printer.Print(TAB(5), "Porosity: ", TAB(28), VB6.Format(Carbon.Porosity, "0.000"))
		Printer.Print(TAB(5), "Shape Factor: ", TAB(28), VB6.Format(Carbon.ShapeFactor, "0.000"))
		'Printer.Print Tab(5); "Tortuosity: "; Tab(28); Format$(Carbon.Tortuosity, "0.000")
		Printer.Print()


		Printer.Print()
		'--- Print the results from the table
		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Results for the Equilibrium Column Model")
		Printer.FontUnderline = False
		Printer.FontSize = 10
		Printer.FontBold = False

		Printer.Print()
		Printer.Print("Zone", TAB(9), "Component", TAB(35), "BVF", TAB(44), "Wave Vel.", TAB(54), "TC", TAB(63), "Breakthrough")
		Printer.Print(TAB(45), "cm/s", TAB(54), "m3/kg", TAB(63), Time_Unit)
		For i = 1 To Number_Component_ECM
			Printer.Print("Zone " & VB6.Format(i, "0"), TAB(9), Mid(Trim(Component(IndexW(i)).Name), 1, 25), TAB(35), Format_It(Output_ECM(i).Bed_Volume_Fed, 2), TAB(45), Format_It(Output_ECM(i).Wave_Velocity, 2), TAB(54), Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2), TAB(63), Format_It(Time_Break(i), 2))

			'Change made: (ejo, 3/1/96)
			'==========================
			'was: Format_It(Output_ECM(i).Carbon_Usage_Rate, 2)
			'is now: Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2)

		Next i
		Printer.Print()
		Printer.Print("TC (Treatment Capacity) is in m" & Chr(179) & "  / kg of GAC")
		Printer.Print()

		For i = 1 To Number_Component_ECM
			Printer.FontBold = True
			Printer.Print(Mid(Trim(Component(IndexW(i)).Name), 1, 25))
			Printer.FontBold = False
			Printer.Print("Zone ", TAB(9), "C/Co", TAB(19), "C (mg/L)", TAB(29), "Q (mg/L)")
			For J = 1 To Number_Component_ECM
				'UPGRADE_WARNING: Couldn't resolve default property of object Solid_ConcW(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Liquid_ConcW(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Printer.Print("Zone " & VB6.Format(J, "0"), TAB(9), Format_It(CoCW(i, J), 2), TAB(19), Format_It(Liquid_ConcW(i, J) / 1000.0#, 2), TAB(29), Format_It(Solid_ConcW(i, J) / 1000.0#, 2))
			Next J
			Printer.Print()
		Next i

		Printer.Print()
		'--- Print the mass balance results
		Printer.FontSize = 12
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print("Mass Balance Results")
		Printer.FontUnderline = False
		Printer.FontSize = 10
		Printer.FontBold = False

		Printer.Print()
		'Printer.Print "Component"; Tab(30); "Left-Hand"; Tab(45); "Right-Hand"; Tab(60); "Percent Err."
		'Printer.Print ""; Tab(30); "(ug/cm2/s)"; Tab(45); "(ug/cm2/s)"; Tab(60); "(%)"
		Printer.Print("Component", TAB(30), "Percent Err.")
		Printer.Print("", TAB(30), "(%)")
		For i = 1 To Number_Component_ECM
			'      Printer.Print Mid$(Trim$(Component(i).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(i), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(i), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(i), 3)
			'      Printer.Print Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(IndexW(i)), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(IndexW(i)), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3)
			Printer.Print(Mid(Trim(Component(IndexW(i)).Name), 1, 25), TAB(30), Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3))
		Next i
		Printer.Print()


		Printer.EndDoc()
		Exit Sub
Print_Error:
		Call Show_Trapped_Error("cmdPrint_Click")
		Resume Exit_Print
Exit_Print:

	End Sub

	Private Sub Command4_Click(sender As Object, e As EventArgs)

	End Sub

	Private Sub frmModelECMResults_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class