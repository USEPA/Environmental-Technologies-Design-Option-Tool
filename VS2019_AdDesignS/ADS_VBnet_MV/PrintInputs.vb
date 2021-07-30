Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmPrintInputs
	Inherits System.Windows.Forms.Form
	Dim rs As New Resizer

	Dim Filename_Input As String





	Const frmPrintInputs_declarations_end As Boolean = True
	
	
	Private Sub chkSelect_Click(ByRef index As Short, ByRef Value As Short)
		'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(index).Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		chkSelect(index).Tag = CStr(Str(Value))
	End Sub
	
	
	Private Sub cmdCancel_Click()
		Me.Close()
	End Sub
	
	
	Private Sub cmdPrint_Click()
		Dim cdlCancel As Object
		Dim cdlOFNPathMustExist As Object
		Dim cdlOFNOverwritePrompt As Object
		Dim Printer As New Printer
		Dim Error_Code, f As Short
		Dim temp As String
		Dim i, DFlag As Short
		Dim Dummy As Double
		Dim Eq1, temporaryname As String
		Dim response As Short
		Dim s As String
		Dim J As Short
		
		DFlag = False
		For i = 0 To 4
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(i).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(i).Value Then DFlag = True
		Next i
		If Not (DFlag) Then
			Call Show_Error("You must select something to print!")
			Exit Sub
		End If
		
		If Print_To_Printer Then
			On Error GoTo Print_Error
			Printer.ScaleLeft = -1080 'Set a 3/4-inch margin
			Printer.ScaleTop = -1080
			Printer.CurrentX = 0
			Printer.CurrentY = 0
			Printer.FontSize = 10
			Printer.Print(Filename)
			Printer.Print()
			'---- Print Component Properties  ---------
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(0).Value Then
				Printer.FontSize = 12
				Printer.FontBold = True
				Printer.Print(TAB(25), "Component Properties")
				Printer.FontSize = 10
				Printer.FontBold = False
				Printer.Print()
				Printer.Print("Component", TAB(30), "K*", TAB(38), "1/n", TAB(47), "C0", TAB(57), "MW", TAB(65), "Vm", TAB(75), "NBP")
				Printer.Print(TAB(39), "-", TAB(46), "mg/L", TAB(56), "g/mol", TAB(65), "cm" & Chr(179) & "/mol", TAB(76), "C")
				For i = 1 To Number_Component
					Printer.Print(Trim(Mid(LTrim(Component(i).Name), 1, 25)), TAB(29), VB6.Format(Component(i).Use_K, "###,##0.000"), TAB(37), VB6.Format(Component(i).Use_OneOverN, "0.000"), TAB(46), Format_It(Component(i).InitialConcentration, 2), TAB(55), Format_It(Component(i).MW, 2), TAB(64), Format_It(Component(i).MolarVolume, 2), TAB(73), Format_It(Component(i).BP, 2))
				Next i
				Printer.Print()
				Printer.Print("* K in (mg/g)*(L/mg)^(1/n) - Vm = Molar Volume at NBP")
				Printer.Print()
			End If
			'---Print bed data-----
			Call GetMoreBedParameters()
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(1).Value Then
				Printer.FontSize = 12
				Printer.FontBold = True
				Printer.Print(TAB(25), "Fixed-Bed Properties")
				Printer.FontSize = 10
				Printer.FontBold = False
				Printer.Print()
				Printer.Print("Bed Length:", TAB(25), VB6.Format(Bed.length, "0.000E+00") & " m")
				Printer.Print("Bed Diameter:", TAB(25), VB6.Format(Bed.Diameter, "0.000E+00") & " m")
				Printer.Print("Weight of GAC:", TAB(25), VB6.Format(Bed.Weight, "0.000E+00") & " kg")
				Printer.Print("Inlet Flowrate:", TAB(25), VB6.Format(Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
				Printer.Print("EBCT:", TAB(25), VB6.Format(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#, "0.000E+00") & " mn")
				Printer.Print("Bed Density:", TAB(25), VB6.Format(Bed.Density, "0.000") & " g/cm3")
				Printer.Print("Bed Porosity:", TAB(25), VB6.Format(Bed.Porosity, "0.000"))
				Printer.Print("Superficial Velocity:", TAB(25), VB6.Format(Bed.SuperficialVelocity * 3600#, "0.00E+00") & " m/hr")
				Printer.Print("Interstitial Velocity:", TAB(25), VB6.Format(Bed.InterstitialVelocity * 3600#, "0.00E+00") & " m/hr")
				Printer.Print()
				Printer.Print("Temperature:", TAB(25), VB6.Format(Bed.Temperature, "0.00") & " C")
				Printer.Print("Water Density:", TAB(25), VB6.Format(Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
				Printer.Print("Water Viscosity:", TAB(25), VB6.Format(Bed.WaterViscosity, "0.00E+0") & " g/cm.s")
				Printer.Print()
			End If
			'--- Print Carbon Properties -----
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(2).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(2).Value Then
				Printer.FontSize = 12
				Printer.FontBold = True
				Printer.Print(TAB(25), "Carbon Properties")
				Printer.FontSize = 10
				Printer.FontBold = False
				Printer.Print()
				Printer.Print("Name:", TAB(19), Trim(Carbon.Name))
				Printer.Print("Apparent Density:", TAB(19), VB6.Format(Carbon.Density, "0.000") & " g/cm" & Chr(179))
				Printer.Print("Particle Radius:", TAB(19), VB6.Format(Carbon.ParticleRadius * 100#, "0.00000") & " cm")
				Printer.Print("Porosity:", TAB(19), VB6.Format(Carbon.Porosity, "0.000"))
				Printer.Print("Shape Factor: ", TAB(19), VB6.Format(Carbon.ShapeFactor, "0.000"))
				'Printer.Print "Tortuosity:"; Tab(19); Format$(Carbon.Tortuosity, "0.000")
				Printer.Print()
			End If
			'--- Print kinetic Parameters -----
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(3).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(3).Value Then
				Printer.FontSize = 12
				Printer.FontBold = True
				Printer.Print(TAB(25), "Kinetic Parameters")
				Printer.FontSize = 10
				Printer.FontBold = False
				Printer.Print()
				Printer.Print("Component", TAB(24), "kf", TAB(33), "Ds", TAB(42), "Dp", TAB(50), "St", TAB(58), "Eds", TAB(67), "Edp", TAB(75), "SPDFR")
				Printer.Print(TAB(23), "cm/s", TAB(32), "cm" & Chr(178) & "/s", TAB(41), "cm" & Chr(178) & "/s", TAB(50), "-", TAB(59), "-", TAB(68), "-", TAB(77), "-")
				For i = 1 To Number_Component
					'UPGRADE_WARNING: Couldn't resolve default property of object Edp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Eds(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Printer.Print(Mid(Trim(Component(i).Name), 1, 20), TAB(22), Format_It(Component(i).kf, 2), TAB(31), VB6.Format(Component(i).Ds, "0.00E+00"), TAB(40), VB6.Format(Component(i).Dp, "0.00E+00"), TAB(49), Format_It(ST(i), 2), TAB(58), Format_It(Eds(i), 2), TAB(67), Format_It(Edp(i), 2), TAB(76), Format_It(Component(i).SPDFR, 2))
				Next i
			End If
			'--- Print fouling correlations ---
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(4).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(4).Value Then
				J = False
				
				For i = 1 To Number_Component
					'if and only if using correlation then print correlation
					If Component(i).Use_Tortuosity_Correlation = True Then J = True
				Next i
				
				If J Then
					Printer.Print()
					Printer.FontSize = 12
					Printer.FontBold = True
					Printer.Print(TAB(25), "Fouling Correlations")
					Printer.FontSize = 10
					Printer.FontBold = False
					Printer.Print()
					
					Printer.Print(" Water type : " & Trim(Bed.Water_Correlation.Name))
					Eq1 = VB6.Format(Bed.Water_Correlation.Coeff(1), "0.00")
					
					If Bed.Water_Correlation.Coeff(2) > 0 Then
						Eq1 = Eq1 & " + " & VB6.Format(Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
					Else
						If Bed.Water_Correlation.Coeff(2) < 0 Then
							Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
						End If
					End If
					If Bed.Water_Correlation.Coeff(3) > 0 Then
						Eq1 = Eq1 & " + " & VB6.Format(Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
					Else
						If Bed.Water_Correlation.Coeff(3) < 0 Then
							Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
						End If
					End If
					If Bed.Water_Correlation.Coeff(3) <> 0 Then
						If Bed.Water_Correlation.Coeff(4) > 0 Then
							Eq1 = Eq1 & VB6.Format(Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
						Else
							If Bed.Water_Correlation.Coeff(4) < 0 Then
								Eq1 = Eq1 & VB6.Format(Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
							End If
						End If
					End If
					Printer.Print("K(t)/K0 = " & Eq1)
					Printer.Print("(t in minutes)")
					Printer.Print()
					
					For i = 1 To Number_Component
						If Component(i).Use_Tortuosity_Correlation = True Then
							Eq1 = ""
							If Component(i).Correlation.Coeff(1) = 1# Then
								Eq1 = "(K/K0) "
							Else
								If Component(i).Correlation.Coeff(1) <> 0 Then Eq1 = VB6.Format(Component(i).Correlation.Coeff(1), "0.00") & " * (K/K0) "
							End If
							If Component(i).Correlation.Coeff(2) > 0 Then
								Eq1 = Eq1 & "+ " & VB6.Format(Component(i).Correlation.Coeff(2), "0.00")
							Else
								If Component(i).Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & VB6.Format(System.Math.Abs(Component(i).Correlation.Coeff(2)), "0.00")
							End If
							If Trim(Eq1) = "" Then
								Eq1 = "K/K0"
							End If
							Printer.Print(Trim(Component(i).Name) & ":")
							Printer.Print(TAB(10), "Correlation type: " & Trim(Component(i).Correlation.Name))
							
							Printer.Print(TAB(10), "K/K0 = " & Eq1)
							If (Component(i).Use_Tortuosity_Correlation) Then
								If (Component(i).Constant_Tortuosity) Then
									Printer.Print("Correlation used when SOC competition is important:")
									Printer.Print(" Tortuosity = 0.782 * EBCT^0.925 ")
								Else
									Printer.Print("Correlation used when NOM fouling is important:")
									Printer.Print(" Tortuosity = 1.0 if t< 70 days")
									Printer.Print(" Tortuosity = 0.334 + 6.610E-06 * t   (t in minutes)")
								End If
							End If
							
							Printer.Print()
						End If
					Next i
					
					Printer.Print()
				End If
			End If
			
			'--- Print Variable Influent Data
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(5).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(5).Value Then
				Printer.Print()
				Printer.Print(TAB(25), "Variable Influent Data")
				Printer.Print()
				s = "Time(days)"
				For J = 1 To Number_Component
					s = s & ",C of " & Trim(Component(J).Name)
				Next J
				s = s & ":"
				Printer.Print(s)
				Printer.Print("(All C in mg/L)")
				For i = 1 To Number_Influent_Points
					s = Trim(Str(T_Influent(i) / 60# / 24#)) 'Convert min--->days
					For J = 1 To Number_Component
						s = s & "," & Trim(Str(C_Influent(J, i)))
					Next J
					Printer.Print(s)
				Next i
			End If
			
			'--- Print Effluent Data
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(6).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(6).Value Then
				Printer.Print()
				Printer.Print(TAB(25), "Effluent Data")
				Printer.Print()
				s = "Time(days)"
				For J = 1 To Number_Component
					s = s & ",C/C0 of " & Trim(Component(J).Name)
				Next J
				s = s & ":"
				Printer.Print(s)
				Printer.Print("(All C/C0 are dimensionless and normalized)")
				For i = 1 To NData_Points
					s = Trim(Str(T_Data_Points(i)))
					For J = 1 To Number_Component
						s = s & "," & Trim(Str(C_Data_Points(J, i)))
					Next J
					Printer.Print(s)
				Next i
			End If
			
			Printer.EndDoc()
		Else
			
			On Error GoTo File_Error
			'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.CancelError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.CancelError = True
			''UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.Filename = ""
			''UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.DialogTitle = "Print to File"
			''UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
			''UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.FilterIndex = 2
			''UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			''UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNPathMustExist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			''UPGRADE_WARNING: Couldn't resolve default property of object cdlOFNOverwritePrompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
			''UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'CMDialog1.Action = 2

			''f = FileNameIsValid(Filename_Input, CMDialog1)
			''If Not (f) Then Exit Sub
			''UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'Filename_Input = CMDialog1.Filename

			f = FreeFile
			FileOpen(f, Filename_Input, OpenMode.Output)
			PrintLine(f, Filename)
			
			'---- Print Component Properties  ---------
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(0).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(0).Value Then
				PrintLine(f, TAB(25), "Component Properties")
				PrintLine(f)
				PrintLine(f, "Component", TAB(30), "K*", TAB(38), "1/n", TAB(47), "C0", TAB(57), "MW", TAB(65), "Vm", TAB(75), "NBP")
				PrintLine(f, TAB(39), "-", TAB(46), "mg/L", TAB(56), "g/mol", TAB(65), "cm" & Chr(179) & "/mol", TAB(76), "C")
				For i = 1 To Number_Component
					PrintLine(f, Trim(Mid(LTrim(Component(i).Name), 1, 25)), TAB(29), VB6.Format(Component(i).Use_K, "###,##0.000"), TAB(37), VB6.Format(Component(i).Use_OneOverN, "0.000"), TAB(46), Format_It(Component(i).InitialConcentration, 2), TAB(55), Format_It(Component(i).MW, 2), TAB(64), Format_It(Component(i).MolarVolume, 2), TAB(73), Format_It(Component(i).BP, 2))
				Next i
				PrintLine(f)
				PrintLine(f, "* K in (mg/g)*(L/mg)^(1/n) - Vm = Molar Volume at NBP")
				PrintLine(f)
			End If
			
			'---Print bed data-----
			Call GetMoreBedParameters()
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(1).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(1).Value Then
				PrintLine(f, TAB(25), "Bed Data")
				PrintLine(f)
				PrintLine(f, "Bed Length:", TAB(18), VB6.Format(Bed.length, "0.000E+00") & " m")
				PrintLine(f, "Bed Diameter:", TAB(18), VB6.Format(Bed.Diameter, "0.000E+00") & " m")
				PrintLine(f, "Weight of GAC:", TAB(18), VB6.Format(Bed.Weight, "0.000E+00") & " kg")
				PrintLine(f, "Inlet Flowrate:", TAB(18), VB6.Format(Bed.Flowrate, "0.000E+00") & " m" & Chr(179) & "/s")
				PrintLine(f, "EBCT:", TAB(18), VB6.Format(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#, "0.000E+00") & " mn")
				PrintLine(f, "Bed Density:", TAB(25), VB6.Format(Bed.Density, "0.000") & " g/cm3")
				PrintLine(f, "Bed Porosity:", TAB(25), VB6.Format(Bed.Porosity, "0.000"))
				PrintLine(f, "Superficial Velocity:", TAB(25), VB6.Format(Bed.SuperficialVelocity * 3600#, "0.00E+00") & " m/hr")
				PrintLine(f, "Interstitial Velocity:", TAB(25), VB6.Format(Bed.InterstitialVelocity * 3600#, "0.00E+00") & " m/hr")
				PrintLine(f)
				PrintLine(f, "Temperature:", TAB(18), VB6.Format(Bed.Temperature, "0.00") & " C")
				PrintLine(f, "Water Density:", TAB(18), VB6.Format(Bed.WaterDensity, "0.0000") & " g/cm" & Chr(179))
				PrintLine(f, "Water Viscosity:", TAB(18), VB6.Format(Bed.WaterViscosity, "0.00E+0") & " g/cm.s")
				PrintLine(f)
			End If
			'--- Print Carbon Properties -----
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(2).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(2).Value Then
				PrintLine(f, TAB(25), "Carbon Properties")
				PrintLine(f)
				PrintLine(f, "Name:", TAB(19), Trim(Carbon.Name))
				PrintLine(f, "Apparent Density:", TAB(19), VB6.Format(Carbon.Density, "0.000") & " g/cm" & Chr(179))
				PrintLine(f, "Particle Radius:", TAB(19), VB6.Format(Carbon.ParticleRadius * 100#, "0.00000") & " cm")
				PrintLine(f, "Porosity:", TAB(19), VB6.Format(Carbon.Porosity, "0.000"))
				PrintLine(f, "Shape Factor: ", TAB(19), VB6.Format(Carbon.ShapeFactor, "0.000"))
				'Print #f, "Tortuosity:"; Tab(19); Format$(Carbon.Tortuosity, "0.000")
				PrintLine(f)
			End If
			'--- Print kinetic Parameters -----
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(3).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(3).Value Then
				PrintLine(f, TAB(25), "Kinetic Parameters")
				PrintLine(f)
				PrintLine(f, "Component", TAB(24), "kf", TAB(33), "Ds", TAB(42), "Dp", TAB(50), "St", TAB(58), "Eds", TAB(67), "Edp", TAB(75), "SPDFR")
				PrintLine(f, TAB(23), "cm/s", TAB(32), "cm" & Chr(178) & "/s", TAB(41), "cm" & Chr(178) & "/s", TAB(50), "-", TAB(59), "-", TAB(68), "-", TAB(77), "-")
				For i = 1 To Number_Component
					'UPGRADE_WARNING: Couldn't resolve default property of object Edp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Eds(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(f, Mid(Trim(Component(i).Name), 1, 20), TAB(22), Format_It(Component(i).kf, 2), TAB(31), VB6.Format(Component(i).Ds, "0.00E+00"), TAB(40), VB6.Format(Component(i).Dp, "0.00E+00"), TAB(49), Format_It(ST(i), 2), TAB(58), Format_It(Eds(i), 2), TAB(67), Format_It(Edp(i), 2), TAB(76), Format_It(Component(i).SPDFR, 2))
				Next i
			End If
			'--- Print fouling correlations ---
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(4).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(4).Value Then
				'check to see if fouling needed
				J = False
				For i = 1 To Number_Component
					'if and only if using correlation then print correlation
					If Component(i).Use_Tortuosity_Correlation = True Then J = True
				Next i
				
				If J Then
					PrintLine(f)
					PrintLine(f, TAB(25), "Fouling Correlations")
					PrintLine(f)
					
					PrintLine(f, " Water type : " & Trim(Bed.Water_Correlation.Name))
					Eq1 = VB6.Format(Bed.Water_Correlation.Coeff(1), "0.00")
					
					If Bed.Water_Correlation.Coeff(2) > 0 Then
						Eq1 = Eq1 & " + " & VB6.Format(Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
					Else
						If Bed.Water_Correlation.Coeff(2) < 0 Then
							Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
						End If
					End If
					If Bed.Water_Correlation.Coeff(3) > 0 Then
						Eq1 = Eq1 & " + " & VB6.Format(Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
					Else
						If Bed.Water_Correlation.Coeff(3) < 0 Then
							Eq1 = Eq1 & " - " & VB6.Format(System.Math.Abs(Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
						End If
					End If
					If Bed.Water_Correlation.Coeff(3) <> 0 Then
						If Bed.Water_Correlation.Coeff(4) > 0 Then
							Eq1 = Eq1 & VB6.Format(Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
						Else
							If Bed.Water_Correlation.Coeff(4) < 0 Then
								Eq1 = Eq1 & VB6.Format(Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
							End If
						End If
					End If
					PrintLine(f, "K(t)/K0 = " & Eq1)
					PrintLine(f, "(t in minutes)")
					PrintLine(f)
					
					For i = 1 To Number_Component
						'if and only if using correlation then print correlation
						If Component(i).Use_Tortuosity_Correlation = True Then
							
							Eq1 = ""
							If Component(i).Correlation.Coeff(1) = 1# Then
								Eq1 = "(K/K0) "
							Else
								If Component(i).Correlation.Coeff(1) <> 0 Then Eq1 = VB6.Format(Component(i).Correlation.Coeff(1), "0.00") & " * (K/K0) "
							End If
							If Component(i).Correlation.Coeff(2) > 0 Then
								Eq1 = Eq1 & "+ " & VB6.Format(Component(i).Correlation.Coeff(2), "0.00")
							Else
								If Component(i).Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & VB6.Format(System.Math.Abs(Component(i).Correlation.Coeff(2)), "0.00")
							End If
							If Trim(Eq1) = "" Then
								Eq1 = "K/K0"
							End If
							PrintLine(f, Trim(Component(i).Name) & ":")
							PrintLine(f, TAB(10), "Correlation type: " & Trim(Component(i).Correlation.Name))
							
							PrintLine(f, TAB(10), "K/K0 = " & Eq1)
							
							If (Component(i).Use_Tortuosity_Correlation) Then
								If (Component(i).Constant_Tortuosity) Then
									PrintLine(f, "Correlation used when SOC competition is important:")
									PrintLine(f, " Tortuosity = 0.782 * EBCT^0.925 ")
								Else
									PrintLine(f, "Correlation used when NOM fouling is important:")
									PrintLine(f, " Tortuosity = 1.0 if t< 70 days")
									PrintLine(f, " Tortuosity = 0.334 + 6.610E-06 * t   (t in minutes)")
								End If
							End If
							PrintLine(f)
						End If
					Next i
				End If
				'If Use_Tortuosity_Correlation Then
				'  If Constant_Tortuosity Then
				'   Print #f, "Correlation used when SOC competition is important:"
				'   Print #f, " Tortuosity = 0.782 * EBCT^0.925 "
				'  Else
				'   Print #f, "Correlation used when NOM fouling is important:"
				'   Print #f, " Tortuosity = 1.0 if t< 70 days"
				'   Print #f, " Tortuosity = 0.334 + 6.610E-06 * EBCT"
				'  End If
				'End If
				PrintLine(f)
			End If
			
			'--- Print Variable Influent Data
			
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(5).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(5).Value Then
				PrintLine(f)
				PrintLine(f, TAB(25), "Variable Influent Data")
				PrintLine(f)
				s = "Time(days)"
				For J = 1 To Number_Component
					s = s & ",C of " & Trim(Component(J).Name)
				Next J
				s = s & ":"
				PrintLine(f, s)
				PrintLine(f, "(All C in mg/L)")
				For i = 1 To Number_Influent_Points
					s = Trim(Str(T_Influent(i) / 60# / 24#)) 'Convert min--->days
					For J = 1 To Number_Component
						s = s & "," & Trim(Str(C_Influent(J, i)))
					Next J
					PrintLine(f, s)
				Next i
			End If
			
			'--- Print Effluent Data
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(6).Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If chkSelect(6).Value Then
				PrintLine(f)
				PrintLine(f, TAB(25), "Effluent Data")
				PrintLine(f)
				s = "Time(days)"
				For J = 1 To Number_Component
					s = s & ",C/C0 of " & Trim(Component(J).Name)
				Next J
				s = s & ":"
				PrintLine(f, s)
				PrintLine(f, "(All C/C0 are dimensionless and normalized)")
				For i = 1 To NData_Points
					s = Trim(Str(T_Data_Points(i)))
					For J = 1 To Number_Component
						s = s & "," & Trim(Str(C_Data_Points(J, i)))
					Next J
					PrintLine(f, s)
				Next i
			End If
			
			FileClose((f))
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object CMDialog1.Filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'CMDialog1.Filename = ""
		Me.Close()
		Exit Sub
		
Print_Error: 
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = cdlCancel) Then
			'DO NOTHING.
		Else
			Call Show_Trapped_Error("cmdPrint_Click")
		End If
		Resume Exit_Print
File_Error: 
		'UPGRADE_WARNING: Couldn't resolve default property of object cdlCancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Err.Number = cdlCancel) Then
			'DO NOTHING.
		Else
			Call Show_Trapped_Error("cmdPrint_Click")
		End If
		Resume Exit_Print
Exit_Print: 
	End Sub
	
	
	Private Sub frmPrintInputs_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		rs.FindAllControls(Me)

		Dim temp As String
		Dim temp2 As String
		Dim temp3 As String
		Dim i As Short
		''''Me.HelpContextID = Hlp_Print_
		Call UserPrefs_Load()
		For i = 0 To 6
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(i).Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(i).Tag = CStr(Str(chkSelect(i).Value))
		Next i
		Call CenterOnForm(Me, frmMain)
		If Print_To_Printer Then
			Me.Text = "Print to printer"
		Else
			Me.Text = "Print to file"
		End If
		If (Number_Component <= 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			temp = chkSelect(0).Tag
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			temp2 = chkSelect(3).Tag
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			temp3 = chkSelect(4).Tag
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(0).Enabled = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(0).Value = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(3).Enabled = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(3).Value = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(4).Enabled = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(4).Value = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(0).Tag = temp
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(3).Tag = temp2
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(4).Tag = temp3
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(0).Value = True
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(3).Value = True
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(4).Value = True
		End If
		If (Number_Influent_Points = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			temp = chkSelect(5).Tag
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(5).Value = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(5).Enabled = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(5).Tag = temp
		End If
		If (NData_Points = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			temp = chkSelect(6).Tag
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(6).Value = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(6).Enabled = False
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(6).Tag = temp
		End If
	End Sub
	Private Sub frmPrintInputs_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call UserPrefs_Save()
	End Sub
	
	
	Private Sub UserPrefs_Load()
		Dim X As Integer
		Dim i As Short
		Dim varname As String
		On Error GoTo err_FRMPRINT_UserPrefs_Load
		For i = 0 To 6
			varname = "FRMPRINT_chkSelect(" & Trim(Str(i)) & ")"
			X = CInt(INI_Getsetting(varname))
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chkSelect(i).Value = X
		Next i
		Exit Sub
resume_err_FRMPRINT_UserPrefs_Load: 
		Call UserPrefs_Save()
		Exit Sub
err_FRMPRINT_UserPrefs_Load: 
		Resume resume_err_FRMPRINT_UserPrefs_Load
	End Sub
	Private Sub UserPrefs_Save()
		Dim X As Integer
		Dim i As Short
		Dim varname As String
		For i = 0 To 6
			varname = "FRMPRINT_chkSelect(" & Trim(Str(i)) & ")"
			'UPGRADE_WARNING: Couldn't resolve default property of object chkSelect().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			X = CInt(chkSelect(i).Value)
			Call INI_PutSetting(varname, Trim(CStr(X)))
		Next i
	End Sub

	Private Sub frmPrintInputs_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		rs.ResizeAllControls(Me)

	End Sub
End Class