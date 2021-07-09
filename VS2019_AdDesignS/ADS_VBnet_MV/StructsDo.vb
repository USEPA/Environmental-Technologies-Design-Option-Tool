Option Strict Off
Option Explicit On
Module StructsDo
	
	
	'INPUTS:
	'    CompNum =
	'        0 OR ABOVE : COMPONENT NUMBER TO USE FOR CALLS TO
	'                     Dp(), Ds(), Kf()
	'        -1 : DO NOT MAKE CALLS TO Dp(), Ds(), Kf() [SEE BELOW]
	'OUTPUTS:
	'    x = COMPONENT STRUCTURE TO SET DEFAULTS FOR.
	Sub SetComponentDefaults(ByRef X As ComponentPropertyType, ByRef CompNum As Short)
		Dim i As Short
		
		'***** Properties: *****
		X.Name = "New Component"
		'---- Changed by Eric J. Oman 8/8/97 BEGINS:
		X.CAS = 0
		'x.Cas = 79016
		'---- Changed by Eric J. Oman 8/8/97 ENDS.
		X.MW = 131.39
		X.MolarVolume = 102#
		X.BP = 87
		X.InitialConcentration = 50#
		X.Use_K = 98#
		X.Use_OneOverN = 0.43
		X.Source_KandOneOverN = KNSOURCE_USERINPUT
		X.UserEntered_K = 98#
		X.UserEntered_OneOverN = 0.43
		X.Treatment_Objective = 0.005
		
		'***** IPES *****
		X.IPES_OrderOfMagnitude = 4#
		X.IPES_NumRegressionPts = 50
		X.IPES_RelativeHumidity = 0#
		X.IPES_EstimationMethod = IPESMETHOD_LIQ_3PARAM
		X.Liquid_Density = 1.53
		X.Aqueous_Solubility = 1100#
		X.Vapor_Pressure = 9830#
		X.Refractive_Index = 1.475
		X.IPESResult_K = -1
		X.IPESResult_OneOverN = -1
		
		'***** Isotherm Database *****
		X.IsothermDB_Component_Name = ""
		X.IsothermDB_Range_Num = -1
		X.IsothermDB_K = -1
		X.IsothermDB_OneOverN = -1
		
		'***** Kinetics *****
		X.SPDFR = 5#
		X.SPDFR_Low_Concentration = True
		X.Use_SPDFR_Correlation = False
		X.Tortuosity = 1#
		X.Use_Tortuosity_Correlation = False
		X.Constant_Tortuosity = True
		If (IsNothing(X.Corr) Or IsNothing(X.KP_User_Input)) Then   'Shang added to avoid memory violation
			X.Initialize()
		End If

		For i = 1 To 3
			X.Corr(i) = True
		Next i
		If (CompNum <> -1) Then
			X.Dp = Dp(CompNum)
			X.Ds = Ds(CompNum)
			X.kf = kf(CompNum)
		Else
			X.Dp = 0.0000000001
			X.Ds = 0.0000000001
			X.kf = 0.0000000001
		End If
		X.KP_User_Input(1) = X.kf
		X.KP_User_Input(2) = X.Ds
		X.KP_User_Input(3) = X.Dp
		
		'***** ECM *****
		X.K_Reduction = False
		'Note: leaving x.Correlation (K reduction) unset.
		
		X.Is_Selected_On_List = False
		
	End Sub
	Sub SetBedDefaults(ByRef BedPhase As Short)
		'Liquid Phase
		If (BedPhase = 0) Then
			Bed.Phase = 0
			Bed.length = 2.765
			Bed.Diameter = 3.048
			Bed.Weight = 9072#
			Bed.Flowrate = 0.03577
			Bed.NumberOfBeds = 1
			Bed.WaterDensity = 0.99915
			Bed.Temperature = 15#
			Bed.WaterViscosity = 0.0115
			Bed.Pressure = 1#
		End If
		'Gas Phase
		If (BedPhase = 1) Then
			Bed.Phase = 1
			Bed.length = 1.09
			Bed.Diameter = 3.66
			Bed.Weight = 4540#
			Bed.Flowrate = 1.09
			Bed.NumberOfBeds = 1
			Bed.WaterDensity = 0.99569
			Bed.Temperature = 30#
			Bed.WaterViscosity = 0.00815
			Bed.Pressure = 1#
		End If
	End Sub
	Sub SetCarbonDefaults(ByRef BedPhase As Short)
		If (BedPhase = 0) Then
			Carbon.Name = "Calgon F 400"
			Carbon.Density = 0.803 'g/cm^3
			Carbon.ParticleRadius = 0.0513 / 100# 'cm ---> m
			Carbon.Porosity = 0.641 '(-)
			Carbon.ShapeFactor = 1#
			Carbon.Tortuosity = 1# 'UNUSED!
			Carbon.W0 = 0.63
			Carbon.BB = 0.02766
			Carbon.PolanyiExponent = 1.208
		End If
		If (BedPhase = 1) Then
			Carbon.Name = "Calgon BPL 4x6 mesh"
			Carbon.Density = 0.85
			Carbon.ParticleRadius = 0.186 / 100# 'cm ---> m
			Carbon.Porosity = 0.595
			Carbon.ShapeFactor = 1#
			Carbon.Tortuosity = 1# 'UNUSED!
			Carbon.W0 = 0.46
			Carbon.BB = 0.0000000337
			Carbon.PolanyiExponent = 2#
		End If
	End Sub
	
	
	'BedPhase:
	'    0 = LIQUID PHASE.
	'    1 = GAS PHASE.
	Sub Initialize_All_Data(ByRef BedPhase As Short)
		Dim i As Short
		
		Number_Component = 0
		
		Bed.Phase = BedPhase
		Call chem_phase(BedPhase)
		
		Filename = ""
		frmMain.Text = AppName_For_Display_Short & "  -  (Untitled)"
		FileNote = ""


		'Set Default PFPSDM Simulation Data:
		Number_Influent_Points = 0
		MC = 8
		NC = 3
		TimeP.Init = 19000#
		TimeP.End_Renamed = 250000#
		TimeP.np = 1900
		TimeP.Step_Renamed = 600#
		
		'Set Default Carbon Data:
		Call SetCarbonDefaults(BedPhase) 'Set Carbon defaults for Liquid Phase (0).
		
		'Set Default Bed Data:
		Call SetBedDefaults(BedPhase) 'Set Bed defaults for Liquid Phase (0).
		Bed.Initialize()   'shang added otherwise memory violation
		'Set Default Water Correlation Data:
		Bed.Water_Correlation.Name = "Organic Free Water"
		Bed.Water_Correlation.Coeff(1) = 1#
		Bed.Water_Correlation.Coeff(2) = 0#
		Bed.Water_Correlation.Coeff(3) = 0#
		Bed.Water_Correlation.Coeff(4) = 0#
		
		''Display Carbon and Bed Values on Window:
		'frmMain.txtCarbon(0) = Trim$(Carbon.name)
		'frmMain.txtCarbon(1) = Format$(Carbon.Density, "0.000")
		'frmMain.txtCarbon(2) = Format$(Carbon.ParticleRadius * 100#, "0.00000")
		'frmMain.txtCarbon(3) = Format$(Carbon.Porosity, "0.000")
		'frmMain.txtCarbon(4) = Format$(Carbon.ShapeFactor, "0.000")
		'frmMain.txtBedValue(0) = Format_It(Bed.Length, 3)
		'frmMain.txtBedValue(1) = Format_It(Bed.Diameter, 3)
		'frmMain.txtBedValue(2) = Format_It(Bed.Weight, 2)
		'frmMain.txtBedValue(3) = Format$(Bed.Flowrate, "0.000E+00")
		
		'Initialization for Kinetic Parameters
		Use_Tortuosity_Correlation = False 'UNUSED!
		Constant_Tortuosity = True 'UNUSED!
		
		'Initialization for Water Properties
		State_Check_Water(1) = 1
		State_Check_Water(2) = 1

		'Set Default Units:
		frmMain.txtBedUnits(0).SelectedIndex = 0
		frmMain.txtBedUnits(1).SelectedIndex = 0
		frmMain.txtBedUnits(2).SelectedIndex = 0
		frmMain.txtBedUnits(3).SelectedIndex = 0
		frmMain.txtBedUnits(4).SelectedIndex = 0
		frmMain.txtCarbonUnits(1).SelectedIndex = 0
		frmMain.txtCarbonUnits(2).SelectedIndex = 0
		PropertyUnits.MW = "mg/mmol"
		PropertyUnits.MolarVolume = "mL/gmol"
		PropertyUnits.BP = "C"
		PropertyUnits.InitialConcentration = "mg/L"
		PropertyUnits.Liquid_Density = "g/mL"
		PropertyUnits.Aqueous_Solubility = "mg/L"
		PropertyUnits.Vapor_Pressure = "Pa"
		PropertyUnits.k = "(mg/g)*(L/mg)^(1/n)"
		
		'Temporary kludge until we allow saveable-settable units:
		Call unitsys_set_units(frmMain.txtTime(0), "d")
		Call unitsys_set_units(frmMain.txtTime(1), "d")
		Call unitsys_set_units(frmMain.txtTime(2), "d")
		
		'REFRESH THE MAIN WINDOW.
		Call frmMain_Refresh()

		'OLD STUFF:
		''Call Update_Display       'This routine calculates EBCT from txtBedValue(0...3)
		'''''''frmmain.txtBedValue(4) = Format_It(21.03, 2)

		'	'''Update the display in the main window
		'		''Call Update_Display_Data
		'		''Call Update_Display_Kinetic
		'		''Call Update_Bed_Density_Display
		'		'OLD STUFF [ENDS].

		'SET PSDMR MODEL PARAMETERS.
		RoomParams.COUNT_CONTAMINANT = 0
		RoomParams.ROOM_VOL = 40776259# / 100# / 100# / 100.0#
		RoomParams.ROOM_FLOWRATE = 3397.89 / 100.0# / 100.0# / 100.0#
		RoomParams.Initialize()    'shang added otherwise memory violation
		For i = 1 To Number_Compo_Max
			RoomParams.ROOM_C0(i) = 0#
			RoomParams.ROOM_EMIT(i) = 1.7
			RoomParams.INITIAL_ROOM_CONC(i) = 0#
			RoomParams.RXN_RATE_CONSTANT(i) = 0#
			RoomParams.RXN_PRODUCT(i) = 0
			RoomParams.RXN_RATIO(i) = 0#
		Next i
		RoomParams.ROOM_VOL_Units = "m" & Chr(179)
		RoomParams.ROOM_FLOWRATE_Units = "m" & Chr(179) & "/s"
		RoomParams.ROOM_C0_Units = "mg/L"
		RoomParams.ROOM_EMIT_Units = "µg/s"
		RoomParams.INITIAL_ROOM_CONC_Units = "mg/L"
		Call RoomParam_Recalculate(RoomParams)
		'---- NEW AS OF 11/11/99 BEGINS: ----
		With RoomParams
			'
			'/////////   TIME-VARIABLE Co   //////////////////////////////////
			'UPGRADE_WARNING: Lower bound of array .bool_ROOM_COINI_ISTIMEVAR was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .bool_ROOM_COINI_ISTIMEVAR(Number_Compo_Max)
			For i = 1 To Number_Compo_Max
				.bool_ROOM_COINI_ISTIMEVAR(i) = False
			Next i
			'UPGRADE_WARNING: Lower bound of array .int_ROOM_NCOINI was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .int_ROOM_NCOINI(Number_Compo_Max)
			For i = 1 To Number_Compo_Max
				.int_ROOM_NCOINI(i) = 0
			Next i
			'UPGRADE_WARNING: Lower bound of array .dbl_ROOM_TCOINI was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .dbl_ROOM_TCOINI(Number_Compo_Max, Max_int_ROOM_NCOINI)
			'UPGRADE_WARNING: Lower bound of array .dbl_ROOM_COINI was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .dbl_ROOM_COINI(Number_Compo_Max, Max_int_ROOM_NCOINI)
			.u_ROOM_TCOINI = "min"
			.u_ROOM_COINI = "µg/L"
			'
			'/////////   TIME-VARIABLE w*A   /////////////////////////////////
			'UPGRADE_WARNING: Lower bound of array .bool_ROOM_EMITINI_ISTIMEVAR was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .bool_ROOM_EMITINI_ISTIMEVAR(Number_Compo_Max)
			For i = 1 To Number_Compo_Max
				.bool_ROOM_EMITINI_ISTIMEVAR(i) = False
			Next i
			'UPGRADE_WARNING: Lower bound of array .int_ROOM_NEMITINI was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .int_ROOM_NEMITINI(Number_Compo_Max)
			For i = 1 To Number_Compo_Max
				.int_ROOM_NEMITINI(i) = 0
			Next i
			'UPGRADE_WARNING: Lower bound of array .dbl_ROOM_TEMITINI was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .dbl_ROOM_TEMITINI(Number_Compo_Max, Max_int_ROOM_NEMITINI)
			'UPGRADE_WARNING: Lower bound of array .dbl_ROOM_EMITINI was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .dbl_ROOM_EMITINI(Number_Compo_Max, Max_int_ROOM_NEMITINI)
			.u_ROOM_TEMITINI = "min"
			.u_ROOM_EMITINI = "µg/s" 'MASS_EMISSION_RATE
		End With
		'---- NEW AS OF 11/11/99 ENDS. ----
		'---- NEW AS OF 1/17/00 BEGINS: ----
		With RoomParams
			'
			'/////////   TIME-VARIABLE K   /////////////////////////////////
			'UPGRADE_WARNING: Lower bound of array .bool_ROOM_KINI_ISTIMEVAR was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .bool_ROOM_KINI_ISTIMEVAR(Number_Compo_Max)
			For i = 1 To Number_Compo_Max
				.bool_ROOM_KINI_ISTIMEVAR(i) = False
			Next i
			'UPGRADE_WARNING: Lower bound of array .int_ROOM_NKINI was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .int_ROOM_NKINI(Number_Compo_Max)
			For i = 1 To Number_Compo_Max
				.int_ROOM_NKINI(i) = 0
			Next i
			'UPGRADE_WARNING: Lower bound of array .dbl_ROOM_TKINI was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .dbl_ROOM_TKINI(Number_Compo_Max, Max_int_ROOM_NKINI)
			'UPGRADE_WARNING: Lower bound of array .dbl_ROOM_KINI was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim .dbl_ROOM_KINI(Number_Compo_Max, Max_int_ROOM_NKINI)
			.u_ROOM_TKINI = "min"
			.u_ROOM_KINI = "(mg/g)*(L/mg)^(1/n)" 'freundlich_k
		End With
		'---- NEW AS OF 1/17/00 ENDS. ----
		
	End Sub
	
	
	Sub GetMoreBedParameters()
		Dim EBCT As Double
		Bed.Area = PI * Bed.Diameter * Bed.Diameter / 4
		Bed.Volume = Bed.Area * Bed.length
		Bed.Density = Bed.Weight / Bed.Volume / 1000
		Bed.Porosity = 1# - Bed.Density / Carbon.Density
		Bed.SuperficialVelocity = Bed.Flowrate / Bed.Area
		Bed.InterstitialVelocity = Bed.SuperficialVelocity / Bed.Porosity
		EBCT = Bed.Volume / Bed.Flowrate / 60
		Bed.TAU = EBCT * Bed.Porosity
	End Sub
	Sub Update_Bed_Density_Display()
		Call GetMoreBedParameters()
		frmMain.lblBedDensityDisplay.Text = VB6.Format(Bed.Density, "0.0000")
	End Sub
	
	
	'change_type:
	'    1 = RE-CALCULATE AND DISPLAY BED POROSITY.
	'    2 = RE-CALCULATE AND DISPLAY SUPERFICIAL VELOCITY AND INTERSTITIAL VELOCITY.
	'    3 = RE-CALCULATE AND DISPLAY (ALL OF THE ABOVE).
	Sub Update_Several_Bed_Properties(ByRef change_type As Short)
		Select Case change_type
			Case 1
				Bed.Porosity = 1# - Bed.Density / Carbon.Density
				Bed.InterstitialVelocity = Bed.SuperficialVelocity / Bed.Porosity
				'display superficial velocity and interstitial velocity in units of m/hr
				frmMain.lblporosity.Text = Format_It(Bed.Porosity, 3)
				frmMain.lblInterstitialVelocity.Text = Format_It(Bed.InterstitialVelocity * 3600#, 3)
			Case 2
				Bed.SuperficialVelocity = Bed.Flowrate / Bed.Area
				Bed.InterstitialVelocity = Bed.SuperficialVelocity / Bed.Porosity
				'display superficial velocity and interstitial velocity in units of m/hr
				frmMain.lblSuperficialVelocity.Text = Format_It(Bed.SuperficialVelocity * 3600#, 3)
				frmMain.lblInterstitialVelocity.Text = Format_It(Bed.InterstitialVelocity * 3600#, 3)
			Case 3
				Bed.Porosity = 1# - Bed.Density / Carbon.Density
				Bed.SuperficialVelocity = Bed.Flowrate / Bed.Area
				Bed.InterstitialVelocity = Bed.SuperficialVelocity / Bed.Porosity
				'display superficial velocity and interstitial velocity in units of m/hr
				frmMain.lblporosity.Text = Format_It(Bed.Porosity, 3)
				frmMain.lblSuperficialVelocity.Text = Format_It(Bed.SuperficialVelocity * 3600#, 3)
				frmMain.lblInterstitialVelocity.Text = Format_It(Bed.InterstitialVelocity * 3600#, 3)
		End Select
	End Sub
	
	
	Sub chem_phase(ByRef Index As Short)
		'Dim response As Integer
		'Dim temp As String
		'Dim p1 As String, p2 As String
		''DO NOT RUN THIS CODE IF THE PHASE HAS ALREADY BEEN CHANGED
		''(AS INDICATED BY THE CHECKMARK NEXT TO THE PHASE LABEL
		''ON THE MENU).
		'If (frmMain.mnuPhaseItem(Index).Checked) Then Exit Sub
		'If (Not Flag_Openfile) Then
		'  'If loading in a data file, don't ask the user
		'  'about changing phases.
		'Select Case index
		'  Case 0
		'    p1 = "gas"
		'    p2 = "liquid"
		'  Case 1
		'    p1 = "liquid"
		'    p2 = "gas"
		'End Select
		'  'temp = "Changing phase from " & p1 & " to " & p2 & " will destroy all current data."
		'  'temp = temp & NL & "Change phase anyway?"
		'  'response = MsgBox(temp, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, AppName_For_Display_long)
		'  'If (response <> IDYES) Then Exit Sub
		'End If
		
		Select Case Index
			Case 0
				frmMain.mnuPhaseItem(0).Checked = True
				frmMain.mnuPhaseItem(1).Checked = False
				Bed.Phase = 0
				'If (Not Flag_Openfile) Then Call Initialize_All_Data(0)
				'frmMain.cmdEditWater.Caption = "Edi&t Water Properties"
				'Call H2ODens(Bed.WaterDensity, Bed.Temperature + 273.15)
				'Call H2OVisc(Bed.WaterViscosity, Bed.Temperature + 273.15)
				'Bed.WaterDensity = Bed.WaterDensity / 1000#
				'Bed.WaterViscosity = Bed.WaterViscosity * 10#
				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.ssframe_Water.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmMain.watergroup.Text = "Water Properties:"
			Case 1
				frmMain.mnuPhaseItem(0).Checked = False
				frmMain.mnuPhaseItem(1).Checked = True
				Bed.Phase = 1
				'If (Not Flag_Openfile) Then Call Initialize_All_Data(1)
				'frmMain.cmdEditWater.Caption = "Edi&t Air Properties"
				'Call AIRDens(Bed.WaterDensity, Bed.Temperature + 273.15, Bed.Pressure)
				'Call AirVisc(Bed.WaterViscosity, Bed.Temperature + 273.15)
				'Bed.WaterDensity = Bed.WaterDensity / 1000#
				'Bed.WaterViscosity = Bed.WaterViscosity * 10#
				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.ssframe_Water.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmMain.watergroup.Text = "Air Properties:"
		End Select
		Call Update_FluidDensity(Bed.Temperature, Bed.Pressure, Bed.WaterDensity)
		Call Update_FluidViscosity(Bed.Temperature, Bed.WaterViscosity)
		Call Update_KP_Values()
		'Call Update_Display_Kinetic
		'Call Update_Bed_Density_Display
	End Sub
End Module