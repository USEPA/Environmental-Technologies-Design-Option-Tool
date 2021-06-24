Option Strict Off
Option Explicit On
Module Refresh
	
	
	
	Private Sub frmMain_Repopulate_Values()
		'UPDATE NUMERIC VALUES TO WINDOW.
		
		'WATER/AIR PROPERTIES.
		Call unitsys_set_number_in_base_units(frmMain.txtWater(0), Bed.Temperature)
		Call unitsys_set_number_in_base_units(frmMain.txtWater(1), Bed.Pressure)
		'frmMain.txtWater(0) = Format$(Bed.Temperature, "0.00")
		'frmMain.txtWater(1) = Format$(Bed.Pressure, "0.000")
		''txtWater(2) = Format$(Bed.WaterDensity, "0.000E+00")
		''txtWater(3) = Format$(Bed.WaterViscosity, "0.00E+00")

		'SIM PARAMS FOR PSDM ONLY.
		'Call unitsys_set_number_in_base_units((frmMain.txtNumberOfBeds), CDbl(Bed.NumberOfBeds))
		'Call unitsys_set_number_in_base_units(frmMain.txtNPoint(0), CDbl(MC))
		'Call unitsys_set_number_in_base_units(frmMain.txtNPoint(1), CDbl(NC))
		Call unitsys_set_number_in_base_units(frmMain.txtTime(0), TimeP.End_Renamed)
		Call unitsys_set_number_in_base_units(frmMain.txtTime(1), TimeP.Init)
		Call unitsys_set_number_in_base_units(frmMain.txtTime(2), TimeP.Step_Renamed)
		'frmMain.txtNumberOfBeds = Format$(Bed.NumberOfBeds, "0")
		'frmMain.txtNPoint(0) = Format$(MC, "0")
		'frmMain.txtNPoint(1) = Format$(NC, "0")

		'set components to save
		frmMain.NumericUpDown1.Value = Format$(Bed.NumberOfBeds, "0")
		frmMain.NumericUpDown2.Value = Format$(MC, "0")
		frmMain.NumericUpDown3.Value = Format$(NC, "0")

		'frmMain.txtTime(0) = Format_It(TimeP.End / 60# / 24#, 2)
		'frmMain.txtTime(1) = Format_It(TimeP.Init / 60# / 24#, 2)
		'frmMain.txtTime(2) = Format_It(TimeP.Step / 60# / 24#, 2)

		'FIXED BED PROPERTIES.
		Call unitsys_set_number_in_base_units(frmMain.txtBedValue(0), Bed.length)
		Call unitsys_set_number_in_base_units(frmMain.txtBedValue(1), Bed.Diameter)
		Call unitsys_set_number_in_base_units(frmMain.txtBedValue(2), Bed.Weight)
		'ConversionFactor = LengthConversionFactor(CInt(frmMain.txtBedUnits(0).ListIndex))
		'frmMain.txtBedValue(0) = Format_It(Bed.Length * ConversionFactor, 3)
		'ConversionFactor = LengthConversionFactor(CInt(frmMain.txtBedUnits(1).ListIndex))
		'frmMain.txtBedValue(1) = Format_It(Bed.Diameter * ConversionFactor, 3)
		'ConversionFactor = MassConversionFactor(CInt(frmMain.txtBedUnits(2).ListIndex))
		'frmMain.txtBedValue(2) = Format_It(Bed.Weight * ConversionFactor, 2)
		''** Note: Update_Display() takes care of Flowrate and EBCT.
		
		'ADSORBENT PROPERTIES.
		Call AssignTextAndTag(frmMain.txtCarbon(0), Trim(Carbon.Name))
		Call unitsys_set_number_in_base_units(frmMain.txtCarbon(1), Carbon.Density)
		Call unitsys_set_number_in_base_units(frmMain.txtCarbon(2), Carbon.ParticleRadius)
		Call unitsys_set_number_in_base_units(frmMain.txtCarbon(3), Carbon.Porosity)
		Call unitsys_set_number_in_base_units(frmMain.txtCarbon(4), Carbon.ShapeFactor)
		'frmMain.txtCarbon(0) = Carbon.name
		'ConversionFactor = DensityConversionFactor(CInt(frmMain.txtCarbonUnits(1).ListIndex))
		'frmMain.txtCarbon(1) = Format$(Carbon.Density * ConversionFactor, "0.000")
		'ConversionFactor = LengthConversionFactor(CInt(frmMain.txtCarbonUnits(2).ListIndex))
		'frmMain.txtCarbon(2) = Format$(Carbon.ParticleRadius * ConversionFactor, "0.00000")
		'frmMain.txtCarbon(3) = Format$(Carbon.Porosity, "0.000")
		'frmMain.txtCarbon(4) = Format$(Carbon.ShapeFactor, "0.000")
	End Sub
	Sub frmMain_Refresh()
		Dim i As Short
		Dim ConversionFactor As Double
		Dim dd As Double
		Dim T As Double
		
		Dim Enabled_Add As Boolean
		Dim Enabled_Delete As Boolean
		Dim Enabled_Edit As Boolean
		'Dim Enabled_PSDM_Results As Boolean
		'Dim Enabled_CPHSDM_Results As Boolean
		'Dim Enabled_ECM_Results As Boolean
		'Dim Enabled_PSDM_Comparison As Boolean
		'Dim Enabled_CPHSDM_Comparison As Boolean
		Dim Enabled_OptionsMenu As Boolean
		Dim Enabled_RunMenu As Boolean
		Dim Enabled_Save As Boolean
		Dim Enabled_ViewDimless As Boolean
		Dim Is_At_Least_One_Component As Boolean
		Dim SAVE_OLD_POSITION As Short
		
		'/////////// FORMERLY NAMED Update_Display_Data() ////////////////////////////////
		'UPDATE COMPONENT SELECTION LIST AND SCROLLBOX.
		If (frmMain.cboSelectCompo.Items.Count >= 1) And (frmMain.cboSelectCompo.SelectedIndex >= 0) Then
			SAVE_OLD_POSITION = frmMain.cboSelectCompo.SelectedIndex
		Else
			SAVE_OLD_POSITION = -1
		End If
		frmMain.lstComponents.Items.Clear()
		frmMain.cboSelectCompo.Items.Clear()
		For i = 1 To Number_Component
			frmMain.cboSelectCompo.Items.Add(Component(i).Name)
			frmMain.lstComponents.Items.Add(Component(i).Name)
			frmMain.lstComponents.SetSelected(i - 1, Component(i).Is_Selected_On_List)
		Next i
		If (SAVE_OLD_POSITION <> -1) And (SAVE_OLD_POSITION <= frmMain.cboSelectCompo.Items.Count - 1) Then
			frmMain.cboSelectCompo.SelectedIndex = SAVE_OLD_POSITION
		Else
			If (frmMain.cboSelectCompo.Items.Count >= 1) Then
				frmMain.cboSelectCompo.SelectedIndex = SAVE_OLD_POSITION = 0
			End If
		End If
		
		'COMPONENT SELECTION STUFF.
		If (Number_Component > 0) Then
			frmMain.cboSelectCompo.Enabled = True
			''''''''frmMain.cboSelectCompo.ListIndex = 0
		Else
			frmMain.cboSelectCompo.Enabled = False
			Component_Number_Selected = 0
		End If
		
		'ENABLE/DISABLE ADD/DELETE/EDIT.
		Enabled_Add = True 'ENABLE ADD.
		Enabled_Delete = True 'ENABLE DELETE.
		Enabled_Edit = True 'ENABLE EDIT.
		If (Number_Component = Number_Compo_Max) Then
			Enabled_Add = False 'DISABLE ADD.
		End If
		If (Number_Component = 0) Then
			Enabled_Delete = False 'DISABLE DELETE.
			Enabled_Edit = False 'DISABLE EDIT.
		End If
		'ENABLE/DISABLE OPTIONS MENU, RUN MENU, AND SAVE/SAVE-AS.
		Is_At_Least_One_Component = (Number_Component >= 1)
		Enabled_OptionsMenu = Is_At_Least_One_Component
		Enabled_RunMenu = Is_At_Least_One_Component
		Enabled_Save = Is_At_Least_One_Component
		Enabled_ViewDimless = Is_At_Least_One_Component
		'ACTUATE ENABLE/DISABLE VARIABLES TO CONTROLS/MENUS.
		'---- ADD/DELETE/EDIT.
		frmMain.cmdADEComponent(0).Enabled = Enabled_Add
		frmMain.cmdADEComponent(1).Enabled = Enabled_Delete
		frmMain.cmdADEComponent(2).Enabled = Enabled_Edit
		''---- RESULTS MENU: PSDM, CPHSDM, ECM, COMPARE PSDM, COMPARE CPHSDM.
		'frmMain.mnuResultsItem(0).Enabled = Enabled_PSDM_Results
		'frmMain.mnuResultsItem(1).Enabled = Enabled_CPHSDM_Results
		'frmMain.mnuResultsItem(2).Enabled = Enabled_ECM_Results
		'frmMain.mnuResultsItem(3).Enabled = Enabled_PSDM_Comparison
		'frmMain.mnuResultsItem(4).Enabled = Enabled_CPHSDM_Comparison
		'---- OPTIONS MENU: FOULING, INFLUENT CONC, EFFLUENT CONC.
		frmMain.mnuOptionsItem(0).Enabled = Enabled_OptionsMenu
		frmMain.mnuOptionsItem(1).Enabled = Enabled_OptionsMenu
		frmMain.mnuOptionsItem(2).Enabled = Enabled_OptionsMenu
		'---- RUN MENU: PSDM, CPHSDM, ECM.
		frmMain.mnuRunItem(0).Enabled = Enabled_RunMenu
		frmMain.mnuRunItem(1).Enabled = Enabled_RunMenu
		frmMain.mnuRunItem(2).Enabled = Enabled_RunMenu
		frmMain.mnuRunItem(10).Enabled = False 'PSDMR-IN-ROOM. ' false to make inactive for now
		frmMain.mnuRunItem(20).Enabled = Enabled_RunMenu 'PSDMR ALONE.
		'---- FILE MENU: SAVE AND SAVE-AS.
		frmMain.mnuFileItem(2).Enabled = Enabled_Save
		frmMain.mnuFileItem(3).Enabled = Enabled_Save
		'---- VIEW DIM'LESS GROUPS.
		frmMain.cmdViewDimensionless.Enabled = Enabled_ViewDimless
		'
		' DEMO SETTINGS.
		'
		Call frmMain.frmMain_Reset_DemoVersionDisablings()
		'
		' RE-DISPLAY ALL VALUES.
		'
		Call frmMain_Repopulate_Values()
		'/////////// FORMERLY NAMED Update_Display_Data() [ENDS] ////////////////////////////////
		'
		' RE-CALCULATE AND DISPLAY BED DENSITY.
		'
		Call Update_Bed_Density_Display()
		'
		' RE-CALCULATE AND DISPLAY BED POROSITY,
		' SUPERFICIAL VELOCITY AND INTERSTITIAL VELOCITY.
		'
		Call Update_Several_Bed_Properties(3)
		
		'/////////// FORMERLY NAMED Update_Display() ////////////////////////////////
		'RE-CALCULATE AND DISPLAY EBCT.
		dd = Bed.Flowrate
		'dd = dd * FlowConversionFactor(CInt(frmMain.txtBedUnits(3).ListIndex))
		'frmMain.txtBedValue(3) = Format$(dd, "0.000E+00")
		Call unitsys_set_number_in_base_units(frmMain.txtBedValue(3), dd) 'FLOW RATE.
		'  dd = Bed.Length * Pi * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#   'EBCT in min
		dd = Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate 'EBCT in sec
		'dd = dd * TimeConversionFactor(CInt(frmMain.txtBedUnits(4).ListIndex))
		'frmMain.txtBedValue(4) = Format_It(dd, 2)
		Call unitsys_set_number_in_base_units(frmMain.txtBedValue(4), dd) 'EBCT.
		
		'RE-CALCULATE SPDFR CORRELATION VALUE FOR EACH COMPONENT.
		If (Number_Component > 0) Then
			For i = 1 To Number_Component
				If (Component(i).Use_SPDFR_Correlation) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object SPDFR_Corr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Component(i).SPDFR = SPDFR_Corr(Component_Number_Selected)
				End If
			Next i
		End If
		'/////////// FORMERLY NAMED Update_Display() [ENDS] //////////////////////////////////
		
		'ADDED 9/4/98.
		If (FileNote = "") Then
			frmMain.cmdNote(0).Visible = True
			frmMain.cmdNote(1).Visible = False
		Else
			frmMain.cmdNote(0).Visible = False
			frmMain.cmdNote(1).Visible = True
		End If
		
	End Sub
	
	
	Private Sub frmCompoProp_Repopulate_Values()
		Dim Frm As frmCompoProp
		Frm = frmCompoProp
		'UPDATE NUMERIC VALUES TO WINDOW.
		'---- MAIN BLOCK.
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call AssignTextAndTag(Frm.txtDataComponentProperty(0), Trim(Component(0).Name))
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(1), Component(0).MW)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(2), Component(0).MolarVolume)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(3), Component(0).BP)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(4), Component(0).InitialConcentration)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(10), Component(0).Liquid_Density)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(9), Component(0).Aqueous_Solubility)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(7), Component(0).Vapor_Pressure)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(8), Component(0).Refractive_Index)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(11), CDbl(Component(0).CAS))
		'---- FREUNDLICH K AND 1/N BLOCK.
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(5), Component(0).Use_K)
		'UPGRADE_ISSUE: Control txtDataComponentProperty could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDataComponentProperty(6), Component(0).Use_OneOverN)
	End Sub
	Sub frmCompoProp_Refresh()
		'RE-DISPLAY ALL VALUES.
		Call frmCompoProp_Repopulate_Values()
		
		
		
	End Sub
	
	
	Sub frmInputParamsPSDMInRoom_Repopulate_Values(ByRef Temp_RP As RoomParam_Type, ByRef in_NOW_CONTAMINANT As Short)
		Dim Frm As frmInputParamsPSDMInRoom
		Frm = frmInputParamsPSDMInRoom
		'UPDATE NUMERIC VALUES TO WINDOW.
		'---- MAIN BLOCK.
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(0), Temp_RP.ROOM_VOL)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(1), Temp_RP.ROOM_FLOWRATE)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(2), Temp_RP.ROOM_C0(in_NOW_CONTAMINANT))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(3), Temp_RP.ROOM_EMIT(in_NOW_CONTAMINANT))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(4), Temp_RP.INITIAL_ROOM_CONC(in_NOW_CONTAMINANT))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(5), Temp_RP.RXN_RATE_CONSTANT(in_NOW_CONTAMINANT))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(6), Temp_RP.RXN_RATIO(in_NOW_CONTAMINANT))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(7), Component(in_NOW_CONTAMINANT).Use_K)
	End Sub
	Sub frmInputParamsPSDMInRoom_Refresh(ByRef Temp_RP As RoomParam_Type, ByRef in_NOW_CONTAMINANT As Short)
		Dim Frm As frmInputParamsPSDMInRoom
		Frm = frmInputParamsPSDMInRoom
		Dim boolNewSetting As Boolean
		Dim intNewTag As Short
		'
		' IMPORTANT TO DO THIS HERE:
		'
		Component(0).MW = Component(in_NOW_CONTAMINANT).MW
		Component(0).Use_OneOverN = Component(in_NOW_CONTAMINANT).Use_OneOverN
		'
		' OTHER CODE CONTINUES ... .
		'
		'
		' RE-DISPLAY ALL VALUES.
		'
		Call frmInputParamsPSDMInRoom_Repopulate_Values(Temp_RP, in_NOW_CONTAMINANT)
		'Dim newVal As Double
		'Dim ConversionFactor As Double
		''VOLUME OF ROOM.
		'ConversionFactor = VolumeConversionFactor(CInt(cboUnits(0).ListIndex))
		'newVal = TempData.ROOM_VOL * ConversionFactor
		'Call AssignTextAndTag_WithRange(txtData(0), newVal, 1E-20, 1E+20)
		''FLOW RATE OF AIR THROUGH ROOM.
		'ConversionFactor = FlowConversionFactor(CInt(cboUnits(1).ListIndex))
		'newVal = TempData.ROOM_FLOWRATE * ConversionFactor
		'Call AssignTextAndTag_WithRange(txtData(1), newVal, 1E-20, 1E+20)
		'DISPLAY CALCULATED PARAMETERS.
		'UPGRADE_ISSUE: Control lblAirRate could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblAirRate.Text = NumberToMFBString(Temp_RP.ROOM_CHANGE_RATE)
		'ENABLING/DISABLING VARIOUS STUFF.
		If (Temp_RP.COUNT_CONTAMINANT = 0) Then
			'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.txtData(2).Enabled = False
			'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.txtData(3).Enabled = False
			'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.txtData(4).Enabled = False
			'UPGRADE_ISSUE: Control cboChemical could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.cboChemical.Enabled = False
			'DISPLAY CALCULATED PARAMETERS.
			'UPGRADE_ISSUE: Control lblSSValue could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.lblSSValue.Enabled = False
			'UPGRADE_ISSUE: Control lblSSValue could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.lblSSValue.Text = "n/a"
		Else
			''CONCENTRATION OF CONTAMINANT IN THE AIR STREAM INFLUENT TO THE ROOM.
			'ConversionFactor = ConcentrationConversionFactor(CInt(cboUnits(2).ListIndex))
			'newVal = TempData.ROOM_C0(NOW_CONTAMINANT) * ConversionFactor
			'Call AssignTextAndTag_WithRange(txtData(2), newVal, 0#, 1E+20)
			''MASS EMISSION RATE OF CONTAMINANT.
			'ConversionFactor = MassEmissionRateConversionFactor(CInt(cboUnits(3).ListIndex))
			'newVal = TempData.ROOM_EMIT(NOW_CONTAMINANT) * ConversionFactor
			'Call AssignTextAndTag_WithRange(txtData(3), newVal, 0#, 1E+20)
			''CONCENTRATION OF CONTAMINANT IN ROOM AT TIME = ZERO.
			'ConversionFactor = ConcentrationConversionFactor(CInt(cboUnits(4).ListIndex))
			'newVal = TempData.INITIAL_ROOM_CONC(NOW_CONTAMINANT) * ConversionFactor
			'Call AssignTextAndTag_WithRange(txtData(4), newVal, 0#, 1E+20)
			'ENABLE TEXT BOXES.
			'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.txtData(2).Enabled = True
			'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.txtData(3).Enabled = True
			'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.txtData(4).Enabled = True
			'UPGRADE_ISSUE: Control cboChemical could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.cboChemical.Enabled = True
			'DISPLAY CALCULATED PARAMETERS.
			'UPGRADE_ISSUE: Control lblSSValue could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.lblSSValue.Enabled = True
			'UPGRADE_ISSUE: Control lblSSValue could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.lblSSValue.Text = NumberToMFBString(Temp_RP.ROOM_SS_VALUE(in_NOW_CONTAMINANT))
		End If
		'UPGRADE_ISSUE: Control cboChemical could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If (Frm.cboChemical.Items.Count > 0) Then
			''''''Frm.ssframe_ContaminantProps.Caption = "Properties of " & Frm.cboChemical.List(Frm.cboChemical.ListIndex) & ":"
			''Frm.sspContaminantProps.Caption = "Properties of " & Frm.cboChemical.List(Frm.cboChemical.ListIndex) & ":"
		Else
			''''''Frm.ssframe_ContaminantProps.Caption = "No Contaminants Defined"
			''Frm.sspContaminantProps.Caption = "No Contaminants Defined"
		End If
		'UPGRADE_ISSUE: Control sspContaminantProps could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'
		' LOOK UP INDEX FOR THIS CHEMICAL.
		'
		'UPGRADE_ISSUE: Control HALT_cbo_RXN_PRODUCT could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.HALT_cbo_RXN_PRODUCT = True
		Dim i As Short
		Dim Ctl As ComboBox
		'UPGRADE_ISSUE: Control cbo_RXN_PRODUCT could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Ctl = Frm.cbo_RXN_PRODUCT
		'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ctl.SelectedIndex = -1
		'UPGRADE_ISSUE: Control cbo_RXN_PRODUCT could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		For i = 0 To Frm.cbo_RXN_PRODUCT.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (Ctl.Items(i).ItemData = Temp_RP.RXN_PRODUCT(in_NOW_CONTAMINANT)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Ctl.SelectedIndex = i
				Exit For
			End If
		Next i
		'UPGRADE_ISSUE: Control HALT_cbo_RXN_PRODUCT could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.HALT_cbo_RXN_PRODUCT = False
		'
		' SELECT APPROPRIATE optTimeVarConc SETTING.
		'
		boolNewSetting = Temp_RP.bool_ROOM_COINI_ISTIMEVAR(in_NOW_CONTAMINANT)
		intNewTag = IIf(boolNewSetting, 1, 0)
		'UPGRADE_ISSUE: Control HALT_ALL_CONTROLS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.HALT_ALL_CONTROLS = True
		'UPGRADE_ISSUE: Control optTimeVarConc could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If intNewTag = 0 Then
			'Frm._optTimeVarConc_0.Value = True
			'UPGRADE_ISSUE: Control optTimeVarConc could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm._optTimeVarConc_1.Value = False
			'UPGRADE_ISSUE: Control optTimeVarConc could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Else
			'Frm._optTimeVarConc_1.Value = True
			'Frm._optTimeVarConc_0.Value = False
		End If


		'Frm._optTimeVarConc_0.Enabled = True
		'UPGRADE_ISSUE: Control optTimeVarConc could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm._optTimeVarConc_1.Enabled = True
		'UPGRADE_ISSUE: Control optTimeVarConc could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm._optTimeVarConc_0.Tag = Trim(Str(intNewTag))
		'UPGRADE_ISSUE: Control cmdTimeVarConc could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm.cmdTimeVarConc.Enabled = boolNewSetting
		'UPGRADE_ISSUE: Control HALT_ALL_CONTROLS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.HALT_ALL_CONTROLS = False
		'
		' SELECT APPROPRIATE optTimeVarEmit SETTING.
		'
		boolNewSetting = Temp_RP.bool_ROOM_EMITINI_ISTIMEVAR(in_NOW_CONTAMINANT)
		intNewTag = IIf(boolNewSetting, 1, 0)
		'UPGRADE_ISSUE: Control HALT_ALL_CONTROLS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.HALT_ALL_CONTROLS = True
		'UPGRADE_ISSUE: Control optTimeVarEmit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If intNewTag = 0 Then
			'Frm._optTimeVarEmit_0.Value = True
			'UPGRADE_ISSUE: Control optTimeVarEmit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm._optTimeVarEmit_1.Value = False
		Else
			'Frm._optTimeVarEmit_1.Value = True
			'UPGRADE_ISSUE: Control optTimeVarEmit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm._optTimeVarEmit_0.Value = False
		End If

		'UPGRADE_ISSUE: Control optTimeVarEmit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm._optTimeVarEmit_0.Enabled = True
		''UPGRADE_ISSUE: Control optTimeVarEmit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm._optTimeVarEmit_1.Enabled = True
		''UPGRADE_ISSUE: Control optTimeVarEmit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm._optTimeVarEmit_0.Tag = Trim(Str(intNewTag))
		''UPGRADE_ISSUE: Control cmdTimeVarEmit could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm.cmdTimeVarEmit.Enabled = boolNewSetting
		'UPGRADE_ISSUE: Control HALT_ALL_CONTROLS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.HALT_ALL_CONTROLS = False
		'
		' SELECT APPROPRIATE optTimeVarK SETTING.
		'
		boolNewSetting = Temp_RP.bool_ROOM_KINI_ISTIMEVAR(in_NOW_CONTAMINANT)
		intNewTag = IIf(boolNewSetting, 1, 0)
		'UPGRADE_ISSUE: Control HALT_ALL_CONTROLS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.HALT_ALL_CONTROLS = True
		'UPGRADE_ISSUE: Control optTimeVarK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If intNewTag = 0 Then
			'Frm._optTimeVarK_0.Value = True
			''UPGRADE_ISSUE: Control optTimeVarK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm._optTimeVarK_1.Value = False
			''UPGRADE_ISSUE: Control optTimeVarK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Else
			'Frm._optTimeVarK_1.Value = True
			''UPGRADE_ISSUE: Control optTimeVarK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm._optTimeVarK_0.Value = False
		End If

		'Frm._optTimeVarK_0.Enabled = True
		'	'UPGRADE_ISSUE: Control optTimeVarK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'	Frm._optTimeVarK_1.Enabled = True
		''UPGRADE_ISSUE: Control optTimeVarK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm._optTimeVarK_0.Tag = Trim(Str(intNewTag))
		''UPGRADE_ISSUE: Control cmdTimeVarK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm.cmdTimeVarK.Enabled = boolNewSetting
		'UPGRADE_ISSUE: Control HALT_ALL_CONTROLS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.HALT_ALL_CONTROLS = False
	End Sub
	
	
	Sub frmKinetic_Repopulate_Values()
		Dim Frm As frmKinetic
		Frm = frmKinetic
		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtKF could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtKF, Component(0).KP_User_Input(1))
		'UPGRADE_ISSUE: Control txtDS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDS, Component(0).KP_User_Input(2))
		'UPGRADE_ISSUE: Control txtDP could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtDP, Component(0).KP_User_Input(3))
		'UPGRADE_ISSUE: Control txtSPDFR could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtSPDFR, Component(0).SPDFR)
		'UPGRADE_ISSUE: Control txtTort could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtTort, Component(0).Tortuosity)
	End Sub
	Sub frmKinetic_Refresh()
		Dim Frm As frmKinetic
		Frm = frmKinetic
		'RE-DISPLAY ALL VALUES.
		Call frmKinetic_Repopulate_Values()
		'DISPLAY CORRELATION NAMES.
		'UPGRADE_ISSUE: Control lblCorrelationKF could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblCorrelationKF.Text = Get_Correlation_Description(0)
		'UPGRADE_ISSUE: Control lblCorrelationDS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblCorrelationDS.Text = Get_Correlation_Description(1)
		'UPGRADE_ISSUE: Control lblCorrelationDP could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblCorrelationDP.Text = Get_Correlation_Description(2)
		'DISPLAY USER/CORRELATION OPTIONBOXES.
		'UPGRADE_ISSUE: Control lblCorrelationKF could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control optKF could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblCorrelationKF.Enabled = Frm.optKF(1).Checked
		'UPGRADE_ISSUE: Control lblCorrelationDS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control optDS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblCorrelationDS.Enabled = Frm.optDS(1).Checked
		'UPGRADE_ISSUE: Control lblCorrelationDP could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control optDP could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblCorrelationDP.Enabled = Frm.optDP(1).Checked
		'DISPLAY CORRELATION OUTPUTS.
		'UPGRADE_ISSUE: Control lblKF could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblKF.Text = VB6.Format(kf(0), "0.00E+00")
		'UPGRADE_ISSUE: Control lblDS could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblDS.Text = VB6.Format(Ds(0), "0.00E+00")
		'UPGRADE_ISSUE: Control lblDP could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.lblDP.Text = VB6.Format(Dp(0), "0.00E+00")




	End Sub
	
	
	Sub frmFreundlich_Show_KNData(ByRef lblx As System.Windows.Forms.Control, ByRef NowVal As Double, ByRef UseFormat As String)
		If (NowVal = -1#) Then
			lblx.ForeColor = System.Drawing.ColorTranslator.FromOle(QBColor(12))
			lblx.Text = "Unavailable"
		Else
			lblx.ForeColor = System.Drawing.ColorTranslator.FromOle(QBColor(0))
			lblx.Text = VB6.Format(NowVal, UseFormat)
		End If
	End Sub
	Sub frmFreundlich_Repopulate_Values()
		Dim Frm As frmFreundlich
		Frm = frmFreundlich
		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtInput(11), Component(0).IPES_OrderOfMagnitude)
		'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtInput(12), CDbl(Component(0).IPES_NumRegressionPts))
		'UPGRADE_ISSUE: Control UserOneOverN could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.UserOneOverN, Component(0).UserEntered_OneOverN)
		'UPGRADE_ISSUE: Control UserK could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.UserK, Component(0).UserEntered_K)
	End Sub
	Sub frmFreundlich_Refresh()
		Dim Frm As frmFreundlich
		Frm = frmFreundlich
		Dim Ctl As GroupBox
		Dim Ctl_Inv1 As GroupBox
		Dim Ctl_Inv2 As GroupBox
		Dim Avail_Height As Integer
		Dim Avail_Width As Integer
		Dim XXX As Integer
		Dim yyy As Integer
		Dim WhichSelected As Short
		Dim SelectedOption As RadioButton
		Dim temp As String
		'Debug.Print "frmFreundlich_Refresh()"
		'REDISPLAY ALL VALUES.
		Call frmFreundlich_Repopulate_Values()
		'CENTER THE TOP FRAME.
		'UPGRADE_ISSUE: Control fraSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.grpSource.Left = (Frm.ClientRectangle.Width - Frm.grpSource.Width) / 2
		'SET UP FRAMES APPROPRIATELY.
		'default to radio button 1 stuff
		Ctl = Frm.grpIsothermDB
		'UPGRADE_ISSUE: Control fraIPES could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Ctl_Inv1 = Frm.grpIPES
		'UPGRADE_ISSUE: Control fraUserInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Ctl_Inv2 = Frm.grpUserInput


		'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If (Frm.RadioButton1.Checked) Then
			'UPGRADE_ISSUE: Control fraIsothermDB could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl = Frm.grpIsothermDB
			'UPGRADE_ISSUE: Control fraIPES could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl_Inv1 = Frm.grpIPES
			'UPGRADE_ISSUE: Control fraUserInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl_Inv2 = Frm.grpUserInput
		End If
		'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If (Frm.RadioButton2.Checked) Then
			'UPGRADE_ISSUE: Control fraIsothermDB could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl_Inv1 = Frm.grpIsothermDB
			'UPGRADE_ISSUE: Control fraIPES could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl = Frm.grpIPES
			'UPGRADE_ISSUE: Control fraUserInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl_Inv2 = Frm.grpUserInput
		End If
		'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If (Frm.RadioButton3.Checked) Then
			'UPGRADE_ISSUE: Control fraIsothermDB could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl_Inv1 = Frm.grpIsothermDB
			'UPGRADE_ISSUE: Control fraIPES could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl_Inv2 = Frm.grpIPES
			'UPGRADE_ISSUE: Control fraUserInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Ctl = Frm.grpUserInput
		End If
		Ctl.Visible = True
		Ctl_Inv1.Visible = False
		Ctl_Inv2.Visible = False
		'UPGRADE_ISSUE: Control fraSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_ISSUE: Control sspanel_StatusBar could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Avail_Height = Frm.StatusStrip1.Top - (Frm.grpSource.Top + Frm.grpSource.Height)
		Avail_Width = Frm.ClientRectangle.Width
		XXX = (Avail_Width - Ctl.Width) / 2
		'UPGRADE_ISSUE: Control fraSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		yyy = Frm.grpSource.Top + Frm.grpSource.Height + (Avail_Height - Ctl.Height) / 2
		Ctl.SetBounds((XXX), (yyy), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		'VALIDATE/INVALIDATE SOURCES.
		If (Component(0).IsothermDB_K > 0#) And (Component(0).IsothermDB_OneOverN > 0#) Then
			'VALIDATE ISOTHERM DB AS SOURCE.
			'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.RadioButton1.Text = "Isotherm &Database"
		Else
			'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.RadioButton1.Text = "(Isotherm &Database)"
		End If
		If (Component(0).IPESResult_K > 0#) And (Component(0).IPESResult_OneOverN > 0#) Then
			'VALIDATE IPE CALCULATION AS SOURCE.
			'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.RadioButton2.Text = "Isotherm Parameter &Estimation"
		Else
			'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.RadioButton2.Text = "(Isotherm Parameter &Estimation)"
		End If
		'ENSURE PROPER SOURCE IS CHECKED.
		'HALT_OPTFREUNDLICHSOURCE = True
		Select Case Component(0).Source_KandOneOverN
			Case KNSOURCE_ISOTHERMDB
				'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm.RadioButton1.Checked = True
			Case KNSOURCE_IPES
				'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm.RadioButton2.Checked = True
			Case KNSOURCE_USERINPUT
				'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Frm.RadioButton3.Checked = True
		End Select
		'HALT_OPTFREUNDLICHSOURCE = False
		'DETERMINE WHICH OPTION WAS SELECTED.
		'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'

		If (Frm.RadioButton1.Checked) Then
			WhichSelected = 0
			SelectedOption = Frm.RadioButton1
		End If
		'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If (Frm.RadioButton2.Checked) Then
			WhichSelected = 1
			SelectedOption = Frm.RadioButton2
		End If
		'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If (Frm.RadioButton3.Checked) Then
			WhichSelected = 2
			SelectedOption = Frm.RadioButton3
		End If
		'DISPLAY WARNING IF NEEDED.
		'UPGRADE_ISSUE: Control sspanel_Warning could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'Frm.sspanel_Warning.Visible = False
		'UPGRADE_ISSUE: Control optFreundlichSource could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If (Left(SelectedOption.Text, 1) = "(") Then
			'UPGRADE_ISSUE: Control sspanel_Warning could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			'Frm.sspanel_Warning.Visible = True
			Select Case WhichSelected
					Case 0
						temp = "You must select an isotherm from the isotherm " & "database.  To do so, select a component on the left, " & "and then select an isotherm record " & "on the right.  " & "If you do not, K and 1/n source will " & "revert to user-input."
						'UPGRADE_ISSUE: Control lblWarning could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						Frm.lblWarning.Text = temp
					Case 1
						temp = "You must calculate K and 1/n using IPE.  " & "To do so, click on the button marked " & Chr(34) & "Perform IPE Calculations" & Chr(34) & " from within this screen.  If you do not, " & "K and 1/n source will revert to user-input."
						'UPGRADE_ISSUE: Control lblWarning could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
						Frm.lblWarning.Text = temp
				End Select
			End If
			'DISPLAY CURRENT POLANYI PARAMETERS.
			'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Frm.txtInput(13).Text = Trim(Carbon.Name)
		'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtInput(0).Text = VB6.Format(Carbon.W0, "0.000E+00")
		'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtInput(1).Text = VB6.Format(Carbon.BB, "0.000E+00")
		'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtInput(10).Text = VB6.Format(Carbon.PolanyiExponent, "0.000E+00")
		'DISPLAY CURRENT IPE K AND 1/N.
		'UPGRADE_ISSUE: Control lblValue could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call frmFreundlich_Show_KNData(Frm.lblValue(4), Component(0).IPESResult_OneOverN, "0.000")
		'UPGRADE_ISSUE: Control lblValue could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call frmFreundlich_Show_KNData(Frm.lblValue(5), Component(0).IPESResult_K, "###,##0.0")
		'DISPLAY CURRENT ISOTHERM DATABASE K AND 1/N.
		'UPGRADE_ISSUE: Control lblValue could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call frmFreundlich_Show_KNData(Frm.lblValue(1), Component(0).IsothermDB_OneOverN, "0.000")
		'UPGRADE_ISSUE: Control lblValue could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call frmFreundlich_Show_KNData(Frm.lblValue(0), Component(0).IsothermDB_K, "###,##0.0")
		
		
		
		
	End Sub
	
	
	
	Sub frmPolanyi_Repopulate_Values()
		Dim Frm As frmPolanyi
		Frm = frmPolanyi
		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtInput(0), Carbon.W0)
		'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtInput(1), Carbon.BB)
		'UPGRADE_ISSUE: Control txtInput could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtInput(2), Carbon.PolanyiExponent)
	End Sub
	Sub frmPolanyi_Refresh()
		'Dim frm As Form
		'Set frm = frmPolanyi
		'RE-DISPLAY ALL VALUES.
		Call frmPolanyi_Repopulate_Values()
	End Sub
	
	
	Sub frmFluidProps_Repopulate_Values()
		Dim Frm As frmFluidProps
		Frm = frmFluidProps
		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtWater could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtWater(0), Bed.WaterDensity)
		'UPGRADE_ISSUE: Control txtWater could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtWater(1), Bed.WaterViscosity)
	End Sub
	Sub frmFluidProps_Refresh()
		Dim Frm As frmFluidProps
		Frm = frmFluidProps
		Call frmFluidProps_Repopulate_Values()
		'UPDATE CORRELATION USAGE BOXES.
		'UPGRADE_ISSUE: Control chkCorr could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm._chkCorr_0.Checked = State_Check_Water(1)
		'UPGRADE_ISSUE: Control chkCorr could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm._chkCorr_1.Checked = State_Check_Water(2)
		'UPGRADE_ISSUE: Control txtWater could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtWater(0).Enabled = Not State_Check_Water(1)
		'UPGRADE_ISSUE: Control txtWater could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtWater(1).Enabled = Not State_Check_Water(2)
	End Sub
	
	
	Sub frmEditAdsorberData_Repopulate_Values()
		Dim Frm As frmEditAdsorberData
		Frm = frmEditAdsorberData
		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(0), CDbl(Val(frmEditAdsorberData_Record.InternalArea)))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(1), CDbl(Val(frmEditAdsorberData_Record.MaxCapacity)))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(2), CDbl(Val(frmEditAdsorberData_Record.OutsideDiameter)))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(5), CDbl(Val(frmEditAdsorberData_Record.DefaultFlowRate)))
		'TEXT DATA.
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(7).Text = Trim(frmEditAdsorberData_Record.PartNumber)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(3).Text = Trim(frmEditAdsorberData_Record.DesignPressure)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(4).Text = Trim(frmEditAdsorberData_Record.DesignFlowRange)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(6).Text = Trim(frmEditAdsorberData_Record.Note)
	End Sub
	Sub frmEditAdsorberData_Refresh()
		'Dim frm As Form
		'Set frm = frmEditAdsorberData
		Call frmEditAdsorberData_Repopulate_Values()
	End Sub
	
	
	Sub frmEditCarbonData_Repopulate_Values()
		Dim Frm As frmEditCarbonData
		Frm = frmEditCarbonData
		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(0), CDbl(Val(CStr(frmEditCarbonData_Record.AppDen))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(1), CDbl(Val(CStr(frmEditCarbonData_Record.ParticleRadius))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(2), CDbl(Val(CStr(frmEditCarbonData_Record.ParticlePorosity))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(4), CDbl(Val(CStr(frmEditCarbonData_Record.W0))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(5), CDbl(Val(CStr(frmEditCarbonData_Record.BB))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(6), CDbl(Val(CStr(frmEditCarbonData_Record.PolanyiExponent))))
		'TEXT DATA.
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(7).Text = Trim(frmEditCarbonData_Record.Name)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(3).Text = Trim(frmEditCarbonData_Record.AdsType)
	End Sub
	Sub frmEditCarbonData_Refresh()
		Dim Frm As System.Windows.Forms.Form
		Frm = frmEditCarbonData
		Call frmEditCarbonData_Repopulate_Values()
	End Sub
	
	
	Sub frmEditIsothermData_Repopulate_Values()
		Dim Frm As frmEditIsothermData
		Frm = frmEditIsothermData
		'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(0), CDbl(Val(CStr(frmEditIsothermData_Record.k))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(1), CDbl(Val(CStr(frmEditIsothermData_Record.Cmin))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(2), CDbl(Val(CStr(frmEditIsothermData_Record.pHmin))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(3), CDbl(Val(frmEditIsothermData_Record.Tmin)))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(4), CDbl(Val(CStr(frmEditIsothermData_Record.OneOverN))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(5), CDbl(Val(CStr(frmEditIsothermData_Record.Cmax))))
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call unitsys_set_number_in_base_units(Frm.txtData(6), CDbl(Val(CStr(frmEditIsothermData_Record.pHmax))))
		'TEXT DATA.
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(7).Text = Trim(frmEditIsothermData_Record.CarbonName)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(8).Text = Trim(frmEditIsothermData_Record.CAS)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(9).Text = Trim(frmEditIsothermData_Record.Name)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(10).Text = Trim(frmEditIsothermData_Record.Source)
		'UPGRADE_ISSUE: Control txtData could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Frm.txtData(11).Text = Trim(frmEditIsothermData_Record.Comments)
	End Sub
	Sub frmEditIsothermData_Refresh()
		Dim Frm As System.Windows.Forms.Form
		Frm = frmEditIsothermData
		Call frmEditIsothermData_Repopulate_Values()
	End Sub
End Module