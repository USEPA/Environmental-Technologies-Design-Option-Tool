Attribute VB_Name = "Refresh"
Option Explicit



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
  Call unitsys_set_number_in_base_units(frmMain.txtNumberOfBeds, CDbl(Bed.NumberOfBeds))
  Call unitsys_set_number_in_base_units(frmMain.txtNPoint(0), CDbl(MC))
  Call unitsys_set_number_in_base_units(frmMain.txtNPoint(1), CDbl(NC))
  Call unitsys_set_number_in_base_units(frmMain.txtTime(0), TimeP.End)
  Call unitsys_set_number_in_base_units(frmMain.txtTime(1), TimeP.Init)
  Call unitsys_set_number_in_base_units(frmMain.txtTime(2), TimeP.Step)
  'frmMain.txtNumberOfBeds = Format$(Bed.NumberOfBeds, "0")
  'frmMain.txtNPoint(0) = Format$(MC, "0")
  'frmMain.txtNPoint(1) = Format$(NC, "0")
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
  Call AssignTextAndTag(frmMain.txtCarbon(0), Trim$(Carbon.Name))
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
Dim i As Integer
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
Dim SAVE_OLD_POSITION As Integer

  '/////////// FORMERLY NAMED Update_Display_Data() ////////////////////////////////
  'UPDATE COMPONENT SELECTION LIST AND SCROLLBOX.
  If (frmMain.cboSelectCompo.ListCount >= 1) And _
      (frmMain.cboSelectCompo.ListIndex >= 0) Then
    SAVE_OLD_POSITION = frmMain.cboSelectCompo.ListIndex
  Else
    SAVE_OLD_POSITION = -1
  End If
  frmMain.lstComponents.Clear
  frmMain.cboSelectCompo.Clear
  For i = 1 To Number_Component
    frmMain.cboSelectCompo.AddItem Component(i).Name
    frmMain.lstComponents.AddItem Component(i).Name
    frmMain.lstComponents.Selected(i - 1) = Component(i).Is_Selected_On_List
  Next i
  If (SAVE_OLD_POSITION <> -1) And _
      (SAVE_OLD_POSITION <= frmMain.cboSelectCompo.ListCount - 1) Then
    frmMain.cboSelectCompo.ListIndex = SAVE_OLD_POSITION
  Else
    If (frmMain.cboSelectCompo.ListCount >= 1) Then
      frmMain.cboSelectCompo.ListIndex = SAVE_OLD_POSITION = 0
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
  Enabled_Add = True              'ENABLE ADD.
  Enabled_Delete = True           'ENABLE DELETE.
  Enabled_Edit = True             'ENABLE EDIT.
  If (Number_Component = Number_Compo_Max) Then
    Enabled_Add = False             'DISABLE ADD.
  End If
  If (Number_Component = 0) Then
    Enabled_Delete = False        'DISABLE DELETE.
    Enabled_Edit = False          'DISABLE EDIT.
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
  frmMain.mnuRunItem(10).Enabled = Enabled_RunMenu    'PSDMR-IN-ROOM.
  frmMain.mnuRunItem(20).Enabled = Enabled_RunMenu    'PSDMR ALONE.
  '---- FILE MENU: SAVE AND SAVE-AS.
  frmMain.mnuFileItem(2).Enabled = Enabled_Save
  frmMain.mnuFileItem(3).Enabled = Enabled_Save
  '---- VIEW DIM'LESS GROUPS.
  frmMain.cmdViewDimensionless.Enabled = Enabled_ViewDimless
  '
  ' DEMO SETTINGS.
  '
  Call frmMain.frmMain_Reset_DemoVersionDisablings
  '
  ' RE-DISPLAY ALL VALUES.
  '
  Call frmMain_Repopulate_Values
  '/////////// FORMERLY NAMED Update_Display_Data() [ENDS] ////////////////////////////////
  '
  ' RE-CALCULATE AND DISPLAY BED DENSITY.
  '
  Call Update_Bed_Density_Display
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
  Call unitsys_set_number_in_base_units(frmMain.txtBedValue(3), dd)   'FLOW RATE.
'  dd = Bed.Length * Pi * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#   'EBCT in min
  dd = Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate          'EBCT in sec
  'dd = dd * TimeConversionFactor(CInt(frmMain.txtBedUnits(4).ListIndex))
  'frmMain.txtBedValue(4) = Format_It(dd, 2)
  Call unitsys_set_number_in_base_units(frmMain.txtBedValue(4), dd)   'EBCT.
  
  'RE-CALCULATE SPDFR CORRELATION VALUE FOR EACH COMPONENT.
  If (Number_Component > 0) Then
    For i = 1 To Number_Component
      If (Component(i).Use_SPDFR_Correlation) Then
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
Dim Frm As Form
Set Frm = frmCompoProp
  'UPDATE NUMERIC VALUES TO WINDOW.
  '---- MAIN BLOCK.
  Call AssignTextAndTag( _
      Frm.txtDataComponentProperty(0), Trim$(Component(0).Name))
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(1), Component(0).MW)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(2), Component(0).MolarVolume)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(3), Component(0).BP)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(4), Component(0).InitialConcentration)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(10), Component(0).Liquid_Density)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(9), Component(0).Aqueous_Solubility)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(7), Component(0).Vapor_Pressure)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(8), Component(0).Refractive_Index)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(11), CDbl(Component(0).CAS))
  '---- FREUNDLICH K AND 1/N BLOCK.
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(5), Component(0).Use_K)
  Call unitsys_set_number_in_base_units( _
      Frm.txtDataComponentProperty(6), Component(0).Use_OneOverN)
End Sub
Sub frmCompoProp_Refresh()
  'RE-DISPLAY ALL VALUES.
  Call frmCompoProp_Repopulate_Values



End Sub


Sub frmInputParamsPSDMInRoom_Repopulate_Values( _
    Temp_RP As RoomParam_Type, _
    in_NOW_CONTAMINANT As Integer)
Dim Frm As Form
Set Frm = frmInputParamsPSDMInRoom
  'UPDATE NUMERIC VALUES TO WINDOW.
  '---- MAIN BLOCK.
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(0), Temp_RP.ROOM_VOL)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(1), Temp_RP.ROOM_FLOWRATE)
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(2), Temp_RP.ROOM_C0(in_NOW_CONTAMINANT))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(3), Temp_RP.ROOM_EMIT(in_NOW_CONTAMINANT))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(4), Temp_RP.INITIAL_ROOM_CONC(in_NOW_CONTAMINANT))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(5), Temp_RP.RXN_RATE_CONSTANT(in_NOW_CONTAMINANT))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(6), Temp_RP.RXN_RATIO(in_NOW_CONTAMINANT))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(7), Component(in_NOW_CONTAMINANT).Use_K)
End Sub
Sub frmInputParamsPSDMInRoom_Refresh( _
    Temp_RP As RoomParam_Type, _
    in_NOW_CONTAMINANT As Integer)
Dim Frm As Form
Set Frm = frmInputParamsPSDMInRoom
Dim boolNewSetting As Boolean
Dim intNewTag As Integer
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
  Call frmInputParamsPSDMInRoom_Repopulate_Values( _
      Temp_RP, _
      in_NOW_CONTAMINANT)
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
  Frm.lblAirRate.Caption = NumberToMFBString(Temp_RP.ROOM_CHANGE_RATE)
  'ENABLING/DISABLING VARIOUS STUFF.
  If (Temp_RP.COUNT_CONTAMINANT = 0) Then
    Frm.txtData(2).Enabled = False
    Frm.txtData(3).Enabled = False
    Frm.txtData(4).Enabled = False
    Frm.cboChemical.Enabled = False
    'DISPLAY CALCULATED PARAMETERS.
    Frm.lblSSValue.Enabled = False
    Frm.lblSSValue.Caption = "n/a"
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
    Frm.txtData(2).Enabled = True
    Frm.txtData(3).Enabled = True
    Frm.txtData(4).Enabled = True
    Frm.cboChemical.Enabled = True
    'DISPLAY CALCULATED PARAMETERS.
    Frm.lblSSValue.Enabled = True
    Frm.lblSSValue.Caption = NumberToMFBString(Temp_RP.ROOM_SS_VALUE(in_NOW_CONTAMINANT))
  End If
  If (Frm.cboChemical.ListCount > 0) Then
    ''''''Frm.ssframe_ContaminantProps.Caption = "Properties of " & Frm.cboChemical.List(Frm.cboChemical.ListIndex) & ":"
    ''Frm.sspContaminantProps.Caption = "Properties of " & Frm.cboChemical.List(Frm.cboChemical.ListIndex) & ":"
  Else
    ''''''Frm.ssframe_ContaminantProps.Caption = "No Contaminants Defined"
    ''Frm.sspContaminantProps.Caption = "No Contaminants Defined"
  End If
  Frm.sspContaminantProps.Caption = ""
  '
  ' LOOK UP INDEX FOR THIS CHEMICAL.
  '
  Frm.HALT_cbo_RXN_PRODUCT = True
Dim i As Integer
Dim Ctl As Control
Set Ctl = Frm.cbo_RXN_PRODUCT
  Ctl.ListIndex = -1
  For i = 0 To Frm.cbo_RXN_PRODUCT.ListCount - 1
    If (Ctl.ItemData(i) = Temp_RP.RXN_PRODUCT(in_NOW_CONTAMINANT)) Then
      Ctl.ListIndex = i
      Exit For
    End If
  Next i
  Frm.HALT_cbo_RXN_PRODUCT = False
  '
  ' SELECT APPROPRIATE optTimeVarConc SETTING.
  '
  boolNewSetting = Temp_RP.bool_ROOM_COINI_ISTIMEVAR(in_NOW_CONTAMINANT)
  intNewTag = IIf(boolNewSetting, 1, 0)
  Frm.HALT_ALL_CONTROLS = True
  Frm.optTimeVarConc(intNewTag).Value = True
  Frm.optTimeVarConc(1 - intNewTag).Value = False
  Frm.optTimeVarConc(0).Enabled = True
  Frm.optTimeVarConc(1).Enabled = True
  Frm.optTimeVarConc(0).Tag = Trim$(Str$(intNewTag))
  Frm.cmdTimeVarConc.Enabled = boolNewSetting
  Frm.HALT_ALL_CONTROLS = False
  '
  ' SELECT APPROPRIATE optTimeVarEmit SETTING.
  '
  boolNewSetting = Temp_RP.bool_ROOM_EMITINI_ISTIMEVAR(in_NOW_CONTAMINANT)
  intNewTag = IIf(boolNewSetting, 1, 0)
  Frm.HALT_ALL_CONTROLS = True
  Frm.optTimeVarEmit(intNewTag).Value = True
  Frm.optTimeVarEmit(1 - intNewTag).Value = False
  Frm.optTimeVarEmit(0).Enabled = True
  Frm.optTimeVarEmit(1).Enabled = True
  Frm.optTimeVarEmit(0).Tag = Trim$(Str$(intNewTag))
  Frm.cmdTimeVarEmit.Enabled = boolNewSetting
  Frm.HALT_ALL_CONTROLS = False
  '
  ' SELECT APPROPRIATE optTimeVarK SETTING.
  '
  boolNewSetting = Temp_RP.bool_ROOM_KINI_ISTIMEVAR(in_NOW_CONTAMINANT)
  intNewTag = IIf(boolNewSetting, 1, 0)
  Frm.HALT_ALL_CONTROLS = True
  Frm.optTimeVarK(intNewTag).Value = True
  Frm.optTimeVarK(1 - intNewTag).Value = False
  Frm.optTimeVarK(0).Enabled = True
  Frm.optTimeVarK(1).Enabled = True
  Frm.optTimeVarK(0).Tag = Trim$(Str$(intNewTag))
  Frm.cmdTimeVarK.Enabled = boolNewSetting
  Frm.HALT_ALL_CONTROLS = False
End Sub


Sub frmKinetic_Repopulate_Values()
Dim Frm As Form
Set Frm = frmKinetic
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtKF, Component(0).KP_User_Input(1))
  Call unitsys_set_number_in_base_units( _
      Frm.txtDS, Component(0).KP_User_Input(2))
  Call unitsys_set_number_in_base_units( _
      Frm.txtDP, Component(0).KP_User_Input(3))
  Call unitsys_set_number_in_base_units( _
      Frm.txtSPDFR, Component(0).SPDFR)
  Call unitsys_set_number_in_base_units( _
      Frm.txtTort, Component(0).Tortuosity)
End Sub
Sub frmKinetic_Refresh()
Dim Frm As Form
Set Frm = frmKinetic
  'RE-DISPLAY ALL VALUES.
  Call frmKinetic_Repopulate_Values
  'DISPLAY CORRELATION NAMES.
  Frm.lblCorrelationKF.Caption = Get_Correlation_Description(0)
  Frm.lblCorrelationDS.Caption = Get_Correlation_Description(1)
  Frm.lblCorrelationDP.Caption = Get_Correlation_Description(2)
  'DISPLAY USER/CORRELATION OPTIONBOXES.
  Frm.lblCorrelationKF.Enabled = Frm.optKF(1).Value
  Frm.lblCorrelationDS.Enabled = Frm.optDS(1).Value
  Frm.lblCorrelationDP.Enabled = Frm.optDP(1).Value
  'DISPLAY CORRELATION OUTPUTS.
  Frm.lblKF = Format$(kf(0), "0.00E+00")
  Frm.lblDS = Format$(Ds(0), "0.00E+00")
  Frm.lblDP = Format$(Dp(0), "0.00E+00")




End Sub


Sub frmFreundlich_Show_KNData( _
    lblx As Control, _
    NowVal As Double, _
    UseFormat As String)
  If (NowVal = -1#) Then
    lblx.ForeColor = QBColor(12)
    lblx.Caption = "Unavailable"
  Else
    lblx.ForeColor = QBColor(0)
    lblx.Caption = Format$(NowVal, UseFormat)
  End If
End Sub
Sub frmFreundlich_Repopulate_Values()
Dim Frm As Form
Set Frm = frmFreundlich
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtInput(11), Component(0).IPES_OrderOfMagnitude)
  Call unitsys_set_number_in_base_units( _
      Frm.txtInput(12), CDbl(Component(0).IPES_NumRegressionPts))
  Call unitsys_set_number_in_base_units( _
      Frm.UserOneOverN, Component(0).UserEntered_OneOverN)
  Call unitsys_set_number_in_base_units( _
      Frm.UserK, Component(0).UserEntered_K)
End Sub
Sub frmFreundlich_Refresh()
Dim Frm As Form
Set Frm = frmFreundlich
Dim Ctl As Control
Dim Ctl_Inv1 As Control
Dim Ctl_Inv2 As Control
Dim Avail_Height As Long
Dim Avail_Width As Long
Dim XXX As Long
Dim yyy As Long
Dim WhichSelected As Integer
Dim temp As String
'Debug.Print "frmFreundlich_Refresh()"
  'REDISPLAY ALL VALUES.
  Call frmFreundlich_Repopulate_Values
  'CENTER THE TOP FRAME.
  Frm.fraSource.Left = (Frm.ScaleWidth - Frm.fraSource.Width) / 2
  'SET UP FRAMES APPROPRIATELY.
  If (Frm.optFreundlichSource(0).Value) Then
    Set Ctl = Frm.fraIsothermDB
    Set Ctl_Inv1 = Frm.fraIPES
    Set Ctl_Inv2 = Frm.fraUserInput
  End If
  If (Frm.optFreundlichSource(1).Value) Then
    Set Ctl_Inv1 = Frm.fraIsothermDB
    Set Ctl = Frm.fraIPES
    Set Ctl_Inv2 = Frm.fraUserInput
  End If
  If (Frm.optFreundlichSource(2).Value) Then
    Set Ctl_Inv1 = Frm.fraIsothermDB
    Set Ctl_Inv2 = Frm.fraIPES
    Set Ctl = Frm.fraUserInput
  End If
  Ctl.Visible = True
  Ctl_Inv1.Visible = False
  Ctl_Inv2.Visible = False
  Avail_Height = Frm.sspanel_StatusBar.Top - _
      (Frm.fraSource.Top + Frm.fraSource.Height)
  Avail_Width = Frm.ScaleWidth
  XXX = (Avail_Width - Ctl.Width) / 2
  yyy = Frm.fraSource.Top + Frm.fraSource.Height + _
      (Avail_Height - Ctl.Height) / 2
  Ctl.Move XXX, yyy
  'VALIDATE/INVALIDATE SOURCES.
  If (Component(0).IsothermDB_K > 0#) And (Component(0).IsothermDB_OneOverN > 0#) Then
    'VALIDATE ISOTHERM DB AS SOURCE.
    Frm.optFreundlichSource(0).Caption = "Isotherm &Database"
  Else
    Frm.optFreundlichSource(0).Caption = "(Isotherm &Database)"
  End If
  If (Component(0).IPESResult_K > 0#) And (Component(0).IPESResult_OneOverN > 0#) Then
    'VALIDATE IPE CALCULATION AS SOURCE.
    Frm.optFreundlichSource(1).Caption = "Isotherm Parameter &Estimation"
  Else
    Frm.optFreundlichSource(1).Caption = "(Isotherm Parameter &Estimation)"
  End If
  'ENSURE PROPER SOURCE IS CHECKED.
  'HALT_OPTFREUNDLICHSOURCE = True
  Select Case Component(0).Source_KandOneOverN
    Case KNSOURCE_ISOTHERMDB: Frm.optFreundlichSource(0).Value = True
    Case KNSOURCE_IPES: Frm.optFreundlichSource(1).Value = True
    Case KNSOURCE_USERINPUT: Frm.optFreundlichSource(2).Value = True
  End Select
  'HALT_OPTFREUNDLICHSOURCE = False
  'DETERMINE WHICH OPTION WAS SELECTED.
  If (Frm.optFreundlichSource(0)) Then WhichSelected = 0
  If (Frm.optFreundlichSource(1)) Then WhichSelected = 1
  If (Frm.optFreundlichSource(2)) Then WhichSelected = 2
  'DISPLAY WARNING IF NEEDED.
  Frm.sspanel_Warning.Visible = False
  If (Left$(Frm.optFreundlichSource(WhichSelected).Caption, 1) = "(") Then
    Frm.sspanel_Warning.Visible = True
    Select Case WhichSelected
      Case 0
        temp = "You must select an isotherm from the isotherm " & _
            "database.  To do so, select a component on the left, " & _
            "and then select an isotherm record " & _
            "on the right.  " & _
            "If you do not, K and 1/n source will " & _
            "revert to user-input."
        Frm.lblWarning.Caption = temp
      Case 1
        temp = "You must calculate K and 1/n using IPE.  " & _
            "To do so, click on the button marked " & Chr$(34) & _
            "Perform IPE Calculations" & Chr$(34) & _
            " from within this screen.  If you do not, " & _
            "K and 1/n source will revert to user-input."
        Frm.lblWarning.Caption = temp
    End Select
  End If
  'DISPLAY CURRENT POLANYI PARAMETERS.
  Frm.txtInput(13) = Trim$(Carbon.Name)
  Frm.txtInput(0) = Format$(Carbon.W0, "0.000E+00")
  Frm.txtInput(1) = Format$(Carbon.BB, "0.000E+00")
  Frm.txtInput(10) = Format$(Carbon.PolanyiExponent, "0.000E+00")
  'DISPLAY CURRENT IPE K AND 1/N.
  Call frmFreundlich_Show_KNData( _
      Frm.lblValue(4), Component(0).IPESResult_OneOverN, "0.000")
  Call frmFreundlich_Show_KNData( _
      Frm.lblValue(5), Component(0).IPESResult_K, "###,##0.0")
  'DISPLAY CURRENT ISOTHERM DATABASE K AND 1/N.
  Call frmFreundlich_Show_KNData( _
      Frm.lblValue(1), Component(0).IsothermDB_OneOverN, "0.000")
  Call frmFreundlich_Show_KNData( _
      Frm.lblValue(0), Component(0).IsothermDB_K, "###,##0.0")
  



End Sub



Sub frmPolanyi_Repopulate_Values()
Dim Frm As Form
Set Frm = frmPolanyi
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtInput(0), Carbon.W0)
  Call unitsys_set_number_in_base_units( _
      Frm.txtInput(1), Carbon.BB)
  Call unitsys_set_number_in_base_units( _
      Frm.txtInput(2), Carbon.PolanyiExponent)
End Sub
Sub frmPolanyi_Refresh()
'Dim frm As Form
'Set frm = frmPolanyi
  'RE-DISPLAY ALL VALUES.
  Call frmPolanyi_Repopulate_Values
End Sub


Sub frmFluidProps_Repopulate_Values()
Dim Frm As Form
Set Frm = frmFluidProps
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtWater(0), Bed.WaterDensity)
  Call unitsys_set_number_in_base_units( _
      Frm.txtWater(1), Bed.WaterViscosity)
End Sub
Sub frmFluidProps_Refresh()
Dim Frm As Form
Set Frm = frmFluidProps
  Call frmFluidProps_Repopulate_Values
  'UPDATE CORRELATION USAGE BOXES.
  Frm.chkCorr(0).Value = State_Check_Water(1)
  Frm.chkCorr(1).Value = State_Check_Water(2)
  Frm.txtWater(0).Locked = State_Check_Water(1)
  Frm.txtWater(1).Locked = State_Check_Water(2)
End Sub


Sub frmEditAdsorberData_Repopulate_Values()
Dim Frm As Form
Set Frm = frmEditAdsorberData
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(0), CDbl(Val(frmEditAdsorberData_Record.InternalArea)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(1), CDbl(Val(frmEditAdsorberData_Record.MaxCapacity)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(2), CDbl(Val(frmEditAdsorberData_Record.OutsideDiameter)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(5), CDbl(Val(frmEditAdsorberData_Record.DefaultFlowRate)))
  'TEXT DATA.
  Frm.txtData(7) = Trim$(frmEditAdsorberData_Record.PartNumber)
  Frm.txtData(3) = Trim$(frmEditAdsorberData_Record.DesignPressure)
  Frm.txtData(4) = Trim$(frmEditAdsorberData_Record.DesignFlowRange)
  Frm.txtData(6) = Trim$(frmEditAdsorberData_Record.Note)
End Sub
Sub frmEditAdsorberData_Refresh()
'Dim frm As Form
'Set frm = frmEditAdsorberData
  Call frmEditAdsorberData_Repopulate_Values
End Sub


Sub frmEditCarbonData_Repopulate_Values()
Dim Frm As Form
Set Frm = frmEditCarbonData
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(0), CDbl(Val(frmEditCarbonData_Record.AppDen)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(1), CDbl(Val(frmEditCarbonData_Record.ParticleRadius)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(2), CDbl(Val(frmEditCarbonData_Record.ParticlePorosity)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(4), CDbl(Val(frmEditCarbonData_Record.W0)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(5), CDbl(Val(frmEditCarbonData_Record.BB)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(6), CDbl(Val(frmEditCarbonData_Record.PolanyiExponent)))
  'TEXT DATA.
  Frm.txtData(7) = Trim$(frmEditCarbonData_Record.Name)
  Frm.txtData(3) = Trim$(frmEditCarbonData_Record.AdsType)
End Sub
Sub frmEditCarbonData_Refresh()
Dim Frm As Form
Set Frm = frmEditCarbonData
  Call frmEditCarbonData_Repopulate_Values
End Sub


Sub frmEditIsothermData_Repopulate_Values()
Dim Frm As Form
Set Frm = frmEditIsothermData
  'DISPLAY CURRENT VALUES FOR UNIT CONTROLS.
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(0), CDbl(Val(frmEditIsothermData_Record.k)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(1), CDbl(Val(frmEditIsothermData_Record.Cmin)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(2), CDbl(Val(frmEditIsothermData_Record.pHmin)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(3), CDbl(Val(frmEditIsothermData_Record.Tmin)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(4), CDbl(Val(frmEditIsothermData_Record.OneOverN)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(5), CDbl(Val(frmEditIsothermData_Record.Cmax)))
  Call unitsys_set_number_in_base_units( _
      Frm.txtData(6), CDbl(Val(frmEditIsothermData_Record.pHmax)))
  'TEXT DATA.
  Frm.txtData(7) = Trim$(frmEditIsothermData_Record.CarbonName)
  Frm.txtData(8) = Trim$(frmEditIsothermData_Record.CAS)
  Frm.txtData(9) = Trim$(frmEditIsothermData_Record.Name)
  Frm.txtData(10) = Trim$(frmEditIsothermData_Record.Source)
  Frm.txtData(11) = Trim$(frmEditIsothermData_Record.Comments)
End Sub
Sub frmEditIsothermData_Refresh()
Dim Frm As Form
Set Frm = frmEditIsothermData
  Call frmEditIsothermData_Repopulate_Values
End Sub


