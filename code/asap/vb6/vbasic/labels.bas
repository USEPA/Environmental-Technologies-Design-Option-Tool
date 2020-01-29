Attribute VB_Name = "LabelsMod"
Option Explicit

Sub LabelsAirWaterPropertiesSI()

    frmAirWaterProperties!lblAirWaterProperties(0).Caption = "Water Density (kg/m" & Chr$(179) & ")"
    frmAirWaterProperties!lblAirWaterProperties(1).Caption = "Water Viscosity (kg/m/sec)"
    frmAirWaterProperties!lblAirWaterProperties(2).Caption = "Water Surface Tension (N/m)"
    frmAirWaterProperties!lblAirWaterProperties(3).Caption = "Air Density (kg/m" & Chr$(179) & ")"
    frmAirWaterProperties!lblAirWaterProperties(4).Caption = "Air Viscosity (kg/m/sec)"

End Sub

Sub LabelsBubble(UnitsType As Integer)
Dim ThisUnit As Integer

  'Operating Conditions.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = PRESSURE_PA
    Case UNITSTYPE_ENGLISH: ThisUnit = PRESSURE_ATM
  End Select
  Call Populate_Pressure_Units(frmBubble!UnitsOpCond(0), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = TEMPERATURE_C
    Case UNITSTYPE_ENGLISH: ThisUnit = TEMPERATURE_F
  End Select
  Call Populate_Temperature_Units(frmBubble!UnitsOpCond(1), ThisUnit)
  
  'Oxygen (reference compound).
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = DIFFUSIVITY_M2_per_S
    Case UNITSTYPE_ENGLISH: ThisUnit = DIFFUSIVITY_FT2_per_S
  End Select
  Call Populate_Diffusivity_Units(frmBubble!UnitsOxygenRef(1), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = INVERSETIME_S
    Case UNITSTYPE_ENGLISH: ThisUnit = INVERSETIME_S
  End Select
  Call Populate_InverseTime_Units(frmBubble!UnitsOxygenRef(2), ThisUnit)
  
  'Design Contaminant.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = CONCENTRATION_UG_per_L
    Case UNITSTYPE_ENGLISH: ThisUnit = CONCENTRATION_UG_per_L
  End Select
  Call Populate_Concentration_Units(frmBubble!UnitsDesignContam(0), ThisUnit)
  Call Populate_Concentration_Units(frmBubble!UnitsDesignContam(1), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = INVERSETIME_S
    Case UNITSTYPE_ENGLISH: ThisUnit = INVERSETIME_S
  End Select
  Call Populate_InverseTime_Units(frmBubble!UnitsDesignContam(3), ThisUnit)

  'Flow Parameters.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = FLOW_M3_per_S
    Case UNITSTYPE_ENGLISH: ThisUnit = FLOW_GPM
  End Select
  Call Populate_FlowRate_Units(frmBubble!UnitsFlowParam(0), ThisUnit)
  Call Populate_FlowRate_Units(frmBubble!UnitsFlowParam(3), ThisUnit)

  'Tank Parameters.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = TIME_HR
    Case UNITSTYPE_ENGLISH: ThisUnit = TIME_HR
  End Select
  Call Populate_Time_Units(frmBubble!UnitsTankParam(1), ThisUnit)
  Call Populate_Time_Units(frmBubble!UnitsTankParam(2), ThisUnit)

  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = VOLUME_M3
    Case UNITSTYPE_ENGLISH: ThisUnit = VOLUME_FT3
  End Select
  Call Populate_Volume_Units(frmBubble!UnitsTankParam(3), ThisUnit)
  Call Populate_Volume_Units(frmBubble!UnitsTankParam(4), ThisUnit)

  'Concentration Results.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = CONCENTRATION_UG_per_L
    Case UNITSTYPE_ENGLISH: ThisUnit = CONCENTRATION_UG_per_L
  End Select
  Call Populate_Concentration_Units(frmBubble!UnitsConcResults(1), ThisUnit)
  Call Populate_Concentration_Units(frmBubble!UnitsConcResults(2), ThisUnit)
  Call Populate_Concentration_Units(frmBubble!UnitsConcResults(3), ThisUnit)






  ''Mass Transfer Parameters.
  'Call Populate_InverseTime_Units(frmPTADScreen1!UnitsMassTransfer(0), INVERSETIME_S)
  'Call Populate_InverseTime_Units(frmPTADScreen1!UnitsMassTransfer(2), INVERSETIME_S)
  '
  ''Tower Parameters.
  'Call Populate_Area_Units(frmPTADScreen1!lblTowerUnits(0), AREA_M2)
  'Call Populate_Length_Units(frmPTADScreen1!lblTowerUnits(1), LENGTH_M)
  'Call Populate_Length_Units(frmPTADScreen1!lblTowerUnits(2), LENGTH_M)
  'Call Populate_Volume_Units(frmPTADScreen1!lblTowerUnits(3), VOLUME_M3)







    'frmBubble!lblOperatingPressure.Caption = "Pressure (Pa)"
    'frmBubble!lblOperatingTemperature.Caption = "Temperature (C)"
    '
    'frmBubble!lblOxygenLabel(1).Caption = "Liquid Diffusivity (m" & Chr$(178) & "/sec)"
    'frmBubble!lblOxygenLabel(2).Caption = "Mass Transfer Coeff. (1/sec)"
    '
    'frmBubble!lblDesignConcentration(0).Caption = "Influent Concentration (" & Chr$(181) & "g/L)"
    'frmBubble!lblDesignConcentration(1).Caption = "Treatment Objective (" & Chr$(181) & "g/L)"
    'frmBubble!lblDesignConcentration(2).Caption = "Desired Percent Removal"
    'frmBubble!lblDesignConcentration(3).Caption = "Mass Transfer Coeff. (1/sec)"
    '
    'frmBubble!lblFlowParametersLabel(0).Caption = "Water Flow Rate (m" & Chr$(179) & "/sec)"
    'frmBubble!lblFlowParametersLabel(1).Caption = "Min Air to Water Ratio (m" & Chr$(179) & "/m" & Chr$(179) & ")"
    'frmBubble!lblFlowParametersLabel(2).Caption = "Air to Water Ratio (m" & Chr$(179) & "/m" & Chr$(179) & ")"
    'frmBubble!lblFlowParametersLabel(3).Caption = "Air Flow Rate (m" & Chr$(179) & "/sec)"
    '
    'frmBubble!lblTankParametersLabel(0).Caption = "No. of Tanks (in series) (-)"
    'frmBubble!lblTankParametersLabel(1).Caption = "Tank Fluid Residence Time (hr)"
    'frmBubble!lblTankParametersLabel(2).Caption = "Total Fluid Residence Time (hr)"
    'frmBubble!lblTankParametersLabel(3).Caption = "Volume of Each Tank (m" & Chr$(179) & ")"
    'frmBubble!lblTankParametersLabel(4).Caption = "Volume of All Tanks (m" & Chr$(179) & ")"
    '
    'frmBubble!lblStantonLabel.Caption = "Stanton Number (-)"
    '
    'frmBubble!lblConcentrationResultsLabel(0).Caption = "Name:"
    'frmBubble!lblConcentrationResultsLabel(1).Caption = "Ci to Tank 1 (" & Chr$(181) & "g/L)"
    'frmBubble!lblConcentrationResultsLabel(2).Caption = "Yi to All Tanks (" & Chr$(181) & "g/L)"
    'frmBubble!lblConcentrationResultsLabel(3).Caption = "Ce from Last Tank (" & Chr$(181) & "g/L)"
    'frmBubble!lblConcentrationResultsLabel(4).Caption = "Achieved Percent Removal"
    
End Sub

Sub LabelsBubbleKLaO2SI()
    
    frmOxygenMassTransferCoeff!lblDataParametersLabel(0).Caption = "SOTE (%)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(1).Caption = "SOTR (kg O2/d)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(2).Caption = "Air Flow (std. m" & Chr$(179) & "/hr)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(3).Caption = "Barometric Pres. (Pa)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(4).Caption = "Water Depth (m)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(5).Caption = "Water Volume (m" & Chr$(179) & ")"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(6).Caption = "C* (mg/L)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(7).Caption = "Apparent KLa,20 (1/s)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(8).Caption = "Phi (1/s)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(9).Caption = "True KLa,20 (1/s)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(10).Caption = "Theta (-)"
    frmOxygenMassTransferCoeff!lblDataParametersLabel(11).Caption = "KLa,O2 at Op. T (1/s)"

End Sub

Sub LabelsBubblePowerSI()

    frmBubblePower!lblPowerLabel(0).Caption = "Inlet Air Temperature (C)"
    frmBubblePower!lblPowerLabel(1).Caption = "Blower Efficiency (%)"
    frmBubblePower!lblPowerLabel(2).Caption = "Tank Water Depth (m)"
    frmBubblePower!lblPowerLabel(3).Caption = "Brake Power (kW/Blower)"
    frmBubblePower!lblPowerLabel(4).Caption = "No. of Tanks (in series)"
    frmBubblePower!lblPowerLabel(5).Caption = "No. of Blowers per Tank"
    frmBubblePower!lblPowerLabel(6).Caption = "Total Brake Power (kW)"

End Sub

Sub LabelsOptimizeContaminantsSI()

    frmOptimizeContaminant.lblOptimizationFormLabel(2).Caption = "Influent Concentration (" & Chr$(181) & "g/L)"
    frmOptimizeContaminant.lblOptimizationFormLabel(3).Caption = "Treatment Objective (" & Chr$(181) & "g/L)"
    frmOptimizeContaminant.lblOptimizationFormLabel(4).Caption = "Effluent Concentration (" & Chr$(181) & "g/L)"
    frmOptimizeContaminant.lblOptimizationFormLabel(6).Caption = "Influent Concentration (" & Chr$(181) & "g/L)"
    frmOptimizeContaminant.lblOptimizationFormLabel(7).Caption = "Treatment Objective (" & Chr$(181) & "g/L)"
    frmOptimizeContaminant.lblOptimizationFormLabel(8).Caption = "Effluent Concentration (" & Chr$(181) & "g/L)"

End Sub

Sub LabelsPowerScreen2SI()

    frmPowerScreen2!lblPowerLabel(0).Caption = "Inlet Air Temperature (C)"
    frmPowerScreen2!lblPowerLabel(1).Caption = "Blower Efficiency (%)"
    frmPowerScreen2!lblPowerLabel(2).Caption = "Blower Brake Power (kW)"
    frmPowerScreen2!lblPowerLabel(3).Caption = "Pump Efficiency (%)"
    frmPowerScreen2!lblPowerLabel(4).Caption = "Pump Brake Power (kW)"
    frmPowerScreen2!lblPowerLabel(5).Caption = "Total Brake Power (kW)"

End Sub

Sub LabelsPowerSI()

    frmPower!lblPowerLabel(0).Caption = "Inlet Air Temperature (C)"
    frmPower!lblPowerLabel(1).Caption = "Blower Efficiency (%)"
    frmPower!lblPowerLabel(2).Caption = "Blower Brake Power (kW)"
    frmPower!lblPowerLabel(3).Caption = "Pump Efficiency (%)"
    frmPower!lblPowerLabel(4).Caption = "Pump Brake Power (kW)"
    frmPower!lblPowerLabel(5).Caption = "Total Brake Power (kW)"

End Sub

Sub LabelsPropContaminant(UnitsType As Integer)
Dim ThisUnit As Integer

  ThisUnit = MOLECULAR_WEIGHT_MG_per_MMOL
  Call Populate_MolecularWeight_Units(frmContaminantPropertyEdit!UnitsProp(0), ThisUnit)

  ThisUnit = MOLAR_VOLUME_M3_per_KMOL
  Call Populate_MolarVolume_Units(frmContaminantPropertyEdit!UnitsProp(2), ThisUnit)

  ThisUnit = TEMPERATURE_C
  Call Populate_Temperature_Units(frmContaminantPropertyEdit!UnitsProp(3), ThisUnit)

  ThisUnit = DIFFUSIVITY_M2_per_S
  Call Populate_Diffusivity_Units(frmContaminantPropertyEdit!UnitsProp(4), ThisUnit)

  ThisUnit = DIFFUSIVITY_M2_per_S
  Call Populate_Diffusivity_Units(frmContaminantPropertyEdit!UnitsProp(5), ThisUnit)

  ThisUnit = CONCENTRATION_UG_per_L
  Call Populate_Concentration_Units(frmContaminantPropertyEdit!UnitsConc(0), ThisUnit)

  ThisUnit = CONCENTRATION_UG_per_L
  Call Populate_Concentration_Units(frmContaminantPropertyEdit!UnitsConc(1), ThisUnit)

'    frmPropContaminant!lblContaminantProperties(1).Caption = "Molecular Weight (kg/kmol)"
'    frmPropContaminant!lblContaminantProperties(2).Caption = "Henry's Constant (-)"
'    frmPropContaminant!lblContaminantProperties(3).Caption = "Molar Volume (m" & Chr$(179) & "/kmol)"
'    frmPropContaminant!lblContaminantProperties(4).Caption = "Normal Boiling Point (C)"
'    frmPropContaminant!lblContaminantProperties(5).Caption = "Liquid Diffusivity (m" & Chr$(178) & "/sec)"
'    frmPropContaminant!lblContaminantProperties(6).Caption = "Gas Diffusivity (m" & Chr$(178) & "/sec)"
'    frmPropContaminant!lblContaminantProperties(7).Caption = "Influent Conc. (" & Chr$(181) & "g/L)"
'    frmPropContaminant!lblContaminantProperties(8).Caption = "Treatment Obj. (" & Chr$(181) & "g/L)"

End Sub

Sub LabelsPropContaminantBubbleSI()

    'frmPropContaminantBubble!lblContaminantProperties(1).Caption = "Molecular Weight (kg/kmol)"
    'frmPropContaminantBubble!lblContaminantProperties(2).Caption = "Henry's Constant (-)"
    'frmPropContaminantBubble!lblContaminantProperties(3).Caption = "Molar Volume (m" & Chr$(179) & "/kmol)"
    'frmPropContaminantBubble!lblContaminantProperties(5).Caption = "Liquid Diffusivity (m" & Chr$(178) & "/sec)"
    'frmPropContaminantBubble!lblContaminantProperties(7).Caption = "Influent Conc. (" & Chr$(181) & "g/L)"
    'frmPropContaminantBubble!lblContaminantProperties(8).Caption = "Treatment Obj. (" & Chr$(181) & "g/L)"
    'frmPropContaminantBubble!lblContaminantProperties(9).Caption = "Pressure (Pa)"
    'frmPropContaminantBubble!lblContaminantProperties(10).Caption = "Temperature (C)"

End Sub

Sub LabelsPropContaminantSurfaceSI()

    'frmPropContaminantSurface!lblContaminantProperties(1).Caption = "Molecular Weight (kg/kmol)"
    'frmPropContaminantSurface!lblContaminantProperties(2).Caption = "Henry's Constant (-)"
    'frmPropContaminantSurface!lblContaminantProperties(3).Caption = "Molar Volume (m" & Chr$(179) & "/kmol)"
    'frmPropContaminantSurface!lblContaminantProperties(5).Caption = "Liquid Diffusivity (m" & Chr$(178) & "/sec)"
    'frmPropContaminantSurface!lblContaminantProperties(7).Caption = "Influent Conc. (" & Chr$(181) & "g/L)"
    'frmPropContaminantSurface!lblContaminantProperties(8).Caption = "Treatment Obj. (" & Chr$(181) & "g/L)"
    'frmPropContaminantSurface!lblContaminantProperties(9).Caption = "Pressure (Pa)"
    'frmPropContaminantSurface!lblContaminantProperties(10).Caption = "Temperature (C)"

End Sub

Sub LabelsPTADScreen1(UnitsType As Integer)
Dim ThisUnit As Integer

  'Operating Conditions.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = PRESSURE_PA
    Case UNITSTYPE_ENGLISH: ThisUnit = PRESSURE_ATM
  End Select
  Call Populate_Pressure_Units(frmptadscreen1!txtPUnits, ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = TEMPERATURE_C
    Case UNITSTYPE_ENGLISH: ThisUnit = TEMPERATURE_F
  End Select
  Call Populate_Temperature_Units(frmptadscreen1!txtTUnits, ThisUnit)
  
  'Flows and Loadings.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = FLOW_M3_per_S
    Case UNITSTYPE_ENGLISH: ThisUnit = FLOW_GPM
  End Select
  Call Populate_FlowRate_Units(frmptadscreen1!txtFlowsUnits(0), ThisUnit)
  Call Populate_FlowRate_Units(frmptadscreen1!txtFlowsUnits(4), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = PRESSUREPERLENGTH_PA_per_M
    Case UNITSTYPE_ENGLISH: ThisUnit = PRESSUREPERLENGTH_PSI_per_FT
  End Select
  Call Populate_PressurePerLength_Units(frmptadscreen1!txtFlowsUnits(5), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = MASSLOADINGRATE_KG_M2_S
    Case UNITSTYPE_ENGLISH: ThisUnit = MASSLOADINGRATE_LBM_FT2_S
  End Select
  Call Populate_MassLoadingRate_Units(frmptadscreen1!lblFlowsUnits(6), ThisUnit)
  Call Populate_MassLoadingRate_Units(frmptadscreen1!lblFlowsUnits(7), ThisUnit)

  'Mass Transfer Parameters.
  ThisUnit = INVERSETIME_S
  Call Populate_InverseTime_Units(frmptadscreen1!UnitsMassTransfer(0), ThisUnit)
  Call Populate_InverseTime_Units(frmptadscreen1!UnitsMassTransfer(2), ThisUnit)

  'Tower Parameters.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = AREA_M2
    Case UNITSTYPE_ENGLISH: ThisUnit = AREA_FT2
  End Select
  Call Populate_Area_Units(frmptadscreen1!lblTowerUnits(0), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = LENGTH_M
    Case UNITSTYPE_ENGLISH: ThisUnit = LENGTH_FT
  End Select
  Call Populate_Length_Units(frmptadscreen1!lblTowerUnits(1), ThisUnit)
  Call Populate_Length_Units(frmptadscreen1!lblTowerUnits(2), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = VOLUME_M3
    Case UNITSTYPE_ENGLISH: ThisUnit = VOLUME_FT3
  End Select
  Call Populate_Volume_Units(frmptadscreen1!lblTowerUnits(3), ThisUnit)

End Sub

Sub LabelsPTADScreen2(UnitsType As Integer)
Dim ThisUnit As Integer

  'Design Based On. ===========================================================
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = LENGTH_M
    Case UNITSTYPE_ENGLISH: ThisUnit = LENGTH_FT
  End Select
  Call Populate_Length_Units(frmPTADScreen2!UnitsDesignBasis(0), ThisUnit)
  Call Populate_Length_Units(frmPTADScreen2!UnitsDesignBasis(1), ThisUnit)
  
  
  'Tower Parameters. ==========================================================
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = LENGTH_M
    Case UNITSTYPE_ENGLISH: ThisUnit = LENGTH_FT
  End Select
  Call Populate_Length_Units(frmPTADScreen2!UnitsTowerParam(0), ThisUnit)
  Call Populate_Length_Units(frmPTADScreen2!UnitsTowerParam(1), ThisUnit)

  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = AREA_M2
    Case UNITSTYPE_ENGLISH: ThisUnit = AREA_FT2
  End Select
  Call Populate_Area_Units(frmPTADScreen2!UnitsTowerParam(2), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = VOLUME_M3
    Case UNITSTYPE_ENGLISH: ThisUnit = VOLUME_FT3
  End Select
  Call Populate_Volume_Units(frmPTADScreen2!UnitsTowerParam(3), ThisUnit)


  'Operating Conditions. ======================================================
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = PRESSURE_PA
    Case UNITSTYPE_ENGLISH: ThisUnit = PRESSURE_ATM
  End Select
  Call Populate_Pressure_Units(frmPTADScreen2!UnitsOpCond(0), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = TEMPERATURE_C
    Case UNITSTYPE_ENGLISH: ThisUnit = TEMPERATURE_F
  End Select
  Call Populate_Temperature_Units(frmPTADScreen2!UnitsOpCond(1), ThisUnit)
  
  
  'Flows and Loadings. ========================================================
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = FLOW_M3_per_S
    Case UNITSTYPE_ENGLISH: ThisUnit = FLOW_GPM
  End Select
  Call Populate_FlowRate_Units(frmPTADScreen2!UnitsFlows(0), ThisUnit)
  Call Populate_FlowRate_Units(frmPTADScreen2!UnitsFlows(1), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = MASSLOADINGRATE_KG_M2_S
    Case UNITSTYPE_ENGLISH: ThisUnit = MASSLOADINGRATE_LBM_FT2_S
  End Select
  Call Populate_MassLoadingRate_Units(frmPTADScreen2!UnitsFlows(3), ThisUnit)
  Call Populate_MassLoadingRate_Units(frmPTADScreen2!UnitsFlows(4), ThisUnit)
  
  
  'Contaminant of Interest. ===================================================
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = INVERSETIME_S
    Case UNITSTYPE_ENGLISH: ThisUnit = INVERSETIME_S
  End Select
  Call Populate_InverseTime_Units(frmPTADScreen2!UnitsInterest(0), ThisUnit)
  Call Populate_InverseTime_Units(frmPTADScreen2!UnitsInterest(2), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = CONCENTRATION_UG_per_L
    Case UNITSTYPE_ENGLISH: ThisUnit = CONCENTRATION_UG_per_L
  End Select
  Call Populate_Concentration_Units(frmPTADScreen2!UnitsInterest(3), ThisUnit)
  Call Populate_Concentration_Units(frmPTADScreen2!UnitsInterest(4), ThisUnit)
  Call Populate_Concentration_Units(frmPTADScreen2!UnitsInterest(5), ThisUnit)
         
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = PRESSUREPERLENGTH_PA_per_M
    Case UNITSTYPE_ENGLISH: ThisUnit = PRESSUREPERLENGTH_PSI_per_FT
  End Select
  Call Populate_PressurePerLength_Units(frmPTADScreen2!UnitsInterest(7), ThisUnit)


  

  '
  ''Flows and Loadings.
  'Call Populate_Flowrate_Units(frmPTADScreen1!txtFlowsUnits(0), FLOW_M3_per_S)
  'Call Populate_Flowrate_Units(frmPTADScreen1!txtFlowsUnits(4), FLOW_M3_per_S)
  'Call Populate_PressurePerLength_Units(frmPTADScreen1!txtFlowsUnits(5), PRESSUREPERLENGTH_PA_per_M)
  'Call Populate_MassLoadingRate_Units(frmPTADScreen1!lblFlowsUnits(6), MASSLOADINGRATE_KG_M2_S)
  'Call Populate_MassLoadingRate_Units(frmPTADScreen1!lblFlowsUnits(7), MASSLOADINGRATE_KG_M2_S)
  '
  ''Mass Transfer Parameters.
  'Call Populate_InverseTime_Units(frmPTADScreen1!lblTransferUnits(0), INVERSETIME_S)
  'Call Populate_InverseTime_Units(frmPTADScreen1!txtTransferUnits(2), INVERSETIME_S)
  '
  ''Tower Parameters.
  'Call Populate_Area_Units(frmPTADScreen1!lblTowerUnits(0), AREA_M2)
  'Call Populate_Length_Units(frmPTADScreen1!lblTowerUnits(1), LENGTH_M)
  'Call Populate_Length_Units(frmPTADScreen1!lblTowerUnits(2), LENGTH_M)
  'Call Populate_Volume_Units(frmPTADScreen1!lblTowerUnits(3), VOLUME_M3)
    


    'frmPTADScreen2!lblDesignParametersLabel(0).Caption = "Tower Diameter (m)"
    'frmPTADScreen2!lblDesignParametersLabel(1).Caption = "Tower Height (m)"
    'frmPTADScreen2!lblTowerParametersLabel(0).Caption = "Specify Tower Diameter (m)"
    'frmPTADScreen2!lblTowerParametersLabel(1).Caption = "Specify Tower Height (m)"
    'frmPTADScreen2!lblTowerParametersLabel(2).Caption = "Tower Area (m" & Chr$(178) & ")"
    'frmPTADScreen2!lblTowerParametersLabel(3).Caption = "Tower Volume (m" & Chr$(179) & ")"
    'frmPTADScreen2!lblOperatingPressure.Caption = "Pressure (Pa)"
    'frmPTADScreen2!lblOperatingTemperature.Caption = "Temperature (C)"
    'frmPTADScreen2!lblFlowsLoadingsLabel(0).Caption = "Water Flow Rate (m" & Chr$(179) & "/sec)"
    'frmPTADScreen2!lblFlowsLoadingsLabel(1).Caption = "Air Flow Rate (m" & Chr$(179) & "/sec)"
    'frmPTADScreen2!lblFlowsLoadingsLabel(2).Caption = "Air to Water Ratio (m" & Chr$(179) & "/m" & Chr$(179) & ")"
    'frmPTADScreen2!lblFlowsLoadingsLabel(3).Caption = "Water Loading Rate (kg/m" & Chr$(178) & "/s)"
    'frmPTADScreen2!lblFlowsLoadingsLabel(4).Caption = "Air Loading Rate (kg/m" & Chr$(178) & "/s)"
    'frmPTADScreen2!lblDesignConcentration(0).Caption = "Onda KLa (1/sec)"
    'frmPTADScreen2!lblDesignConcentration(1).Caption = "KLa Safety Factor (-)"
    'frmPTADScreen2!lblDesignConcentration(2).Caption = "Design KLa (1/sec)"
    'frmPTADScreen2!lblDesignConcentration(3).Caption = "Influent Concentration (" & Chr$(181) & "g/L)"
    'frmPTADScreen2!lblDesignConcentration(4).Caption = "Treatment Objective (" & Chr$(181) & "g/L)"
    'frmPTADScreen2!lblDesignConcentration(5).Caption = "Effluent Concentration (" & Chr$(181) & "g/L)"
    'frmPTADScreen2!lblDesignConcentration(6).Caption = "Percent Removal"
    'frmPTADScreen2!lblDesignConcentration(7).Caption = "Air Pressure Drop (N/m" & Chr$(178) & "/m)"

End Sub

Sub LabelsSelectPackingPropertiesSI()

    frmSelectPacking!lblPackingProperties(1).Caption = "Nominal Size (m)"
    frmSelectPacking!lblPackingProperties(2).Caption = "Packing Factor (-)"
    frmSelectPacking!lblPackingProperties(3).Caption = "Sp. Surf. Area (m" & Chr$(178) & "/m" & Chr$(179) & ")"
    frmSelectPacking!lblPackingProperties(4).Caption = "Crit. Surf. Tension (N/m)"

End Sub

Sub LabelsShowOndaKLaSI()

    frmShowOndaKLaProperties!lblOndaPropertiesLabel(0).Caption = "Reynold's Number (-)"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(1).Caption = "Froude Number (-)"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(2).Caption = "Weber Number (-)"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(3).Caption = "Packing Wetted Surf. Area (m" & Chr$(178) & "/m" & Chr$(179) & ")"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(4).Caption = "Liq. Phase M. T. Resistance (sec)"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(5).Caption = "Gas Phase M. T. Resistance (sec)"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(6).Caption = "Total Mass Transfer Resistance (sec)"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(7).Caption = "Liq. Phase M. T. Coefficient (m/sec)"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(8).Caption = "Gas Phase M. T. Coefficient (m/sec)"
    frmShowOndaKLaProperties!lblOndaPropertiesLabel(9).Caption = "Overall M. T. Coeff. = Onda KLa (1/s)"

End Sub

Sub LabelsShowPackingPropertiesSI()

    frmShowPackingProperties!lblShowPackingPropertesLabel(1).Caption = "Nominal Size (m)"
    frmShowPackingProperties!lblShowPackingPropertesLabel(2).Caption = "Packing Factor (-)"
    frmShowPackingProperties!lblShowPackingPropertesLabel(3).Caption = "Sp. Surf. Area (m" & Chr$(178) & "/m" & Chr$(179) & ")"
    frmShowPackingProperties!lblShowPackingPropertesLabel(4).Caption = "Critical Surface Tension (N/m)"

End Sub

Sub LabelsSurface(UnitsType As Integer)
Dim ThisUnit As Integer

  'Operating Conditions.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = PRESSURE_PA
    Case UNITSTYPE_ENGLISH: ThisUnit = PRESSURE_ATM
  End Select
  Call Populate_Pressure_Units(frmsurface!UnitsOpCond(0), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = TEMPERATURE_C
    Case UNITSTYPE_ENGLISH: ThisUnit = TEMPERATURE_F
  End Select
  Call Populate_Temperature_Units(frmsurface!UnitsOpCond(1), ThisUnit)
  
  'Power Input, P/V.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = POWERPERVOLUME_W_per_M3
    Case UNITSTYPE_ENGLISH: ThisUnit = POWERPERVOLUME_HP_per_FT3
  End Select
  Call Populate_PowerPerVolume_Units(frmsurface!UnitsPowerInput, ThisUnit)
  
  'Oxygen (reference compound).
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = DIFFUSIVITY_M2_per_S
    Case UNITSTYPE_ENGLISH: ThisUnit = DIFFUSIVITY_FT2_per_S
  End Select
  Call Populate_Diffusivity_Units(frmsurface!UnitsOxygenRef(1), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = INVERSETIME_S
    Case UNITSTYPE_ENGLISH: ThisUnit = INVERSETIME_S
  End Select
  Call Populate_InverseTime_Units(frmsurface!UnitsOxygenRef(2), ThisUnit)
  
  'Design Contaminant.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = CONCENTRATION_UG_per_L
    Case UNITSTYPE_ENGLISH: ThisUnit = CONCENTRATION_UG_per_L
  End Select
  Call Populate_Concentration_Units(frmsurface!UnitsDesignContam(0), ThisUnit)
  Call Populate_Concentration_Units(frmsurface!UnitsDesignContam(1), ThisUnit)
  
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = INVERSETIME_S
    Case UNITSTYPE_ENGLISH: ThisUnit = INVERSETIME_S
  End Select
  Call Populate_InverseTime_Units(frmsurface!UnitsDesignContam(3), ThisUnit)
  
  'Flow Parameters.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = FLOW_M3_per_S
    Case UNITSTYPE_ENGLISH: ThisUnit = FLOW_GPM
  End Select
  Call Populate_FlowRate_Units(frmsurface!UnitsFlowParam(0), ThisUnit)
  
  'Tank Parameters.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = TIME_HR
    Case UNITSTYPE_ENGLISH: ThisUnit = TIME_HR
  End Select
  Call Populate_Time_Units(frmsurface!UnitsTankParam(1), ThisUnit)
  Call Populate_Time_Units(frmsurface!UnitsTankParam(2), ThisUnit)

  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = VOLUME_M3
    Case UNITSTYPE_ENGLISH: ThisUnit = VOLUME_FT3
  End Select
  Call Populate_Volume_Units(frmsurface!UnitsTankParam(3), ThisUnit)
  Call Populate_Volume_Units(frmsurface!UnitsTankParam(4), ThisUnit)

  'Concentration Results.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = CONCENTRATION_UG_per_L
    Case UNITSTYPE_ENGLISH: ThisUnit = CONCENTRATION_UG_per_L
  End Select
  Call Populate_Concentration_Units(frmsurface!UnitsConcResults(1), ThisUnit)
  Call Populate_Concentration_Units(frmsurface!UnitsConcResults(3), ThisUnit)
  
  'Power Calculation.
  Select Case UnitsType
    Case UNITSTYPE_SI: ThisUnit = POWER_KW
    Case UNITSTYPE_ENGLISH: ThisUnit = POWER_HP
  End Select
  Call Populate_Power_Units(frmsurface!UnitsPowerCalc(1), ThisUnit)
  Call Populate_Power_Units(frmsurface!UnitsPowerCalc(2), ThisUnit)




    ''for main surface form
    '
    'frmSurface!lblOperatingPressure.Caption = "Pressure (Pa)"
    'frmSurface!lblOperatingTemperature.Caption = "Temperature (C)"
    '
    'frmSurface!lblPowerInputLabel.Caption = "Power Input, P/V (W/m" & Chr$(179) & ")"
    '
    'frmSurface!lblOxygenLabel(1).Caption = "Liquid Diffusivity (m" & Chr$(178) & "/sec)"
    'frmSurface!lblOxygenLabel(2).Caption = "Mass Transfer Coeff. (1/sec)"
    '
    'frmSurface!lblDesignConcentration(0).Caption = "Influent Concentration (" & Chr$(181) & "g/L)"
    'frmSurface!lblDesignConcentration(1).Caption = "Treatment Objective (" & Chr$(181) & "g/L)"
    'frmSurface!lblDesignConcentration(2).Caption = "Desired Percent Removal"
    'frmSurface!lblDesignConcentration(3).Caption = "Mass Transfer Coeff. (1/sec)"
    '
    'frmSurface!lblFlowParametersLabel(0).Caption = "Water Flow Rate (m" & Chr$(179) & "/sec)"
    '
    'frmSurface!lblTankParametersLabel(0).Caption = "No. of Tanks (in series) (-)"
    'frmSurface!lblTankParametersLabel(1).Caption = "Tank Fluid Residence Time (hr)"
    'frmSurface!lblTankParametersLabel(2).Caption = "Total Fluid Residence Time (hr)"
    'frmSurface!lblTankParametersLabel(3).Caption = "Volume of Each Tank (m" & Chr$(179) & ")"
    'frmSurface!lblTankParametersLabel(4).Caption = "Volume of All Tanks (m" & Chr$(179) & ")"
    '
    'frmSurface!lblConcentrationResultsLabel(0).Caption = "Name:"
    'frmSurface!lblConcentrationResultsLabel(1).Caption = "Ci to Tank 1 (" & Chr$(181) & "g/L)"
    'frmSurface!lblConcentrationResultsLabel(3).Caption = "Ce from Last Tank (" & Chr$(181) & "g/L)"
    'frmSurface!lblConcentrationResultsLabel(4).Caption = "Achieved Percent Removal"
    '
    'frmSurface!lblPowerCalculationLabel(0).Caption = "Aerator Motor Efficiency (%)"
    'frmSurface!lblPowerCalculationLabel(1).Caption = "Power Required per Tank (kW)"
    'frmSurface!lblPowerCalculationLabel(2).Caption = "Total Power Required (kW)"
    
End Sub

Sub LabelsWaterPropertiesBubbleSI()

    frmWaterPropertiesBubble!lblAirWaterProperties(0).Caption = "Water Density (kg/m" & Chr$(179) & ")"
    frmWaterPropertiesBubble!lblAirWaterProperties(1).Caption = "Water Viscosity (kg/m/sec)"

End Sub

Sub LabelsWaterPropertiesSurfaceSI()

    frmWaterPropertiesSurface!lblAirWaterProperties(0).Caption = "Water Density (kg/m" & Chr$(179) & ")"
    frmWaterPropertiesSurface!lblAirWaterProperties(1).Caption = "Water Viscosity (kg/m/sec)"

End Sub

Sub xxxLabelsPropContaminantScreen2SI()

    'frmPropContaminantScreen2!lblContaminantProperties(1).Caption = "Molecular Weight (kg/kmol)"
    'frmPropContaminantScreen2!lblContaminantProperties(2).Caption = "Henry's Constant (-)"
    'frmPropContaminantScreen2!lblContaminantProperties(3).Caption = "Molar Volume (m" & Chr$(179) & "/kmol)"
    'frmPropContaminantScreen2!lblContaminantProperties(4).Caption = "Normal Boiling Point (C)"
    'frmPropContaminantScreen2!lblContaminantProperties(5).Caption = "Liquid Diffusivity (m" & Chr$(178) & "/sec)"
    'frmPropContaminantScreen2!lblContaminantProperties(6).Caption = "Gas Diffusivity (m" & Chr$(178) & "/sec)"
    'frmPropContaminantScreen2!lblContaminantProperties(7).Caption = "Influent Conc. (" & Chr$(181) & "g/L)"
    'frmPropContaminantScreen2!lblContaminantProperties(8).Caption = "Treatment Obj. (" & Chr$(181) & "g/L)"
    'frmPropContaminantScreen2!lblContaminantProperties(9).Caption = "Pressure (Pa)"
    'frmPropContaminantScreen2!lblContaminantProperties(10).Caption = "Temperature (C)"

End Sub

