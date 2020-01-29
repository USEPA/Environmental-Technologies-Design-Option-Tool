Attribute VB_Name = "StructsDo"
Option Explicit



Const StructsDo_declarations_end = True


Sub DataSrc_Set( _
    Pp As TYPE_PlantDiagram, _
    idx_This As Integer, _
    SelType As String, _
    in_Val_UserInput As Double, _
    in_Val_StEPP As Double, _
    in_Val_Corr As Double _
    )
  With Pp.ChemicalData.DataSources(idx_This)
    Select Case Trim$(UCase$(SelType))
      Case UCase$("U"): .SourceType = DATASOURCETYPE_USERINPUT
      Case UCase$("S"): .SourceType = DATASOURCETYPE_STEPP
      Case UCase$("C"): .SourceType = DATASOURCETYPE_CORR
    End Select
    .Val_UserInput = in_Val_UserInput
    .Val_StEPP = in_Val_StEPP
    .Val_Corr = in_Val_Corr
  End With
End Sub
Sub Project_Plant_SetDefaults(Pp As TYPE_PlantDiagram)
Dim i As Integer
Dim ThisVal As Double
  With Pp
    '
    ' FORM frmMain.
    '
    .en_InfluentWeir = True
    .en_GritChamber = True
    .en_PrimaryWeir = True
    .en_SecondaryWeir = True
    .Flow = 190500#
    .SolidsConc = 1#
  End With
  With Pp.ChemicalData
    '
    ' FORM frmD0_Props.
    '
    .env_Pressure = 96.5
    .env_Temperature = 20#
    .env_WindVelocity = 2#
    .ContaminantName = "(Unnamed)"
    .InfluentConc = 10#
    .BiodegredationRate = 0.002
    .LogKow = 1.97
    .VOC_HenrysConstant = 0.172
    .VOC_MolecularWeight = 119.5
    .VOC_DiffusivityInH2O = 0.0000094
    .VOC_DiffusivityInGas = 0.07
    '.O2_SaturationConc = 9.17
    '.O2_HenrysConstant = 30#
    '.O2_Diffusivity = 0.0000242
    '.H2O_Density = 999#
    '.H2O_Viscosity = 0.001005
    '.H2O_VaporPressure = 1.6
    .H2O_Alpha = 0.7
    '.AIR_Density = 1.21
    '.AIR_Viscosity = 0.00017
    .O2_CInfinity = 11#
    Call DataSrc_Set(Pp, 0, "U", .env_Pressure, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 1, "U", .env_Temperature, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 2, "U", .env_WindVelocity, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 3, "U", .InfluentConc, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 4, "U", .BiodegredationRate, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 5, "U", .LogKow, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 6, "U", .VOC_HenrysConstant, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 7, "U", .VOC_MolecularWeight, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 8, "U", .VOC_DiffusivityInH2O, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 9, "U", .VOC_DiffusivityInGas, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 10, "C", .O2_SaturationConc, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 11, "C", .O2_HenrysConstant, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 12, "C", .O2_Diffusivity, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 13, "C", .H2O_Density, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 14, "C", .H2O_Viscosity, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 15, "C", .H2O_VaporPressure, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 16, "U", .H2O_Alpha, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 17, "C", .AIR_Density, -1E+20, -1E+20)
    Call DataSrc_Set(Pp, 18, "C", .AIR_Viscosity, -1E+20, -1E+20)
    .UnitsOfDisplay(0) = "kPa"
    .UnitsOfDisplay(1) = "C"
    .UnitsOfDisplay(2) = "m/s"
    .UnitsOfDisplay(3) = "µg/L"
    .UnitsOfDisplay(4) = "L/mg-d"
    .UnitsOfDisplay(5) = "N/A"
    .UnitsOfDisplay(6) = "N/A"
    .UnitsOfDisplay(7) = "g/gmol"
    .UnitsOfDisplay(8) = "cm²/s"
    .UnitsOfDisplay(9) = "cm²/s"
    .UnitsOfDisplay(10) = "mg/L"
    .UnitsOfDisplay(11) = "N/A"
    .UnitsOfDisplay(12) = "cm²/s"
    .UnitsOfDisplay(13) = "kg/m³"
    .UnitsOfDisplay(14) = "kg/m-s"
    .UnitsOfDisplay(15) = "kPa"
    .UnitsOfDisplay(16) = "N/A"
    .UnitsOfDisplay(17) = "kg/m³"
    .UnitsOfDisplay(18) = "kg/m-s"
    '
    ' SET .Val_Corr FOR THE WATER AND AIR CORRELATIONS.
    '
    Call Corr_SetWaterAndAirAndOxygen(Pp)
    '
    ' FOR THIS INITIALIZATION (_SetDefaults) CODE ONLY,
    ' TRANSFER THE .Val_Corr VALUES INTO THE .Val_UserInput VALUES.
    '
    For i = 10 To 18
      If (i <> 16) Then
        .DataSources(i).Val_UserInput = .DataSources(i).Val_Corr
      End If
    Next i
    .O2_SaturationConc = .DataSources(10).Val_Corr
    .O2_HenrysConstant = .DataSources(11).Val_Corr
    .O2_Diffusivity = .DataSources(12).Val_Corr
    .H2O_Density = .DataSources(13).Val_Corr
    .H2O_Viscosity = .DataSources(14).Val_Corr
    .H2O_VaporPressure = .DataSources(15).Val_Corr
    '.H2O_Alpha = 0.7
    .AIR_Density = .DataSources(17).Val_Corr
    .AIR_Viscosity = .DataSources(18).Val_Corr
  End With
  With Pp
    '
    ' FORM frmD1_InfluentWeir.
    '
    .InfluentWeir.ModelingMechanism = WEIR_MODEL_TYPE_POOL    '1
    .InfluentWeir.Width = 25#
    .InfluentWeir.WaterLevelDiff = 0.5
    .InfluentWeir.GasFlow = 0.6
    .InfluentWeir.UnitsOfDisplay(0) = "m"
    .InfluentWeir.UnitsOfDisplay(1) = "m"
    ''''.InfluentWeir.UnitsOfDisplay(2) = "m³/m-h"
    .InfluentWeir.UnitsOfDisplay(2) = "m³/(m-h)"
    '
    ' FORM frmD2_GritChamber.
    '
    .GritChamber.IsCovered = True
    .GritChamber.Count = 1
    .GritChamber.VentilationRate = 100#
    .GritChamber.Depth = 4#
    ''''.GritChamber.GasFlow = 20000#
    .GritChamber.GasFlow = 20000# * 1000#
    .GritChamber.Volume = 83000#
    .GritChamber.SOTR = 150#
    .GritChamber.UnitsOfDisplay(1) = "L/min"
    .GritChamber.UnitsOfDisplay(2) = "m"
    .GritChamber.UnitsOfDisplay(3) = "liter"
    ''''.GritChamber.UnitsOfDisplay(4) = "m³/m-h"
    .GritChamber.UnitsOfDisplay(4) = "m³/min"
    .GritChamber.UnitsOfDisplay(5) = "kg/d"
    '
    ' FORM frmD3_PrimaryClarifier.
    '
    .PrimaryClarifier.IsCovered = True
    .PrimaryClarifier.Count = 1
    .PrimaryClarifier.SorptionRemovalMethod = 0
    .PrimaryClarifier.VolatilizationRemovalMechanism = 1
    .PrimaryClarifier.VentilationRate = 100#
    .PrimaryClarifier.Depth = 4#
    .PrimaryClarifier.Volume = 300000#
    .PrimaryClarifier.WastageFlow = 4700#
    .PrimaryClarifier.PercentageRemoval = 100#      'PRIMARY CLARIFIER ONLY.
    .PrimaryClarifier.EffluentSolidsConc = 0#       'SECONDARY CLARIFIER ONLY.
    .PrimaryClarifier.UnitsOfDisplay(1) = "L/min"
    .PrimaryClarifier.UnitsOfDisplay(2) = "m"
    .PrimaryClarifier.UnitsOfDisplay(3) = "liter"
    .PrimaryClarifier.UnitsOfDisplay(4) = "L/d"
    .PrimaryClarifier.UnitsOfDisplay(5) = "N/A"     'SECONDARY CLARIFIER ONLY.
    '
    ' FORM frmD4_PrimaryWeir.
    '
    .PrimaryWeir.ModelingMechanism = WEIR_MODEL_TYPE_NAPPE   '0
    .PrimaryWeir.Width = 3#
    .PrimaryWeir.WaterLevelDiff = 0.5
    .PrimaryWeir.GasFlow = 0.6
    .PrimaryWeir.UnitsOfDisplay(0) = "m"
    .PrimaryWeir.UnitsOfDisplay(1) = "m"
    ''''.PrimaryWeir.UnitsOfDisplay(2) = "m³/m-h"
    .PrimaryWeir.UnitsOfDisplay(2) = "m³/(m-h)"
    '
    ' FORM frmD5_AerationBasin.
    '
    .AerationBasin.IsCovered = True
    .AerationBasin.Count = 1
    .AerationBasin.ModelingMechanism = 1
    .AerationBasin.AutoCalcBioMass = False
    .AerationBasin.VentilationRate = 100#
    .AerationBasin.Depth = 4#
    .AerationBasin.WastageFlow = 800#
    .AerationBasin.RecycleFlow = 10000#
    .AerationBasin.SOTR = 8500#
    .AerationBasin.SolidsConcInRecycle = 100#
    .AerationBasin.Volume = 800000#
    .AerationBasin.GasFlow = 5600#
    .AerationBasin.BioMass = 10000#
    .AerationBasin.UnitsOfDisplay(1) = "L/min"
    .AerationBasin.UnitsOfDisplay(2) = "m"
    .AerationBasin.UnitsOfDisplay(3) = "L/d"
    .AerationBasin.UnitsOfDisplay(4) = "L/d"
    .AerationBasin.UnitsOfDisplay(5) = "kg/hr"
    
    .AerationBasin.CSTR.Count = 1
    .AerationBasin.CSTR.UseStepFeed = False
    .AerationBasin.CSTR.UniformFeed = True
    .AerationBasin.CSTR.UniformVolume = True
    .AerationBasin.CSTR.UniformGasFlow = True
    .AerationBasin.CSTR.UniformBioMass = True
    For i = 0 To 8
      .AerationBasin.CSTR.Feed(i) = 0#
      .AerationBasin.CSTR.Volume(i) = 0#
      .AerationBasin.CSTR.GasFlow(i) = 0#
      .AerationBasin.CSTR.BioMass(i) = 0#
    Next i
    '
    ' FORM frmD5B_Biomass.
    '
    .AerationBasin.BioTreat.MaxGrowthRate = 3#
    .AerationBasin.BioTreat.HalfVelocityConst = 60#
    .AerationBasin.BioTreat.BacterialDecay = 0.006
    .AerationBasin.BioTreat.YieldCoeff = 6#
    .AerationBasin.BioTreat.BOD5Conc = 84#
    .AerationBasin.BioTreat.UnitsOfDisplay(0) = "1/day"
    .AerationBasin.BioTreat.UnitsOfDisplay(1) = "N/A"
    .AerationBasin.BioTreat.UnitsOfDisplay(2) = "1/day"
    .AerationBasin.BioTreat.UnitsOfDisplay(3) = "N/A"
    .AerationBasin.BioTreat.UnitsOfDisplay(4) = "mg/L"
    '
    ' FORM frmD6_SecondaryClarifier.
    '
    .SecondaryClarifier.IsCovered = True
    .SecondaryClarifier.Count = 1
    .SecondaryClarifier.VentilationRate = 100#
    .SecondaryClarifier.Depth = 4#
    .SecondaryClarifier.Volume = 350000#
    .SecondaryClarifier.PercentageRemoval = 0#        'PRIMARY CLARIFIER ONLY.
    .SecondaryClarifier.EffluentSolidsConc = 0#       'SECONDARY CLARIFIER ONLY.
    .SecondaryClarifier.UnitsOfDisplay(1) = "L/min"
    .SecondaryClarifier.UnitsOfDisplay(2) = "m"
    .SecondaryClarifier.UnitsOfDisplay(3) = "liter"
    .SecondaryClarifier.UnitsOfDisplay(4) = "N/A"
    .SecondaryClarifier.UnitsOfDisplay(5) = "mg/L"    'SECONDARY CLARIFIER ONLY.
    '
    ' FORM frmD7_SecondaryWeir.
    '
    .SecondaryWeir.ModelingMechanism = WEIR_MODEL_TYPE_POOL   '1
    .SecondaryWeir.Width = 25#
    .SecondaryWeir.WaterLevelDiff = 0.5
    .SecondaryWeir.GasFlow = 0.6
    .SecondaryWeir.UnitsOfDisplay(0) = "m"
    .SecondaryWeir.UnitsOfDisplay(1) = "m"
    ''''.SecondaryWeir.UnitsOfDisplay(2) = "m³/m-h"
    .SecondaryWeir.UnitsOfDisplay(2) = "m³/(m-h)"
  End With
End Sub


Sub Project_OutputRec_SetDefaults(ORe As TYPE_OutputRecord)
  With ORe
    .IsDisplayed = False
    
    .TotalAmount.pr_Stripping = 0#        ''OK
    .TotalAmount.pr_Volatilization = 0#   ''Added to pr_stripping
    .TotalAmount.pr_SolidWaste = 0#       ''OK
    .TotalAmount.pr_LiquidWaste = 0#      ''Added to pr_solidwaste
    .TotalAmount.pr_Biodegredation = 0#   ''OK
    .pr_TotalRemoved = 0#
    
    .TotalAmount.Stripping = 0#           ''OK
    .TotalAmount.Volatilization = 0#      ''Added to stripping
    .TotalAmount.SolidWaste = 0#          ''OK
    .TotalAmount.LiquidWaste = 0#         ''Added to SolidWaste
    .TotalAmount.Biodegredation = 0#      ''OK
    .TotalInfluent = 0#
    .TotalEffluent = 0#
    
    ' TAKE CARE OF THE AMOUNT VARIABLES
    .InfluentWeir.EffluentConc = 0#
    .InfluentWeir.Volatilization = 0#
    .InfluentWeir.Stripping = 0#
    .InfluentWeir.SolidWaste = 0#
    .InfluentWeir.LiquidWaste = 0#
    .InfluentWeir.Biodegredation = 0#
 
    .GritChamber.EffluentConc = 0#
    .GritChamber.Stripping = 0#
    .GritChamber.Volatilization = 0#
    
    .PrimaryClarifier.EffluentConc = 0#
    .PrimaryClarifier.Stripping = 0#
    .PrimaryClarifier.Volatilization = 0#
    .PrimaryClarifier.SolidWaste = 0#
    .PrimaryClarifier.LiquidWaste = 0#
    
    .PrimaryWeir.EffluentConc = 0#
    .PrimaryWeir.Stripping = 0#
    
    .AerationBasin.EffluentConc = 0#
    .AerationBasin.Stripping = 0#
    .AerationBasin.Volatilization = 0#
    .AerationBasin.Biodegredation = 0#
    
    .SecondaryClarifier.EffluentConc = 0#
    .SecondaryClarifier.Volatilization = 0#
    .SecondaryClarifier.SolidWaste = 0#
    .SecondaryClarifier.LiquidWaste = 0#
    
    .SecondaryWeir.EffluentConc = 0#
    .SecondaryWeir.Stripping = 0#
    
    
    ' TAKE CARE OF THE PERCENT VARIABLES
    .InfluentWeir.pr_Stripping = 0#
    
    .GritChamber.pr_Stripping = 0#
    .GritChamber.pr_Volatilization = 0#
    
    .PrimaryClarifier.pr_Volatilization = 0#
    .PrimaryClarifier.pr_SolidWaste = 0#
    .PrimaryClarifier.pr_LiquidWaste = 0#
    
    .PrimaryWeir.pr_Stripping = 0#
    
    .AerationBasin.pr_Stripping = 0#
    .AerationBasin.pr_Volatilization = 0#
    .AerationBasin.pr_Biodegredation = 0#
    
    .SecondaryClarifier.pr_Volatilization = 0#
    .SecondaryClarifier.pr_SolidWaste = 0#
    .SecondaryClarifier.pr_LiquidWaste = 0#
    
    .SecondaryWeir.pr_Stripping = 0#
  End With
End Sub


Sub Project_SetDefaults(Prj As Project_Type)
  'Prj.length = 1#
  'Prj.Diameter = 1#
  'Prj.Mass = 1#
  'Prj.FlowRate = 1#
  '
  ' SET INPUT DEFAULTS.
  '
  Call Project_Plant_SetDefaults(Prj.Plant)
  '
  ' SET OUTPUT DEFAULTS.
  '
  Call Project_OutputRec_SetDefaults(Prj.OutputRec)
  '
  ' MISCELLANEOUS.
  '
  Prj.UnitType = UnitType___SI
  Calculated_OK = False
End Sub


