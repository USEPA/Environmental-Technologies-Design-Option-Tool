Attribute VB_Name = "ForDebugMod"
Option Explicit


Sub DLL_PrepCall()
  Call ChangeDir_Main
End Sub


Sub AIRDENS(AirDensity As Double, Temperature As Double, Pressure As Double)
  Call DLL_PrepCall
  'Call system_log("AIRDENS Entry")
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("Temperature =" & Str$(Temperature))
  'Call system_log("Pressure =" & Str$(Pressure))
  Call Fortran_AIRDENS(AirDensity, Temperature, Pressure)
  'Call system_log("AIRDENS Exit")
End Sub

Sub AIRFLO(AirFlowRate As Double, AirToWaterRatio As Double, WaterFlowRate As Double)
  Call DLL_PrepCall
  'Call system_log("AIRFLO Entry")
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  Call Fortran_AIRFLO(AirFlowRate, AirToWaterRatio, WaterFlowRate)
  'Call system_log("AIRFLO Exit")
End Sub

Sub AIRVISC(AirViscosity As Double, Temperature As Double)
  Call DLL_PrepCall
  'Call system_log("AIRVISC Entry")
  'Call system_log("AirViscosity =" & Str$(AirViscosity))
  'Call system_log("Temperature =" & Str$(Temperature))
  Call Fortran_AIRVISC(AirViscosity, Temperature)
  'Call system_log("AIRVISC Exit")
End Sub

Sub AREAPT2(TowerArea As Double, TowerDiameter As Double)
  Call DLL_PrepCall
  'Call system_log("AREAPT2 Entry")
  'Call system_log("TowerArea =" & Str$(TowerArea))
  'Call system_log("TowerDiameter =" & Str$(TowerDiameter))
  Call Fortran_AREAPT2(TowerArea, TowerDiameter)
  'Call system_log("AREAPT2 Exit")
End Sub

Sub AWCALC(PackingWettedSurfaceArea As Double, PackingCriticalSurfaceTension As Double, WaterSurfaceTension As Double, WaterLoadingRate As Double, PackingSpecificSurfaceArea As Double, WaterViscosity As Double, WaterDensity As Double, ReynoldsNumber As Double, FroudeNumber As Double, WeberNumber As Double)
  Call DLL_PrepCall
  'Call system_log("AWCALC Entry")
  'Call system_log("PackingWettedSurfaceArea =" & Str$(PackingWettedSurfaceArea))
  'Call system_log("PackingCriticalSurfaceTension =" & Str$(PackingCriticalSurfaceTension))
  'Call system_log("WaterSurfaceTension =" & Str$(WaterSurfaceTension))
  'Call system_log("WaterLoadingRate =" & Str$(WaterLoadingRate))
  'Call system_log("PackingSpecificSurfaceArea =" & Str$(PackingSpecificSurfaceArea))
  'Call system_log("WaterViscosity =" & Str$(WaterViscosity))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("ReynoldsNumber =" & Str$(ReynoldsNumber))
  'Call system_log("FroudeNumber =" & Str$(FroudeNumber))
  'Call system_log("WeberNumber =" & Str$(WeberNumber))
  Call Fortran_AWCALC(PackingWettedSurfaceArea, PackingCriticalSurfaceTension, WaterSurfaceTension, WaterLoadingRate, PackingSpecificSurfaceArea, WaterViscosity, WaterDensity, ReynoldsNumber, FroudeNumber, WeberNumber)
  'Call system_log("AWCALC Exit")
End Sub

Sub DIFO2(DiffusivityOxygen As Double, Temperature As Double)
  Call DLL_PrepCall
  'Call system_log("DIFO2 Entry")
  'Call system_log("DiffusivityOxygen =" & Str$(DiffusivityOxygen))
  'Call system_log("Temperature =" & Str$(Temperature))
  Call Fortran_DIFO2(DiffusivityOxygen, Temperature)
  'Call system_log("DIFO2 Exit")
End Sub

Sub EFFLBUB(ArrayLiqPhaseEffluentConc As Double, ArrayGasPhaseEffluentConc As Double, HenrysConstOfCompound As Double, LiqPhaseInfluentConc As Double, AirToWaterRatio As Double, NoOfTanks As Long, StantonNo As Double)
  Call DLL_PrepCall
  'Call system_log("EFFLBUB Entry")
  'Call system_log("ArrayLiqPhaseEffluentConc =" & Str$(ArrayLiqPhaseEffluentConc))
  'Call system_log("ArrayGasPhaseEffluentConc =" & Str$(ArrayGasPhaseEffluentConc))
  'Call system_log("HenrysConstOfCompound =" & Str$(HenrysConstOfCompound))
  'Call system_log("LiqPhaseInfluentConc =" & Str$(LiqPhaseInfluentConc))
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("NoOfTanks =" & Str$(NoOfTanks))
  'Call system_log("StantonNo =" & Str$(StantonNo))
  Call Fortran_EFFLBUB(ArrayLiqPhaseEffluentConc, ArrayGasPhaseEffluentConc, HenrysConstOfCompound, LiqPhaseInfluentConc, AirToWaterRatio, NoOfTanks, StantonNo)
  'Call system_log("EFFLBUB Exit")
End Sub

Sub EFFLPT2(EffluentConcentration As Double, AirToWaterRatio As Double, HenrysConstant As Double, WaterFlowRate As Double, TowerArea As Double, TowerLength As Double, DesignMassTransferCoefficient As Double, InfluentConcentration As Double)
  Call DLL_PrepCall
  'Call system_log("EFFLPT2 Entry")
  'Call system_log("EffluentConcentration =" & Str$(EffluentConcentration))
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("HenrysConstant =" & Str$(HenrysConstant))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  'Call system_log("TowerLength =" & Str$(TowerLength))
  'Call system_log("DesignMassTransferCoefficient =" & Str$(DesignMassTransferCoefficient))
  'Call system_log("InfluentConcentration =" & Str$(InfluentConcentration))
  Call Fortran_EFFLPT2(EffluentConcentration, AirToWaterRatio, HenrysConstant, WaterFlowRate, TowerArea, TowerLength, DesignMassTransferCoefficient, InfluentConcentration)
  'Call system_log("EFFLPT2 Exit")
End Sub

Sub GETCSPT(DesignContaminantAirWaterInterfaceConc As Double, AirToWaterRatio As Double, DesignContaminantHenrysConstant As Double, DesignContaminantInfluentConcentration As Double, DesignContaminantTreatmentObjective As Double)
  Call DLL_PrepCall
  'Call system_log("GETCSPT Entry")
  'Call system_log("DesignContaminantAirWaterInterfaceConc =" & Str$(DesignContaminantAirWaterInterfaceConc))
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("DesignContaminantHenrysConstant =" & Str$(DesignContaminantHenrysConstant))
  'Call system_log("DesignContaminantInfluentConcentration =" & Str$(DesignContaminantInfluentConcentration))
  'Call system_log("DesignContaminantTreatmentObjective =" & Str$(DesignContaminantTreatmentObjective))
  Call Fortran_GETCSPT(DesignContaminantAirWaterInterfaceConc, AirToWaterRatio, DesignContaminantHenrysConstant, DesignContaminantInfluentConcentration, DesignContaminantTreatmentObjective)
  'Call system_log("GETCSPT Exit")
End Sub

Sub GETCSTAR(DOSaturationConc As Double, WeightDensityWater As Double, EffectiveSaturationDepth As Double, BarometricPressure As Double, WaterDepth As Double)
  Call DLL_PrepCall
  'Call system_log("GETCSTAR Entry")
  'Call system_log("DOSaturationConc =" & Str$(DOSaturationConc))
  'Call system_log("WeightDensityWater =" & Str$(WeightDensityWater))
  'Call system_log("EffectiveSaturationDepth =" & Str$(EffectiveSaturationDepth))
  'Call system_log("BarometricPressure =" & Str$(BarometricPressure))
  'Call system_log("WaterDepth =" & Str$(WaterDepth))
  Call Fortran_GETCSTAR(DOSaturationConc, WeightDensityWater, EffectiveSaturationDepth, BarometricPressure, WaterDepth)
  'Call system_log("GETCSTAR Exit")
End Sub

Sub GETHTUPT(TransferUnitHeight As Double, WaterFlowRate As Double, TowerArea As Double, DesignMassTransferCoefficient As Double)
  Call DLL_PrepCall
  'Call system_log("GETHTUPT Entry")
  'Call system_log("TransferUnitHeight =" & Str$(TransferUnitHeight))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  'Call system_log("DesignMassTransferCoefficient =" & Str$(DesignMassTransferCoefficient))
  Call Fortran_GETHTUPT(TransferUnitHeight, WaterFlowRate, TowerArea, DesignMassTransferCoefficient)
  'Call system_log("GETHTUPT Exit")
End Sub

Sub GETMULT(MultipleOfMinimumAirToWaterRatio As Double, AirToWaterRatio As Double, MinimumAirToWaterRatio As Double)
  Call DLL_PrepCall
  'Call system_log("GETMULT Entry")
  'Call system_log("MultipleOfMinimumAirToWaterRatio =" & Str$(MultipleOfMinimumAirToWaterRatio))
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("MinimumAirToWaterRatio =" & Str$(MinimumAirToWaterRatio))
  Call Fortran_GETMULT(MultipleOfMinimumAirToWaterRatio, AirToWaterRatio, MinimumAirToWaterRatio)
  'Call system_log("GETMULT Exit")
End Sub

Sub GETNTUPT(NumberOfTransferUnits As Double, DesignContaminantInfluentConcentration As Double, DesignContaminantTreatmentObjective As Double, DesignContaminantAirToWaterInterfaceConc As Double)
  Call DLL_PrepCall
  'Call system_log("GETNTUPT Entry")
  'Call system_log("NumberOfTransferUnits =" & Str$(NumberOfTransferUnits))
  'Call system_log("DesignContaminantInfluentConcentration =" & Str$(DesignContaminantInfluentConcentration))
  'Call system_log("DesignContaminantTreatmentObjective =" & Str$(DesignContaminantTreatmentObjective))
  'Call system_log("DesignContaminantAirToWaterInterfaceConc =" & Str$(DesignContaminantAirToWaterInterfaceConc))
  Call Fortran_GETNTUPT(NumberOfTransferUnits, DesignContaminantInfluentConcentration, DesignContaminantTreatmentObjective, DesignContaminantAirToWaterInterfaceConc)
  'Call system_log("GETNTUPT Exit")
End Sub

Sub GETPHIB(StantonNo As Double, CompoundMassTransCoeff As Double, VolumeOfEaTank As Double, HenrysConstOfCompound As Double, AirFlowRateToEaTank As Double)
  Call DLL_PrepCall
  'Call system_log("GETPHIB Entry")
  'Call system_log("StantonNo =" & Str$(StantonNo))
  'Call system_log("CompoundMassTransCoeff =" & Str$(CompoundMassTransCoeff))
  'Call system_log("VolumeOfEaTank =" & Str$(VolumeOfEaTank))
  'Call system_log("HenrysConstOfCompound =" & Str$(HenrysConstOfCompound))
  'Call system_log("AirFlowRateToEaTank =" & Str$(AirFlowRateToEaTank))
  Call Fortran_GETPHIB(StantonNo, CompoundMassTransCoeff, VolumeOfEaTank, HenrysConstOfCompound, AirFlowRateToEaTank)
  'Call system_log("GETPHIB Exit")
End Sub

Sub GETSAF(KLaSafetyFactor As Double, OndaMassTransferCoefficient As Double, DesignMassTransferCoefficient As Double)
  Call DLL_PrepCall
  'Call system_log("GETSAF Entry")
  'Call system_log("KLaSafetyFactor =" & Str$(KLaSafetyFactor))
  'Call system_log("OndaMassTransferCoefficient =" & Str$(OndaMassTransferCoefficient))
  'Call system_log("DesignMassTransferCoefficient =" & Str$(DesignMassTransferCoefficient))
  Call Fortran_GETSAF(KLaSafetyFactor, OndaMassTransferCoefficient, DesignMassTransferCoefficient)
  'Call system_log("GETSAF Exit")
End Sub

Sub GETSOTE(StandardOxygenTransferEff As Double, StandardOxygenTransferRate As Double, AirFlowRate As Double)
  Call DLL_PrepCall
  'Call system_log("GETSOTE Entry")
  'Call system_log("StandardOxygenTransferEff =" & Str$(StandardOxygenTransferEff))
  'Call system_log("StandardOxygenTransferRate =" & Str$(StandardOxygenTransferRate))
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  Call Fortran_GETSOTE(StandardOxygenTransferEff, StandardOxygenTransferRate, AirFlowRate)
  'Call system_log("GETSOTE Exit")
End Sub

Sub GETSOTR(StandardOxygenTransferRate As Double, StandardOxygenTransferEff As Double, AirFlowRate As Double)
  Call DLL_PrepCall
  'Call system_log("GETSOTR Entry")
  'Call system_log("StandardOxygenTransferRate =" & Str$(StandardOxygenTransferRate))
  'Call system_log("StandardOxygenTransferEff =" & Str$(StandardOxygenTransferEff))
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  Call Fortran_GETSOTR(StandardOxygenTransferRate, StandardOxygenTransferEff, AirFlowRate)
  'Call system_log("GETSOTR Exit")
End Sub

Sub H2ODENS(LiquidDensity As Double, Temperature As Double)
  Call DLL_PrepCall
  'Call system_log("H2ODENS Entry")
  'Call system_log("LiquidDensity =" & Str$(LiquidDensity))
  'Call system_log("Temperature =" & Str$(Temperature))
  Call Fortran_H2ODENS(LiquidDensity, Temperature)
  'Call system_log("H2ODENS Exit")
End Sub

Sub H2OST(LiquidSurfaceTension As Double, Temperature As Double)
  Call DLL_PrepCall
  'Call system_log("H2OST Entry")
  'Call system_log("LiquidSurfaceTension =" & Str$(LiquidSurfaceTension))
  'Call system_log("Temperature =" & Str$(Temperature))
  Call Fortran_H2OST(LiquidSurfaceTension, Temperature)
  'Call system_log("H2OST Exit")
End Sub

Sub H2OVISC(LiquidViscosity As Double, Temperature As Double)
  Call DLL_PrepCall
  'Call system_log("H2OVISC Entry")
  'Call system_log("LiquidViscosity =" & Str$(LiquidViscosity))
  'Call system_log("Temperature =" & Str$(Temperature))
  Call Fortran_H2OVISC(LiquidViscosity, Temperature)
  'Call system_log("H2OVISC Exit")
End Sub

Sub KLA20A(AppOxygenMassTransCoeff As Double, WaterVolumePerTankL As Double, WaterVolumePerTankm3 As Double, DOSaturationConcentration As Double, StandOxygenMassTransRate As Double)
  Call DLL_PrepCall
  'Call system_log("KLA20A Entry")
  'Call system_log("AppOxygenMassTransCoeff =" & Str$(AppOxygenMassTransCoeff))
  'Call system_log("WaterVolumePerTankL =" & Str$(WaterVolumePerTankL))
  'Call system_log("WaterVolumePerTankm3 =" & Str$(WaterVolumePerTankm3))
  'Call system_log("DOSaturationConcentration =" & Str$(DOSaturationConcentration))
  'Call system_log("StandOxygenMassTransRate =" & Str$(StandOxygenMassTransRate))
  Call Fortran_KLA20A(AppOxygenMassTransCoeff, WaterVolumePerTankL, WaterVolumePerTankm3, DOSaturationConcentration, StandOxygenMassTransRate)
  'Call system_log("KLA20A Exit")
End Sub

Sub KLABUB(CompoundMassTransferCoeff As Double, OxygenMassTransferCoeff As Double, DiffusivityLiquidWater As Double, DiffusivityOfOxygen As Double, ExponentInCorrelation As Double, RatioGasLiquidTransfer As Double, HenrysConstant As Double)
  Call DLL_PrepCall
  'Call system_log("KLABUB Entry")
  'Call system_log("CompoundMassTransferCoeff =" & Str$(CompoundMassTransferCoeff))
  'Call system_log("OxygenMassTransferCoeff =" & Str$(OxygenMassTransferCoeff))
  'Call system_log("DiffusivityLiquidWater =" & Str$(DiffusivityLiquidWater))
  'Call system_log("DiffusivityOfOxygen =" & Str$(DiffusivityOfOxygen))
  'Call system_log("ExponentInCorrelation =" & Str$(ExponentInCorrelation))
  'Call system_log("RatioGasLiquidTransfer =" & Str$(RatioGasLiquidTransfer))
  'Call system_log("HenrysConstant =" & Str$(HenrysConstant))
  Call Fortran_KLABUB(CompoundMassTransferCoeff, OxygenMassTransferCoeff, DiffusivityLiquidWater, DiffusivityOfOxygen, ExponentInCorrelation, RatioGasLiquidTransfer, HenrysConstant)
  'Call system_log("KLABUB Exit")
End Sub

Sub KLACOR(DesignMassTransferCoefficient As Double, OndaMassTransferCoefficient As Double, KLaSafetyFactor As Double)
  Call DLL_PrepCall
  'Call system_log("KLACOR Entry")
  'Call system_log("DesignMassTransferCoefficient =" & Str$(DesignMassTransferCoefficient))
  'Call system_log("OndaMassTransferCoefficient =" & Str$(OndaMassTransferCoefficient))
  'Call system_log("KLaSafetyFactor =" & Str$(KLaSafetyFactor))
  Call Fortran_KLACOR(DesignMassTransferCoefficient, OndaMassTransferCoefficient, KLaSafetyFactor)
  'Call system_log("KLACOR Exit")
End Sub

'DLL Declarations for Surface Aeration
Sub KLAO2SUR(OxygenMTCoeff As Double, PowerInput_PoverV As Double)
  Call DLL_PrepCall
  'Call system_log("KLAO2SUR Entry")
  'Call system_log("OxygenMTCoeff =" & Str$(OxygenMTCoeff))
  'Call system_log("PowerInput_PoverV =" & Str$(PowerInput_PoverV))
  Call Fortran_KLAO2SUR(OxygenMTCoeff, PowerInput_PoverV)
  'Call system_log("KLAO2SUR Exit")
End Sub

Sub KLASURF(ContaminantMassTransferCoeff As Double, OxygenMassTransferCoeff As Double, ContaminantLiquidDiffusivity As Double, OxygenLiquidDiffusivity As Double, N_forFindingKLa As Double, kgOVERkl_forFindingKLa As Double, HenrysConstant As Double)
  Call DLL_PrepCall
  'Call system_log("KLASURF Entry")
  'Call system_log("ContaminantMassTransferCoeff =" & Str$(ContaminantMassTransferCoeff))
  'Call system_log("OxygenMassTransferCoeff =" & Str$(OxygenMassTransferCoeff))
  'Call system_log("ContaminantLiquidDiffusivity =" & Str$(ContaminantLiquidDiffusivity))
  'Call system_log("OxygenLiquidDiffusivity =" & Str$(OxygenLiquidDiffusivity))
  'Call system_log("N_forFindingKLa =" & Str$(N_forFindingKLa))
  'Call system_log("kgOVERkl_forFindingKLa =" & Str$(kgOVERkl_forFindingKLa))
  'Call system_log("HenrysConstant =" & Str$(HenrysConstant))
  Call Fortran_KLASURF(ContaminantMassTransferCoeff, OxygenMassTransferCoeff, ContaminantLiquidDiffusivity, OxygenLiquidDiffusivity, N_forFindingKLa, kgOVERkl_forFindingKLa, HenrysConstant)
  'Call system_log("KLASURF Exit")
End Sub

Sub LDAIRPT2(AirLoadingRate As Double, AirFlowRate As Double, AirDensity As Double, TowerArea As Double)
  Call DLL_PrepCall
  'Call system_log("LDAIRPT2 Entry")
  'Call system_log("AirLoadingRate =" & Str$(AirLoadingRate))
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  Call Fortran_LDAIRPT2(AirLoadingRate, AirFlowRate, AirDensity, TowerArea)
  'Call system_log("LDAIRPT2 Exit")
End Sub

Sub LDH2OPT2(WaterLoadingRate As Double, WaterFlowRate As Double, WaterDensity As Double, TowerArea As Double)
  Call DLL_PrepCall
  'Call system_log("LDH2OPT2 Entry")
  'Call system_log("WaterLoadingRate =" & Str$(WaterLoadingRate))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  Call Fortran_LDH2OPT2(WaterLoadingRate, WaterFlowRate, WaterDensity, TowerArea)
  'Call system_log("LDH2OPT2 Exit")
End Sub

Sub ONDAKGPT(OndaGasPhaseMassTransferCoefficient As Double, AirLoadingRate As Double, PackingSpecificSurfaceArea As Double, AirViscosity As Double, AirDensity As Double, GasDiffusivity As Double, PackingNominalSize As Double)
  Call DLL_PrepCall
  'Call system_log("ONDAKGPT Entry")
  'Call system_log("OndaGasPhaseMassTransferCoefficient =" & Str$(OndaGasPhaseMassTransferCoefficient))
  'Call system_log("AirLoadingRate =" & Str$(AirLoadingRate))
  'Call system_log("PackingSpecificSurfaceArea =" & Str$(PackingSpecificSurfaceArea))
  'Call system_log("AirViscosity =" & Str$(AirViscosity))
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("GasDiffusivity =" & Str$(GasDiffusivity))
  'Call system_log("PackingNominalSize =" & Str$(PackingNominalSize))
  Call Fortran_ONDAKGPT(OndaGasPhaseMassTransferCoefficient, AirLoadingRate, PackingSpecificSurfaceArea, AirViscosity, AirDensity, GasDiffusivity, PackingNominalSize)
  'Call system_log("ONDAKGPT Exit")
End Sub

Sub ONDAKLPT(OndaLiquidPhaseMassTransferCoefficient As Double, WaterLoadingRate As Double, PackingWettedSurfaceArea As Double, WaterViscosity As Double, WaterDensity As Double, LiquidDiffusivity As Double, PackingSpecificSurfaceArea As Double, PackingNominalSize As Double)
  Call DLL_PrepCall
  'Call system_log("ONDAKLPT Entry")
  'Call system_log("OndaLiquidPhaseMassTransferCoefficient =" & Str$(OndaLiquidPhaseMassTransferCoefficient))
  'Call system_log("WaterLoadingRate =" & Str$(WaterLoadingRate))
  'Call system_log("PackingWettedSurfaceArea =" & Str$(PackingWettedSurfaceArea))
  'Call system_log("WaterViscosity =" & Str$(WaterViscosity))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("LiquidDiffusivity =" & Str$(LiquidDiffusivity))
  'Call system_log("PackingSpecificSurfaceArea =" & Str$(PackingSpecificSurfaceArea))
  'Call system_log("PackingNominalSize =" & Str$(PackingNominalSize))
  Call Fortran_ONDAKLPT(OndaLiquidPhaseMassTransferCoefficient, WaterLoadingRate, PackingWettedSurfaceArea, WaterViscosity, WaterDensity, LiquidDiffusivity, PackingSpecificSurfaceArea, PackingNominalSize)
  'Call system_log("ONDAKLPT Exit")
End Sub

Sub ONDKLAPT(OndaOverallMassTransferCoefficient As Double, OndaLiquidPhaseResistance As Double, OndaGasPhaseResistance As Double, OndaTotalResistance As Double, OndaLiquidPhaseMassTransferCoefficient As Double, PackingWettedSurfaceArea As Double, OndaGasPhaseMassTransferCoefficient As Double, DesignContaminantHenrysConstant As Double)
  Call DLL_PrepCall
  'Call system_log("ONDKLAPT Entry")
  'Call system_log("OndaOverallMassTransferCoefficient =" & Str$(OndaOverallMassTransferCoefficient))
  'Call system_log("OndaLiquidPhaseResistance =" & Str$(OndaLiquidPhaseResistance))
  'Call system_log("OndaGasPhaseResistance =" & Str$(OndaGasPhaseResistance))
  'Call system_log("OndaTotalResistance =" & Str$(OndaTotalResistance))
  'Call system_log("OndaLiquidPhaseMassTransferCoefficient =" & Str$(OndaLiquidPhaseMassTransferCoefficient))
  'Call system_log("PackingWettedSurfaceArea =" & Str$(PackingWettedSurfaceArea))
  'Call system_log("OndaGasPhaseMassTransferCoefficient =" & Str$(OndaGasPhaseMassTransferCoefficient))
  'Call system_log("DesignContaminantHenrysConstant =" & Str$(DesignContaminantHenrysConstant))
  Call Fortran_ONDKLAPT(OndaOverallMassTransferCoefficient, OndaLiquidPhaseResistance, OndaGasPhaseResistance, OndaTotalResistance, OndaLiquidPhaseMassTransferCoefficient, PackingWettedSurfaceArea, OndaGasPhaseMassTransferCoefficient, DesignContaminantHenrysConstant)
  'Call system_log("ONDKLAPT Exit")
End Sub

Sub OPTMAL(WaterDensity As Double, WaterViscosity As Double, WaterSurfaceTension As Double, AirDensity As Double, AirViscosity As Double, WaterFlowRate As Double, PackingNominalSize As Double, PackingFactor As Double, PackingCriticalSurfaceTension As Double, PackingSpecificSurfaceArea As Double, InfluentConcentrations As Double, TreatmentObjectives As Double, HenrysConstants As Double, NumberOfContaminants As Long, PressureDrop As Double, LiquidDiffusivities As Double, GasDiffusivities As Double, KLaSafetyFactor As Double, ID_OptimalDesignContaminant As Long, MultipleOfMinimumAirToWaterRatio As Double, EffluentConcentrations As Double, ErrorFlag As Long)
  Call DLL_PrepCall
  'Call system_log("OPTMAL Entry")
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("WaterViscosity =" & Str$(WaterViscosity))
  'Call system_log("WaterSurfaceTension =" & Str$(WaterSurfaceTension))
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("AirViscosity =" & Str$(AirViscosity))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  'Call system_log("PackingNominalSize =" & Str$(PackingNominalSize))
  'Call system_log("PackingFactor =" & Str$(PackingFactor))
  'Call system_log("PackingCriticalSurfaceTension =" & Str$(PackingCriticalSurfaceTension))
  'Call system_log("PackingSpecificSurfaceArea =" & Str$(PackingSpecificSurfaceArea))
  'Call system_log("InfluentConcentrations =" & Str$(InfluentConcentrations))
  'Call system_log("TreatmentObjectives =" & Str$(TreatmentObjectives))
  'Call system_log("HenrysConstants =" & Str$(HenrysConstants))
  'Call system_log("NumberOfContaminants =" & Str$(NumberOfContaminants))
  'Call system_log("PressureDrop =" & Str$(PressureDrop))
  'Call system_log("LiquidDiffusivities =" & Str$(LiquidDiffusivities))
  'Call system_log("GasDiffusivities =" & Str$(GasDiffusivities))
  'Call system_log("KLaSafetyFactor =" & Str$(KLaSafetyFactor))
  'Call system_log("ID_OptimalDesignContaminant =" & Str$(ID_OptimalDesignContaminant))
  'Call system_log("MultipleOfMinimumAirToWaterRatio =" & Str$(MultipleOfMinimumAirToWaterRatio))
  'Call system_log("EffluentConcentrations =" & Str$(EffluentConcentrations))
  'Call system_log("ErrorFlag =" & Str$(ErrorFlag))

  Call Fortran_OPTMAL(WaterDensity, WaterViscosity, WaterSurfaceTension, AirDensity, AirViscosity, WaterFlowRate, PackingNominalSize, PackingFactor, PackingCriticalSurfaceTension, PackingSpecificSurfaceArea, InfluentConcentrations, TreatmentObjectives, HenrysConstants, NumberOfContaminants, PressureDrop, LiquidDiffusivities, GasDiffusivities, KLaSafetyFactor, ID_OptimalDesignContaminant, MultipleOfMinimumAirToWaterRatio, EffluentConcentrations, ErrorFlag)

  'Call system_log("OPTMAL Exit")
End Sub

Sub PBLOWPT(BlowerBrakePower As Double, AirFlowRate As Double, TowerArea As Double, OperatingPressure As Double, PressureDrop As Double, TowerHeight As Double, AirDensity As Double, InletAirTemperature As Double, BlowerEfficiency As Double)
  Call DLL_PrepCall
  'Call system_log("PBLOWPT Entry")
  'Call system_log("BlowerBrakePower =" & Str$(BlowerBrakePower))
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  'Call system_log("OperatingPressure =" & Str$(OperatingPressure))
  'Call system_log("PressureDrop =" & Str$(PressureDrop))
  'Call system_log("TowerHeight =" & Str$(TowerHeight))
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("InletAirTemperature =" & Str$(InletAirTemperature))
  'Call system_log("BlowerEfficiency =" & Str$(BlowerEfficiency))
  Call Fortran_PBLOWPT(BlowerBrakePower, AirFlowRate, TowerArea, OperatingPressure, PressureDrop, TowerHeight, AirDensity, InletAirTemperature, BlowerEfficiency)
  'Call system_log("PBLOWPT Exit")
End Sub

Sub PCALCBUB(TotalBrakePowerAllTanks As Double, BlowerBrakePowerForEaTank As Double, OperatingPressure As Double, InletAirTempC As Double, AirFlowRate As Double, BlowerEfficiencyPercent As Double, LiquidDensity As Double, WaterDepth As Double, NoOfTanks As Long, NumberOfBlowersinEachTank As Long)
  Call DLL_PrepCall
  'Call system_log("PCALCBUB Entry")
  'Call system_log("TotalBrakePowerAllTanks =" & Str$(TotalBrakePowerAllTanks))
  'Call system_log("BlowerBrakePowerForEaTank =" & Str$(BlowerBrakePowerForEaTank))
  'Call system_log("OperatingPressure =" & Str$(OperatingPressure))
  'Call system_log("InletAirTempC =" & Str$(InletAirTempC))
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  'Call system_log("BlowerEfficiencyPercent =" & Str$(BlowerEfficiencyPercent))
  'Call system_log("LiquidDensity =" & Str$(LiquidDensity))
  'Call system_log("WaterDepth =" & Str$(WaterDepth))
  'Call system_log("NoOfTanks =" & Str$(NoOfTanks))
  'Call system_log("NumberOfBlowersinEachTank =" & Str$(NumberOfBlowersinEachTank))
  Call Fortran_PCALCBUB(TotalBrakePowerAllTanks, BlowerBrakePowerForEaTank, OperatingPressure, InletAirTempC, AirFlowRate, BlowerEfficiencyPercent, LiquidDensity, WaterDepth, NoOfTanks, NumberOfBlowersinEachTank)
  'Call system_log("PCALCBUB Exit")
End Sub

Sub PCALCSUR(TotalPower As Double, PowerPerTank As Double, PowerInput_PoverV As Double, TotalVolumeAllTanks As Double, NumberOfTanks As Long, AeratorMotorEfficiency As Double)
  Call DLL_PrepCall
  'Call system_log("PCALCSUR Entry")
  'Call system_log("TotalPower =" & Str$(TotalPower))
  'Call system_log("PowerPerTank =" & Str$(PowerPerTank))
  'Call system_log("PowerInput_PoverV =" & Str$(PowerInput_PoverV))
  'Call system_log("TotalVolumeAllTanks =" & Str$(TotalVolumeAllTanks))
  'Call system_log("NumberOfTanks =" & Str$(NumberOfTanks))
  'Call system_log("AeratorMotorEfficiency =" & Str$(AeratorMotorEfficiency))
  Call Fortran_PCALCSUR(TotalPower, PowerPerTank, PowerInput_PoverV, TotalVolumeAllTanks, NumberOfTanks, AeratorMotorEfficiency)
  'Call system_log("PCALCSUR Exit")
End Sub

Sub PDROP(AirPressureDrop As Double, AirToWaterRatio As Double, AirLoadingRate As Double, PackingFactor As Double, WaterViscosity As Double, AirDensity As Double, WaterDensity As Double, InitialPressureDrop As Double, MaximumPressureDrop As Double, PressureDropStep As Double)
  Call DLL_PrepCall
  'Call system_log("PDROP Entry")
  'Call system_log("AirPressureDrop =" & Str$(AirPressureDrop))
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("AirLoadingRate =" & Str$(AirLoadingRate))
  'Call system_log("PackingFactor =" & Str$(PackingFactor))
  'Call system_log("WaterViscosity =" & Str$(WaterViscosity))
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("InitialPressureDrop =" & Str$(InitialPressureDrop))
  'Call system_log("MaximumPressureDrop =" & Str$(MaximumPressureDrop))
  'Call system_log("PressureDropStep =" & Str$(PressureDropStep))
  Call Fortran_PDROP(AirPressureDrop, AirToWaterRatio, AirLoadingRate, PackingFactor, WaterViscosity, AirDensity, WaterDensity, InitialPressureDrop, MaximumPressureDrop, PressureDropStep)
  'Call system_log("PDROP Exit")
End Sub

Sub PPUMPPT(PumpBrakePower As Double, PumpEfficiency As Double, WaterDensity As Double, WaterFlowRate As Double, TowerHeight As Double)
  Call DLL_PrepCall
  'Call system_log("PPUMPPT Entry")
  'Call system_log("PumpBrakePower =" & Str$(PumpBrakePower))
  'Call system_log("PumpEfficiency =" & Str$(PumpEfficiency))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  'Call system_log("TowerHeight =" & Str$(TowerHeight))
  Call Fortran_PPUMPPT(PumpBrakePower, PumpEfficiency, WaterDensity, WaterFlowRate, TowerHeight)
  'Call system_log("PPUMPPT Exit")
End Sub

Sub PT1AREA(TowerArea As Double, WaterFlowRate As Double, WaterDensity As Double, WaterMassLoadingRate As Double)
  Call DLL_PrepCall
  'Call system_log("PT1AREA Entry")
  'Call system_log("TowerArea =" & Str$(TowerArea))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("WaterMassLoadingRate =" & Str$(WaterMassLoadingRate))
  Call Fortran_PT1AREA(TowerArea, WaterFlowRate, WaterDensity, WaterMassLoadingRate)
  'Call system_log("PT1AREA Exit")
End Sub

Sub PT1DTOW(TowerDiameter As Double, TowerArea As Double)
  Call DLL_PrepCall
  'Call system_log("PT1DTOW Entry")
  'Call system_log("TowerDiameter =" & Str$(TowerDiameter))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  Call Fortran_PT1DTOW(TowerDiameter, TowerArea)
  'Call system_log("PT1DTOW Exit")
End Sub

Sub PT1HTOW(TowerHeight As Double, TransferUnitHeight As Double, NumberOfTransferUnits As Double)
  Call DLL_PrepCall
  'Call system_log("PT1HTOW Entry")
  'Call system_log("TowerHeight =" & Str$(TowerHeight))
  'Call system_log("TransferUnitHeight =" & Str$(TransferUnitHeight))
  'Call system_log("NumberOfTransferUnits =" & Str$(NumberOfTransferUnits))
  Call Fortran_PT1HTOW(TowerHeight, TransferUnitHeight, NumberOfTransferUnits)
  'Call system_log("PT1HTOW Exit")
End Sub

Sub PT1LDAIR(AirMassLoadingRate As Double, AirPressureDrop As Double, AirToWaterRatio As Double, AirDensity As Double, WaterDensity As Double, PackingFactor As Double, WaterViscosity As Double)
  Call DLL_PrepCall
  'Call system_log("PT1LDAIR Entry")
  'Call system_log("AirMassLoadingRate =" & Str$(AirMassLoadingRate))
  'Call system_log("AirPressureDrop =" & Str$(AirPressureDrop))
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("PackingFactor =" & Str$(PackingFactor))
  'Call system_log("WaterViscosity =" & Str$(WaterViscosity))
  Call Fortran_PT1LDAIR(AirMassLoadingRate, AirPressureDrop, AirToWaterRatio, AirDensity, WaterDensity, PackingFactor, WaterViscosity)
  'Call system_log("PT1LDAIR Exit")
End Sub

Sub PT1LDH2O(WaterMassLoadingRate As Double, AirToWaterRatio As Double, AirDensity As Double, WaterDensity As Double, AirMassLoadingRate As Double)
  Call DLL_PrepCall
  'Call system_log("PT1LDH2O Entry")
  'Call system_log("WaterMassLoadingRate =" & Str$(WaterMassLoadingRate))
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("AirMassLoadingRate =" & Str$(AirMassLoadingRate))
  Call Fortran_PT1LDH2O(WaterMassLoadingRate, AirToWaterRatio, AirDensity, WaterDensity, AirMassLoadingRate)
  'Call system_log("PT1LDH2O Exit")
End Sub

Sub PT1TVOL(TowerVolume As Double, TowerArea As Double, TowerHeight As Double)
  Call DLL_PrepCall
  'Call system_log("PT1TVOL Entry")
  'Call system_log("TowerVolume =" & Str$(TowerVolume))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  'Call system_log("TowerHeight =" & Str$(TowerHeight))
  Call Fortran_PT1TVOL(TowerVolume, TowerArea, TowerHeight)
  'Call system_log("PT1TVOL Exit")
End Sub

Sub PT1VQMIN(MinimumAirToWaterRatio As Double, InfluentConcentration As Double, TreatmentObjective As Double, HenrysConstant As Double)
  Call DLL_PrepCall
  'Call system_log("PT1VQMIN Entry")
  'Call system_log("MinimumAirToWaterRatio =" & Str$(MinimumAirToWaterRatio))
  'Call system_log("InfluentConcentration =" & Str$(InfluentConcentration))
  'Call system_log("TreatmentObjective =" & Str$(TreatmentObjective))
  'Call system_log("HenrysConstant =" & Str$(HenrysConstant))
  Call Fortran_PT1VQMIN(MinimumAirToWaterRatio, InfluentConcentration, TreatmentObjective, HenrysConstant)
  'Call system_log("PT1VQMIN Exit")
End Sub

Sub PTOTALPT(TotalBrakePower As Double, BlowerBrakePower As Double, PumpBrakePower As Double)
  Call DLL_PrepCall
  'Call system_log("PTOTALPT Entry")
  'Call system_log("TotalBrakePower =" & Str$(TotalBrakePower))
  'Call system_log("BlowerBrakePower =" & Str$(BlowerBrakePower))
  'Call system_log("PumpBrakePower =" & Str$(PumpBrakePower))
  Call Fortran_PTOTALPT(TotalBrakePower, BlowerBrakePower, PumpBrakePower)
  'Call system_log("PTOTALPT Exit")
End Sub

Sub QAIRPT2(AirFlowRate As Double, AirLoadingRate As Double, AirDensity As Double, TowerArea As Double)
  Call DLL_PrepCall
  'Call system_log("QAIRPT2 Entry")
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  'Call system_log("AirLoadingRate =" & Str$(AirLoadingRate))
  'Call system_log("AirDensity =" & Str$(AirDensity))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  Call Fortran_QAIRPT2(AirFlowRate, AirLoadingRate, AirDensity, TowerArea)
  'Call system_log("QAIRPT2 Exit")
End Sub

Sub QH2OPT2(WaterFlowRate As Double, WaterLoadingRate As Double, WaterDensity As Double, TowerArea As Double)
  Call DLL_PrepCall
  'Call system_log("QH2OPT2 Entry")
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  'Call system_log("WaterLoadingRate =" & Str$(WaterLoadingRate))
  'Call system_log("WaterDensity =" & Str$(WaterDensity))
  'Call system_log("TowerArea =" & Str$(TowerArea))
  Call Fortran_QH2OPT2(WaterFlowRate, WaterLoadingRate, WaterDensity, TowerArea)
  'Call system_log("QH2OPT2 Exit")
End Sub

Sub REMOVBUB(ActualLiqPhaseRemovalEfficiency As Double, LiqPhaseInfluentConc As Double, LiqPhaseEffluentConcLastTank As Double)
  Call DLL_PrepCall
  'Call system_log("REMOVBUB Entry")
  'Call system_log("ActualLiqPhaseRemovalEfficiency =" & Str$(ActualLiqPhaseRemovalEfficiency))
  'Call system_log("LiqPhaseInfluentConc =" & Str$(LiqPhaseInfluentConc))
  'Call system_log("LiqPhaseEffluentConcLastTank =" & Str$(LiqPhaseEffluentConcLastTank))
  Call Fortran_REMOVBUB(ActualLiqPhaseRemovalEfficiency, LiqPhaseInfluentConc, LiqPhaseEffluentConcLastTank)
  'Call system_log("REMOVBUB Exit")
End Sub

Sub REMOVPT(RemovalEfficiency As Double, InfluentConcentration As Double, Effluent As Double)
  Call DLL_PrepCall
  'Call system_log("REMOVPT Entry")
  'Call system_log("RemovalEfficiency =" & Str$(RemovalEfficiency))
  'Call system_log("InfluentConcentration =" & Str$(InfluentConcentration))
  'Call system_log("Effluent =" & Str$(Effluent))
  Call Fortran_REMOVPT(RemovalEfficiency, InfluentConcentration, Effluent)
  'Call system_log("REMOVPT Exit")
End Sub

Sub SEFFL(EffluentConcentrations As Double, AchievedRemovalEfficiency As Double, Influent As Double, ContaminantMassTransferCoeff As Double, TankResidenceTime As Double, NumberOfTanks As Long)
  Call DLL_PrepCall
  'Call system_log("SEFFL Entry")
  'Call system_log("EffluentConcentrations =" & Str$(EffluentConcentrations))
  'Call system_log("AchievedRemovalEfficiency =" & Str$(AchievedRemovalEfficiency))
  'Call system_log("Influent =" & Str$(Influent))
  'Call system_log("ContaminantMassTransferCoeff =" & Str$(ContaminantMassTransferCoeff))
  'Call system_log("TankResidenceTime =" & Str$(TankResidenceTime))
  'Call system_log("NumberOfTanks =" & Str$(NumberOfTanks))
  Call Fortran_SEFFL(EffluentConcentrations, AchievedRemovalEfficiency, Influent, ContaminantMassTransferCoeff, TankResidenceTime, NumberOfTanks)
  'Call system_log("SEFFL Exit")
End Sub

Sub SURFEFF(RemovalEfficiency As Double, Influent As Double, EffluentOrTreatmentObjective As Double)
  Call DLL_PrepCall
  'Call system_log("SURFEFF Entry")
  'Call system_log("RemovalEfficiency =" & Str$(RemovalEfficiency))
  'Call system_log("Influent =" & Str$(Influent))
  'Call system_log("EffluentOrTreatmentObjective =" & Str$(EffluentOrTreatmentObjective))
  Call Fortran_SURFEFF(RemovalEfficiency, Influent, EffluentOrTreatmentObjective)
  'Call system_log("SURFEFF Exit")
End Sub

Sub TAUISURF(TankResidenceTime As Double, Influent As Double, TreatmentObjective As Double, NumberOfTanks As Long, ContaminantMassTransferCoeff As Double)
  Call DLL_PrepCall
  'Call system_log("TAUISURF Entry")
  'Call system_log("TankResidenceTime =" & Str$(TankResidenceTime))
  'Call system_log("Influent =" & Str$(Influent))
  'Call system_log("TreatmentObjective =" & Str$(TreatmentObjective))
  'Call system_log("NumberOfTanks =" & Str$(NumberOfTanks))
  'Call system_log("ContaminantMassTransferCoeff =" & Str$(ContaminantMassTransferCoeff))
  Call Fortran_TAUISURF(TankResidenceTime, Influent, TreatmentObjective, NumberOfTanks, ContaminantMassTransferCoeff)
  'Call system_log("TAUISURF Exit")
End Sub

Sub TAUSVOLS(TotalFluidResidenceTime As Double, NoOfTanksInSeries As Long, HydraulicRetentTimeOfEaTank As Double, VolumeOfEachTank As Double, TotalVolumeOfAllTanks As Double, WaterFlowRate As Double, TankParametersCode As Long)
  Call DLL_PrepCall
  'Call system_log("TAUSVOLS Entry")
  'Call system_log("TotalFluidResidenceTime = " & TotalFluidResidenceTime)
  'Call system_log("NoOfTanksInSeries = " & NoOfTanksInSeries)
  'Call system_log("HydraulicRetentTimeOfEaTank = " & HydraulicRetentTimeOfEaTank)
  'Call system_log("VolumeOfEachTank = " & VolumeOfEachTank)
  'Call system_log("TotalVolumeOfAllTanks = " & TotalVolumeOfAllTanks)
  'Call system_log("WaterFlowRate = " & WaterFlowRate)
  'Call system_log("TankParametersCode = " & TankParametersCode)
  Call Fortran_TAUSVOLS(TotalFluidResidenceTime, NoOfTanksInSeries, HydraulicRetentTimeOfEaTank, VolumeOfEachTank, TotalVolumeOfAllTanks, WaterFlowRate, TankParametersCode)
  'Call system_log("TAUSVOLS Exit")
End Sub

Sub TrueKLa(OxygenMassTransferCoeffOperatTemp As Double, AppOxygenMassTransferCoeff20Deg As Double, ParameterUsedInKla As Double, AirFlowRate As Double, WaterVolPerTankL As Double, BarometricPressure As Double, WeightDensityWater As Double, TrueOxygenMassTransferCoef20 As Double, EffectiveSaturatonDepth As Double, OperatingTemp As Double)
  Call DLL_PrepCall
  'Call system_log("TrueKLa Entry")
  'Call system_log("OxygenMassTransferCoeffOperatTemp =" & Str$(OxygenMassTransferCoeffOperatTemp))
  'Call system_log("AppOxygenMassTransferCoeff20Deg =" & Str$(AppOxygenMassTransferCoeff20Deg))
  'Call system_log("ParameterUsedInKla =" & Str$(ParameterUsedInKla))
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  'Call system_log("WaterVolPerTankL =" & Str$(WaterVolPerTankL))
  'Call system_log("BarometricPressure =" & Str$(BarometricPressure))
  'Call system_log("WeightDensityWater =" & Str$(WeightDensityWater))
  'Call system_log("TrueOxygenMassTransferCoef20 =" & Str$(TrueOxygenMassTransferCoef20))
  'Call system_log("EffectiveSaturatonDepth =" & Str$(EffectiveSaturatonDepth))
  'Call system_log("OperatingTemp =" & Str$(OperatingTemp))
  Call Fortran_TrueKLa(OxygenMassTransferCoeffOperatTemp, AppOxygenMassTransferCoeff20Deg, ParameterUsedInKla, AirFlowRate, WaterVolPerTankL, BarometricPressure, WeightDensityWater, TrueOxygenMassTransferCoef20, EffectiveSaturatonDepth, OperatingTemp)
  'Call system_log("TrueKLa Exit")
End Sub

Sub TVOLPT2(TowerVolume As Double, TowerArea As Double, TowerLength As Double)
  Call DLL_PrepCall
  ''Call system_log("TVOLPT2 Entry")
  ''Call system_log("TowerVolume =" & Str$(TowerVolume))
  ''Call system_log("TowerAreaTowerLength =" & Str$(TowerAreaTowerLength))
  ''Call system_log("TowerLength =" & Str$(TowerLength))
  Call Fortran_TVOLPT2(TowerVolume, TowerArea, TowerLength)
  ''Call system_log("TVOLPT2 Exit")
End Sub

Sub VOLBUB(TankVolume As Double, HenrysConstant As Double, AirFlowRate As Double, ContaminantMassTransferCoeff As Double, Influent As Double, TreatmentObjective As Double, NumberOfTanks As Long, WaterFlowRate As Double, ErrorFlag As Long)
  Call DLL_PrepCall
  'Call system_log("VOLBUB Entry")
  'Call system_log("TankVolume =" & Str$(TankVolume))
  'Call system_log("HenrysConstant =" & Str$(HenrysConstant))
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  'Call system_log("ContaminantMassTransferCoeff =" & Str$(ContaminantMassTransferCoeff))
  'Call system_log("Influent =" & Str$(Influent))
  'Call system_log("TreatmentObjective =" & Str$(TreatmentObjective))
  'Call system_log("NumberOfTanks =" & Str$(NumberOfTanks))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  'Call system_log("ErrorFlag =" & Str$(ErrorFlag))
  Call Fortran_VOLBUB(TankVolume, HenrysConstant, AirFlowRate, ContaminantMassTransferCoeff, Influent, TreatmentObjective, NumberOfTanks, WaterFlowRate, ErrorFlag)
  'Call system_log("VOLBUB Exit")
End Sub

Sub VQBUB(AirToWaterRatio As Double, AirFlowRateToEachTank As Double, WaterFlowRate As Double)
  Call DLL_PrepCall
  'Call system_log("VQBUB Entry")
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("AirFlowRateToEachTank =" & Str$(AirFlowRateToEachTank))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  Call Fortran_VQBUB(AirToWaterRatio, AirFlowRateToEachTank, WaterFlowRate)
  'Call system_log("VQBUB Exit")
End Sub

Sub VQCALC(AirToWaterRatio As Double, AirFlowRate As Double, WaterFlowRate As Double)
  Call DLL_PrepCall
  'Call system_log("VQCALC Entry")
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("AirFlowRate =" & Str$(AirFlowRate))
  'Call system_log("WaterFlowRate =" & Str$(WaterFlowRate))
  Call Fortran_VQCALC(AirToWaterRatio, AirFlowRate, WaterFlowRate)
  'Call system_log("VQCALC Exit")
End Sub

Sub VQMINBUB(MinAirToWaterRatio As Double, Influent As Double, TreatmentObjective As Double, HenrysConstant As Double, NumberOfTanks As Long)
  Call DLL_PrepCall
  'Call system_log("VQMINBUB Entry")
  'Call system_log("MinAirToWaterRatio =" & Str$(MinAirToWaterRatio))
  'Call system_log("Influent =" & Str$(Influent))
  'Call system_log("TreatmentObjective =" & Str$(TreatmentObjective))
  'Call system_log("HenrysConstant =" & Str$(HenrysConstant))
  'Call system_log("NumberOfTanks =" & Str$(NumberOfTanks))
  Call Fortran_VQMINBUB(MinAirToWaterRatio, Influent, TreatmentObjective, HenrysConstant, NumberOfTanks)
  'Call system_log("VQMINBUB Exit")
End Sub

Sub vqmltpt1(AirToWaterRatio As Double, MinimumAirToWaterRatio As Double, MultipleOfMinimumAirToWaterRatio As Double)
  Call DLL_PrepCall
  'Call system_log("vqmltpt1 Entry")
  'Call system_log("AirToWaterRatio =" & Str$(AirToWaterRatio))
  'Call system_log("MinimumAirToWaterRatio =" & Str$(MinimumAirToWaterRatio))
  'Call system_log("MultipleOfMinimumAirToWaterRatio =" & Str$(MultipleOfMinimumAirToWaterRatio))
  Call Fortran_VQMLTPT1(AirToWaterRatio, MinimumAirToWaterRatio, MultipleOfMinimumAirToWaterRatio)
  'Call system_log("vqmltpt1 Exit")
End Sub


