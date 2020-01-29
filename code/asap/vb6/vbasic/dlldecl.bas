Attribute VB_Name = "DLL_Decl_Mod"
Option Explicit

'      '
'      ' DLL DECLARATIONS FOR PACKED TOWER AERATION (ASAPPTAD.DLL).
'      ' NOTE: (*) INDICATES THIS ROUTINE IS PRESENT IN THE DLL BUT IS NOT CALLED BY THE ASAP VISUAL BASIC CODE.
'      '
'      ' .......... NON-ALIASED DLL ROUTINES:
'      Declare Sub Fortran_AIRDENS Lib "asapptad.dll" Alias "AIRDENS" (AirDensity As Double, Temperature As Double, Pressure As Double)
'      Declare Sub Fortran_AIRVISC Lib "asapptad.dll" Alias "AIRVISC" (AirViscosity As Double, Temperature As Double)
'      Declare Sub Fortran_AREAPT2 Lib "asapptad.dll" Alias "AREAPT2" (TowerArea As Double, TowerDiameter As Double)
'      'DIFFL -- (*)
'      'DIFGWL -- (*)
'      Declare Sub Fortran_GETSAF Lib "asapptad.dll" Alias "GETSAF" (KLaSafetyFactor As Double, OndaMassTransferCoefficient As Double, DesignMassTransferCoefficient As Double)
'      Declare Sub Fortran_H2ODENS Lib "asapptad.dll" Alias "H2ODENS" (LiquidDensity As Double, Temperature As Double)
'      Declare Sub Fortran_H2OST Lib "asapptad.dll" Alias "H2OST" (LiquidSurfaceTension As Double, Temperature As Double)
'      Declare Sub Fortran_H2OVISC Lib "asapptad.dll" Alias "H2OVISC" (LiquidViscosity As Double, Temperature As Double)
'      Declare Sub Fortran_LDAIRPT2 Lib "asapptad.dll" Alias "LDAIRPT2" (AirLoadingRate As Double, AirFlowRate As Double, AirDensity As Double, TowerArea As Double)
'      Declare Sub Fortran_LDH2OPT2 Lib "asapptad.dll" Alias "LDH2OPT2" (WaterLoadingRate As Double, WaterFlowRate As Double, WaterDensity As Double, TowerArea As Double)
'      Declare Sub Fortran_OPTMAL Lib "asapptad.dll" Alias "OPTMAL" (WaterDensity As Double, WaterViscosity As Double, WaterSurfaceTension As Double, AirDensity As Double, AirViscosity As Double, WaterFlowRate As Double, PackingNominalSize As Double, PackingFactor As Double, PackingCriticalSurfaceTension As Double, PackingSpecificSurfaceArea As Double, InfluentConcentrations As Double, TreatmentObjectives As Double, HenrysConstants As Double, NumberOfContaminants As Long, PressureDrop As Double, LiquidDiffusivities As Double, GasDiffusivities As Double, KLaSafetyFactor As Double, ID_OptimalDesignContaminant As Long, MultipleOfMinimumAirToWaterRatio As Double, EffluentConcentrations As Double, ErrorFlag As Long)
'      Declare Sub Fortran_PBLOWPT Lib "asapptad.dll" Alias "PBLOWPT" (BlowerBrakePower As Double, AirFlowRate As Double, TowerArea As Double, OperatingPressure As Double, PressureDrop As Double, TowerHeight As Double, AirDensity As Double, InletAirTemperature As Double, BlowerEfficiency As Double)
'      Declare Sub Fortran_PDROP Lib "asapptad.dll" Alias "PDROP" (AirPressureDrop As Double, AirToWaterRatio As Double, AirLoadingRate As Double, PackingFactor As Double, WaterViscosity As Double, AirDensity As Double, WaterDensity As Double, InitialPressureDrop As Double, MaximumPressureDrop As Double, PressureDropStep As Double)
'      Declare Sub Fortran_PPUMPPT Lib "asapptad.dll" Alias "PPUMPPT" (PumpBrakePower As Double, PumpEfficiency As Double, WaterDensity As Double, WaterFlowRate As Double, TowerHeight As Double)
'      Declare Sub Fortran_PTOTALPT Lib "asapptad.dll" Alias "PTOTALPT" (TotalBrakePower As Double, BlowerBrakePower As Double, PumpBrakePower As Double)
'      Declare Sub Fortran_QAIRPT2 Lib "asapptad.dll" Alias "QAIRPT2" (AirFlowRate As Double, AirLoadingRate As Double, AirDensity As Double, TowerArea As Double)
'      Declare Sub Fortran_QH2OPT2 Lib "asapptad.dll" Alias "QH2OPT2" (WaterFlowRate As Double, WaterLoadingRate As Double, WaterDensity As Double, TowerArea As Double)
'      Declare Sub Fortran_REMOVPT Lib "asapptad.dll" Alias "REMOVPT" (RemovalEfficiency As Double, InfluentConcentration As Double, Effluent As Double)
'      Declare Sub Fortran_TVOLPT2 Lib "asapptad.dll" Alias "TVOLPT2" (TowerVolume As Double, TowerArea As Double, TowerLength As Double)
'      Declare Sub Fortran_VQCALC Lib "asapptad.dll" Alias "VQCALC" (AirToWaterRatio As Double, AirFlowRate As Double, WaterFlowRate As Double)
'      ' .......... ALIASED DLL ROUTINES:
'      Declare Sub Fortran_AIRFLO Lib "asapptad.dll" Alias "_AIRFLO@12" (AirFlowRate As Double, AirToWaterRatio As Double, WaterFlowRate As Double)
'      Declare Sub Fortran_AWCALC Lib "asapptad.dll" Alias "_AWCALC@40" (PackingWettedSurfaceArea As Double, PackingCriticalSurfaceTension As Double, WaterSurfaceTension As Double, WaterLoadingRate As Double, PackingSpecificSurfaceArea As Double, WaterViscosity As Double, WaterDensity As Double, ReynoldsNumber As Double, FroudeNumber As Double, WeberNumber As Double)
'      'DIFLHL -- (*)
'      'DIFLPOL -- (*)
'      Declare Sub Fortran_EFFLPT2 Lib "asapptad.dll" Alias "_EFFLPT2@32" (EffluentConcentration As Double, AirToWaterRatio As Double, HenrysConstant As Double, WaterFlowRate As Double, TowerArea As Double, TowerLength As Double, DesignMassTransferCoefficient As Double, InfluentConcentration As Double)
'      'FINDKLA -- (*)
'      ''''Declare Sub Fortran_GETCSPT Lib "asapptad.dll" Alias "_GETCSPT@20" (DesignContaminantAirWaterInterfaceConc As Double, AirToWaterRatio As Double, DesignContaminantHenrysConstant As Double, DesignContaminantInfluentConcentration As Double, DesignContaminantTreatmentObjective As Double)
'      Declare Sub Fortran_GETCSPT Lib "asapptad" Alias "_GETCSPT@20" (DesignContaminantAirWaterInterfaceConc As Double, AirToWaterRatio As Double, DesignContaminantHenrysConstant As Double, DesignContaminantInfluentConcentration As Double, DesignContaminantTreatmentObjective As Double)
'      'GETHIVQ -- (*)
'      Declare Sub Fortran_GETHTUPT Lib "asapptad.dll" Alias "_GETHTUPT@16" (TransferUnitHeight As Double, WaterFlowRate As Double, TowerArea As Double, DesignMassTransferCoefficient As Double)
'      Declare Sub Fortran_GETMULT Lib "asapptad.dll" Alias "_GETMULT@12" (MultipleOfMinimumAirToWaterRatio As Double, AirToWaterRatio As Double, MinimumAirToWaterRatio As Double)
'      Declare Sub Fortran_GETNTUPT Lib "asapptad.dll" Alias "_GETNTUPT@16" (NumberOfTransferUnits As Double, DesignContaminantInfluentConcentration As Double, DesignContaminantTreatmentObjective As Double, DesignContaminantAirToWaterInterfaceConc As Double)
'      Declare Sub Fortran_KLACOR Lib "asapptad.dll" Alias "_KLACOR@12" (DesignMassTransferCoefficient As Double, OndaMassTransferCoefficient As Double, KLaSafetyFactor As Double)
'      Declare Sub Fortran_ONDAKGPT Lib "asapptad.dll" Alias "_ONDAKGPT@28" (OndaGasPhaseMassTransferCoefficient As Double, AirLoadingRate As Double, PackingSpecificSurfaceArea As Double, AirViscosity As Double, AirDensity As Double, GasDiffusivity As Double, PackingNominalSize As Double)
'      Declare Sub Fortran_ONDAKLPT Lib "asapptad.dll" Alias "_ONDAKLPT@32" (OndaLiquidPhaseMassTransferCoefficient As Double, WaterLoadingRate As Double, PackingWettedSurfaceArea As Double, WaterViscosity As Double, WaterDensity As Double, LiquidDiffusivity As Double, PackingSpecificSurfaceArea As Double, PackingNominalSize As Double)
'      Declare Sub Fortran_ONDKLAPT Lib "asapptad.dll" Alias "_ONDKLAPT@32" (OndaOverallMassTransferCoefficient As Double, OndaLiquidPhaseResistance As Double, OndaGasPhaseResistance As Double, OndaTotalResistance As Double, OndaLiquidPhaseMassTransferCoefficient As Double, PackingWettedSurfaceArea As Double, OndaGasPhaseMassTransferCoefficient As Double, DesignContaminantHenrysConstant As Double)
'      Declare Sub Fortran_PT1AREA Lib "asapptad.dll" Alias "_PT1AREA@16" (TowerArea As Double, WaterFlowRate As Double, WaterDensity As Double, WaterMassLoadingRate As Double)
'      Declare Sub Fortran_PT1DTOW Lib "asapptad.dll" Alias "_PT1DTOW@8" (TowerDiameter As Double, TowerArea As Double)
'      Declare Sub Fortran_PT1HTOW Lib "asapptad.dll" Alias "_PT1HTOW@12" (TowerHeight As Double, TransferUnitHeight As Double, NumberOfTransferUnits As Double)
'      Declare Sub Fortran_PT1LDAIR Lib "asapptad.dll" Alias "_PT1LDAIR@28" (AirMassLoadingRate As Double, AirPressureDrop As Double, AirToWaterRatio As Double, AirDensity As Double, WaterDensity As Double, PackingFactor As Double, WaterViscosity As Double)
'      Declare Sub Fortran_PT1LDH2O Lib "asapptad.dll" Alias "_PT1LDH2O@20" (WaterMassLoadingRate As Double, AirToWaterRatio As Double, AirDensity As Double, WaterDensity As Double, AirMassLoadingRate As Double)
'      Declare Sub Fortran_PT1TVOL Lib "asapptad.dll" Alias "_PT1TVOL@12" (TowerVolume As Double, TowerArea As Double, TowerHeight As Double)
'      Declare Sub Fortran_PT1VQMIN Lib "asapptad.dll" Alias "_PT1VQMIN@16" (MinimumAirToWaterRatio As Double, InfluentConcentration As Double, TreatmentObjective As Double, HenrysConstant As Double)
'      Declare Sub Fortran_VQMLTPT1 Lib "asapptad.dll" Alias "_VQMLTPT1@12" (AirToWaterRatio As Double, MinimumAirToWaterRatio As Double, MultipleOfMinimumAirToWaterRatio As Double)
'
'
'      '
'      ' DLL DECLARATIONS FOR BUBBLE AERATION (ASAPBUB.DLL).
'      ' NOTE: (*) INDICATES THIS ROUTINE IS PRESENT IN THE DLL BUT IS NOT CALLED BY THE ASAP VISUAL BASIC CODE.
'      '
'      ' .......... NON-ALIASED DLL ROUTINES:
'      'Declare Sub airflo Lib "asapbub.dll" (AirFlowRate As Double, AirToWaterRatio As Double, WaterFlowRate As Double)
'      'AIRFLO -- (*)
'      Declare Sub Fortran_DIFO2 Lib "asapbub.dll" Alias "DIFO2" (DiffusivityOxygen As Double, Temperature As Double)
'      Declare Sub Fortran_EFFLBUB Lib "asapbub.dll" Alias "EFFLBUB" (ArrayLiqPhaseEffluentConc As Double, ArrayGasPhaseEffluentConc As Double, HenrysConstOfCompound As Double, LiqPhaseInfluentConc As Double, AirToWaterRatio As Double, NoOfTanks As Long, StantonNo As Double)
'      Declare Sub Fortran_GETCSTAR Lib "asapbub.dll" Alias "GETCSTAR" (DOSaturationConc As Double, WeightDensityWater As Double, EffectiveSaturationDepth As Double, BarometricPressure As Double, WaterDepth As Double)
'      Declare Sub Fortran_GETPHIB Lib "asapbub.dll" Alias "GETPHIB" (StantonNo As Double, CompoundMassTransCoeff As Double, VolumeOfEaTank As Double, HenrysConstOfCompound As Double, AirFlowRateToEaTank As Double)
'      Declare Sub Fortran_GETSOTE Lib "asapbub.dll" Alias "GETSOTE" (StandardOxygenTransferEff As Double, StandardOxygenTransferRate As Double, AirFlowRate As Double)
'      Declare Sub Fortran_GETSOTR Lib "asapbub.dll" Alias "GETSOTR" (StandardOxygenTransferRate As Double, StandardOxygenTransferEff As Double, AirFlowRate As Double)
'      Declare Sub Fortran_KLA20A Lib "asapbub.dll" Alias "KLA20A" (AppOxygenMassTransCoeff As Double, WaterVolumePerTankL As Double, WaterVolumePerTankm3 As Double, DOSaturationConcentration As Double, StandOxygenMassTransRate As Double)
'      Declare Sub Fortran_KLABUB Lib "asapbub.dll" Alias "KLABUB" (CompoundMassTransferCoeff As Double, OxygenMassTransferCoeff As Double, DiffusivityLiquidWater As Double, DiffusivityOfOxygen As Double, ExponentInCorrelation As Double, RatioGasLiquidTransfer As Double, HenrysConstant As Double)
'      Declare Sub Fortran_PCALCBUB Lib "asapbub.dll" Alias "PCALCBUB" (TotalBrakePowerAllTanks As Double, BlowerBrakePowerForEaTank As Double, OperatingPressure As Double, InletAirTempC As Double, AirFlowRate As Double, BlowerEfficiencyPercent As Double, LiquidDensity As Double, WaterDepth As Double, NoOfTanks As Long, NumberOfBlowersinEachTank As Long)
'      Declare Sub Fortran_REMOVBUB Lib "asapbub.dll" Alias "REMOVBUB" (ActualLiqPhaseRemovalEfficiency As Double, LiqPhaseInfluentConc As Double, LiqPhaseEffluentConcLastTank As Double)
'      Declare Sub Fortran_TAUSVOLS Lib "asapbub.dll" Alias "TAUSVOLS" (TotalFluidResidenceTime As Double, NoOfTanksInSeries As Long, HydraulicRetentTimeOfEaTank As Double, VolumeOfEachTank As Double, TotalVolumeOfAllTanks As Double, WaterFlowRate As Double, TankParametersCode As Long)
'      ''''Declare Sub Fortran_TrueKLa Lib "asapbub.dll" Alias "TrueKLa" (OxygenMassTransferCoeffOperatTemp As Double, AppOxygenMassTransferCoeff20Deg As Double, ParameterUsedInKla As Double, AirFlowRate As Double, WaterVolPerTankL As Double, BarometricPressure As Double, WeightDensityWater As Double, TrueOxygenMassTransferCoef20 As Double, EffectiveSaturatonDepth As Double, OperatingTemp As Double)
'      Declare Sub Fortran_TrueKLa Lib "asapbub.dll" Alias "TRUEKLA" (OxygenMassTransferCoeffOperatTemp As Double, AppOxygenMassTransferCoeff20Deg As Double, ParameterUsedInKla As Double, AirFlowRate As Double, WaterVolPerTankL As Double, BarometricPressure As Double, WeightDensityWater As Double, TrueOxygenMassTransferCoef20 As Double, EffectiveSaturatonDepth As Double, OperatingTemp As Double)
'      Declare Sub Fortran_VOLBUB Lib "asapbub.dll" Alias "VOLBUB" (TankVolume As Double, HenrysConstant As Double, AirFlowRate As Double, ContaminantMassTransferCoeff As Double, Influent As Double, TreatmentObjective As Double, NumberOfTanks As Long, WaterFlowRate As Double, ErrorFlag As Long)
'      Declare Sub Fortran_VQBUB Lib "asapbub.dll" Alias "VQBUB" (AirToWaterRatio As Double, AirFlowRateToEachTank As Double, WaterFlowRate As Double)
'      Declare Sub Fortran_VQMINBUB Lib "asapbub.dll" Alias "VQMINBUB" (MinAirToWaterRatio As Double, Influent As Double, TreatmentObjective As Double, HenrysConstant As Double, NumberOfTanks As Long)
'      ' .......... ALIASED DLL ROUTINES:
'      '
'      ' (THERE ARE NO ALIASED DLL ROUTINES IN THE ASAPBUB.DLL FILE.)
'
'
'      '
'      ' DLL DECLARATIONS FOR SURFACE AERATION (ASAPSURF.DLL).
'      ' NOTE: (*) INDICATES THIS ROUTINE IS PRESENT IN THE DLL BUT IS NOT CALLED BY THE ASAP VISUAL BASIC CODE.
'      '
'      ' .......... NON-ALIASED DLL ROUTINES:
'      'DIFO2 -- (*)
'      Declare Sub Fortran_KLAO2SUR Lib "asapsurf.dll" Alias "KLAO2SUR" (OxygenMTCoeff As Double, PowerInput_PoverV As Double)
'      Declare Sub Fortran_KLASURF Lib "asapsurf.dll" Alias "KLASURF" (ContaminantMassTransferCoeff As Double, OxygenMassTransferCoeff As Double, ContaminantLiquidDiffusivity As Double, OxygenLiquidDiffusivity As Double, N_forFindingKLa As Double, kgOVERkl_forFindingKLa As Double, HenrysConstant As Double)
'      Declare Sub Fortran_PCALCSUR Lib "asapsurf.dll" Alias "PCALCSUR" (TotalPower As Double, PowerPerTank As Double, PowerInput_PoverV As Double, TotalVolumeAllTanks As Double, NumberOfTanks As Long, AeratorMotorEfficiency As Double)
'      Declare Sub Fortran_SEFFL Lib "asapsurf.dll" Alias "SEFFL" (EffluentConcentrations As Double, AchievedRemovalEfficiency As Double, Influent As Double, ContaminantMassTransferCoeff As Double, TankResidenceTime As Double, NumberOfTanks As Long)
'      Declare Sub Fortran_SURFEFF Lib "asapsurf.dll" Alias "SURFEFF" (RemovalEfficiency As Double, Influent As Double, EffluentOrTreatmentObjective As Double)
'      Declare Sub Fortran_TAUISURF Lib "asapsurf.dll" Alias "TAUISURF" (TankResidenceTime As Double, Influent As Double, TreatmentObjective As Double, NumberOfTanks As Long, ContaminantMassTransferCoeff As Double)
'      ' .......... ALIASED DLL ROUTINES:
'      '
'      ' (THERE ARE NO ALIASED DLL ROUTINES IN THE ASAPBUB.DLL FILE.)





'
' DLL DECLARATIONS FOR PACKED TOWER AERATION (DLLS\ASAPPTAD.DLL).
' NOTE: (*) INDICATES THIS ROUTINE IS PRESENT IN THE DLL BUT IS NOT CALLED BY THE ASAP VISUAL BASIC CODE.
'
' .......... NON-ALIASED DLL ROUTINES:
Declare Sub Fortran_AIRDENS Lib "dlls\asapptad.dll" Alias "AIRDENS" (AirDensity As Double, Temperature As Double, Pressure As Double)
Declare Sub Fortran_AIRVISC Lib "dlls\asapptad.dll" Alias "AIRVISC" (AirViscosity As Double, Temperature As Double)
Declare Sub Fortran_AREAPT2 Lib "dlls\asapptad.dll" Alias "AREAPT2" (TowerArea As Double, TowerDiameter As Double)
'DIFFL -- (*)
'DIFGWL -- (*)
Declare Sub Fortran_GETSAF Lib "dlls\asapptad.dll" Alias "GETSAF" (KLaSafetyFactor As Double, OndaMassTransferCoefficient As Double, DesignMassTransferCoefficient As Double)
Declare Sub Fortran_H2ODENS Lib "dlls\asapptad.dll" Alias "H2ODENS" (LiquidDensity As Double, Temperature As Double)
Declare Sub Fortran_H2OST Lib "dlls\asapptad.dll" Alias "H2OST" (LiquidSurfaceTension As Double, Temperature As Double)
Declare Sub Fortran_H2OVISC Lib "dlls\asapptad.dll" Alias "H2OVISC" (LiquidViscosity As Double, Temperature As Double)
Declare Sub Fortran_LDAIRPT2 Lib "dlls\asapptad.dll" Alias "LDAIRPT2" (AirLoadingRate As Double, AirFlowRate As Double, AirDensity As Double, TowerArea As Double)
Declare Sub Fortran_LDH2OPT2 Lib "dlls\asapptad.dll" Alias "LDH2OPT2" (WaterLoadingRate As Double, WaterFlowRate As Double, WaterDensity As Double, TowerArea As Double)
Declare Sub Fortran_OPTMAL Lib "dlls\asapptad.dll" Alias "OPTMAL" (WaterDensity As Double, WaterViscosity As Double, WaterSurfaceTension As Double, AirDensity As Double, AirViscosity As Double, WaterFlowRate As Double, PackingNominalSize As Double, PackingFactor As Double, PackingCriticalSurfaceTension As Double, PackingSpecificSurfaceArea As Double, InfluentConcentrations As Double, TreatmentObjectives As Double, HenrysConstants As Double, NumberOfContaminants As Long, PressureDrop As Double, LiquidDiffusivities As Double, GasDiffusivities As Double, KLaSafetyFactor As Double, ID_OptimalDesignContaminant As Long, MultipleOfMinimumAirToWaterRatio As Double, EffluentConcentrations As Double, ErrorFlag As Long)
Declare Sub Fortran_PBLOWPT Lib "dlls\asapptad.dll" Alias "PBLOWPT" (BlowerBrakePower As Double, AirFlowRate As Double, TowerArea As Double, OperatingPressure As Double, PressureDrop As Double, TowerHeight As Double, AirDensity As Double, InletAirTemperature As Double, BlowerEfficiency As Double)
Declare Sub Fortran_PDROP Lib "dlls\asapptad.dll" Alias "PDROP" (AirPressureDrop As Double, AirToWaterRatio As Double, AirLoadingRate As Double, PackingFactor As Double, WaterViscosity As Double, AirDensity As Double, WaterDensity As Double, InitialPressureDrop As Double, MaximumPressureDrop As Double, PressureDropStep As Double)
Declare Sub Fortran_PPUMPPT Lib "dlls\asapptad.dll" Alias "PPUMPPT" (PumpBrakePower As Double, PumpEfficiency As Double, WaterDensity As Double, WaterFlowRate As Double, TowerHeight As Double)
Declare Sub Fortran_PTOTALPT Lib "dlls\asapptad.dll" Alias "PTOTALPT" (TotalBrakePower As Double, BlowerBrakePower As Double, PumpBrakePower As Double)
Declare Sub Fortran_QAIRPT2 Lib "dlls\asapptad.dll" Alias "QAIRPT2" (AirFlowRate As Double, AirLoadingRate As Double, AirDensity As Double, TowerArea As Double)
Declare Sub Fortran_QH2OPT2 Lib "dlls\asapptad.dll" Alias "QH2OPT2" (WaterFlowRate As Double, WaterLoadingRate As Double, WaterDensity As Double, TowerArea As Double)
Declare Sub Fortran_REMOVPT Lib "dlls\asapptad.dll" Alias "REMOVPT" (RemovalEfficiency As Double, InfluentConcentration As Double, Effluent As Double)
Declare Sub Fortran_TVOLPT2 Lib "dlls\asapptad.dll" Alias "TVOLPT2" (TowerVolume As Double, TowerArea As Double, TowerLength As Double)
Declare Sub Fortran_VQCALC Lib "dlls\asapptad.dll" Alias "VQCALC" (AirToWaterRatio As Double, AirFlowRate As Double, WaterFlowRate As Double)
' .......... ALIASED DLL ROUTINES:
Declare Sub Fortran_AIRFLO Lib "dlls\asapptad.dll" Alias "_AIRFLO@12" (AirFlowRate As Double, AirToWaterRatio As Double, WaterFlowRate As Double)
Declare Sub Fortran_AWCALC Lib "dlls\asapptad.dll" Alias "_AWCALC@40" (PackingWettedSurfaceArea As Double, PackingCriticalSurfaceTension As Double, WaterSurfaceTension As Double, WaterLoadingRate As Double, PackingSpecificSurfaceArea As Double, WaterViscosity As Double, WaterDensity As Double, ReynoldsNumber As Double, FroudeNumber As Double, WeberNumber As Double)
'DIFLHL -- (*)
'DIFLPOL -- (*)
Declare Sub Fortran_EFFLPT2 Lib "dlls\asapptad.dll" Alias "_EFFLPT2@32" (EffluentConcentration As Double, AirToWaterRatio As Double, HenrysConstant As Double, WaterFlowRate As Double, TowerArea As Double, TowerLength As Double, DesignMassTransferCoefficient As Double, InfluentConcentration As Double)
'FINDKLA -- (*)
Declare Sub Fortran_GETCSPT Lib "dlls\asapptad.dll" Alias "_GETCSPT@20" (DesignContaminantAirWaterInterfaceConc As Double, AirToWaterRatio As Double, DesignContaminantHenrysConstant As Double, DesignContaminantInfluentConcentration As Double, DesignContaminantTreatmentObjective As Double)
'GETHIVQ -- (*)
Declare Sub Fortran_GETHTUPT Lib "dlls\asapptad.dll" Alias "_GETHTUPT@16" (TransferUnitHeight As Double, WaterFlowRate As Double, TowerArea As Double, DesignMassTransferCoefficient As Double)
Declare Sub Fortran_GETMULT Lib "dlls\asapptad.dll" Alias "_GETMULT@12" (MultipleOfMinimumAirToWaterRatio As Double, AirToWaterRatio As Double, MinimumAirToWaterRatio As Double)
Declare Sub Fortran_GETNTUPT Lib "dlls\asapptad.dll" Alias "_GETNTUPT@16" (NumberOfTransferUnits As Double, DesignContaminantInfluentConcentration As Double, DesignContaminantTreatmentObjective As Double, DesignContaminantAirToWaterInterfaceConc As Double)
Declare Sub Fortran_KLACOR Lib "dlls\asapptad.dll" Alias "_KLACOR@12" (DesignMassTransferCoefficient As Double, OndaMassTransferCoefficient As Double, KLaSafetyFactor As Double)
Declare Sub Fortran_ONDAKGPT Lib "dlls\asapptad.dll" Alias "_ONDAKGPT@28" (OndaGasPhaseMassTransferCoefficient As Double, AirLoadingRate As Double, PackingSpecificSurfaceArea As Double, AirViscosity As Double, AirDensity As Double, GasDiffusivity As Double, PackingNominalSize As Double)
Declare Sub Fortran_ONDAKLPT Lib "dlls\asapptad.dll" Alias "_ONDAKLPT@32" (OndaLiquidPhaseMassTransferCoefficient As Double, WaterLoadingRate As Double, PackingWettedSurfaceArea As Double, WaterViscosity As Double, WaterDensity As Double, LiquidDiffusivity As Double, PackingSpecificSurfaceArea As Double, PackingNominalSize As Double)
Declare Sub Fortran_ONDKLAPT Lib "dlls\asapptad.dll" Alias "_ONDKLAPT@32" (OndaOverallMassTransferCoefficient As Double, OndaLiquidPhaseResistance As Double, OndaGasPhaseResistance As Double, OndaTotalResistance As Double, OndaLiquidPhaseMassTransferCoefficient As Double, PackingWettedSurfaceArea As Double, OndaGasPhaseMassTransferCoefficient As Double, DesignContaminantHenrysConstant As Double)
Declare Sub Fortran_PT1AREA Lib "dlls\asapptad.dll" Alias "_PT1AREA@16" (TowerArea As Double, WaterFlowRate As Double, WaterDensity As Double, WaterMassLoadingRate As Double)
Declare Sub Fortran_PT1DTOW Lib "dlls\asapptad.dll" Alias "_PT1DTOW@8" (TowerDiameter As Double, TowerArea As Double)
Declare Sub Fortran_PT1HTOW Lib "dlls\asapptad.dll" Alias "_PT1HTOW@12" (TowerHeight As Double, TransferUnitHeight As Double, NumberOfTransferUnits As Double)
Declare Sub Fortran_PT1LDAIR Lib "dlls\asapptad.dll" Alias "_PT1LDAIR@28" (AirMassLoadingRate As Double, AirPressureDrop As Double, AirToWaterRatio As Double, AirDensity As Double, WaterDensity As Double, PackingFactor As Double, WaterViscosity As Double)
Declare Sub Fortran_PT1LDH2O Lib "dlls\asapptad.dll" Alias "_PT1LDH2O@20" (WaterMassLoadingRate As Double, AirToWaterRatio As Double, AirDensity As Double, WaterDensity As Double, AirMassLoadingRate As Double)
Declare Sub Fortran_PT1TVOL Lib "dlls\asapptad.dll" Alias "_PT1TVOL@12" (TowerVolume As Double, TowerArea As Double, TowerHeight As Double)
Declare Sub Fortran_PT1VQMIN Lib "dlls\asapptad.dll" Alias "_PT1VQMIN@16" (MinimumAirToWaterRatio As Double, InfluentConcentration As Double, TreatmentObjective As Double, HenrysConstant As Double)
Declare Sub Fortran_VQMLTPT1 Lib "dlls\asapptad.dll" Alias "_VQMLTPT1@12" (AirToWaterRatio As Double, MinimumAirToWaterRatio As Double, MultipleOfMinimumAirToWaterRatio As Double)


'
' DLL DECLARATIONS FOR BUBBLE AERATION (DLLS\ASAPBUB.DLL).
' NOTE: (*) INDICATES THIS ROUTINE IS PRESENT IN THE DLL BUT IS NOT CALLED BY THE ASAP VISUAL BASIC CODE.
'
' .......... NON-ALIASED DLL ROUTINES:
'Declare Sub airflo Lib "dlls\asapbub.dll" (AirFlowRate As Double, AirToWaterRatio As Double, WaterFlowRate As Double)
'AIRFLO -- (*)
Declare Sub Fortran_DIFO2 Lib "dlls\asapbub.dll" Alias "DIFO2" (DiffusivityOxygen As Double, Temperature As Double)
Declare Sub Fortran_EFFLBUB Lib "dlls\asapbub.dll" Alias "EFFLBUB" (ArrayLiqPhaseEffluentConc As Double, ArrayGasPhaseEffluentConc As Double, HenrysConstOfCompound As Double, LiqPhaseInfluentConc As Double, AirToWaterRatio As Double, NoOfTanks As Long, StantonNo As Double)
Declare Sub Fortran_GETCSTAR Lib "dlls\asapbub.dll" Alias "GETCSTAR" (DOSaturationConc As Double, WeightDensityWater As Double, EffectiveSaturationDepth As Double, BarometricPressure As Double, WaterDepth As Double)
Declare Sub Fortran_GETPHIB Lib "dlls\asapbub.dll" Alias "GETPHIB" (StantonNo As Double, CompoundMassTransCoeff As Double, VolumeOfEaTank As Double, HenrysConstOfCompound As Double, AirFlowRateToEaTank As Double)
Declare Sub Fortran_GETSOTE Lib "dlls\asapbub.dll" Alias "GETSOTE" (StandardOxygenTransferEff As Double, StandardOxygenTransferRate As Double, AirFlowRate As Double)
Declare Sub Fortran_GETSOTR Lib "dlls\asapbub.dll" Alias "GETSOTR" (StandardOxygenTransferRate As Double, StandardOxygenTransferEff As Double, AirFlowRate As Double)
Declare Sub Fortran_KLA20A Lib "dlls\asapbub.dll" Alias "KLA20A" (AppOxygenMassTransCoeff As Double, WaterVolumePerTankL As Double, WaterVolumePerTankm3 As Double, DOSaturationConcentration As Double, StandOxygenMassTransRate As Double)
Declare Sub Fortran_KLABUB Lib "dlls\asapbub.dll" Alias "KLABUB" (CompoundMassTransferCoeff As Double, OxygenMassTransferCoeff As Double, DiffusivityLiquidWater As Double, DiffusivityOfOxygen As Double, ExponentInCorrelation As Double, RatioGasLiquidTransfer As Double, HenrysConstant As Double)
Declare Sub Fortran_PCALCBUB Lib "dlls\asapbub.dll" Alias "PCALCBUB" (TotalBrakePowerAllTanks As Double, BlowerBrakePowerForEaTank As Double, OperatingPressure As Double, InletAirTempC As Double, AirFlowRate As Double, BlowerEfficiencyPercent As Double, LiquidDensity As Double, WaterDepth As Double, NoOfTanks As Long, NumberOfBlowersinEachTank As Long)
Declare Sub Fortran_REMOVBUB Lib "dlls\asapbub.dll" Alias "REMOVBUB" (ActualLiqPhaseRemovalEfficiency As Double, LiqPhaseInfluentConc As Double, LiqPhaseEffluentConcLastTank As Double)
Declare Sub Fortran_TAUSVOLS Lib "dlls\asapbub.dll" Alias "TAUSVOLS" (TotalFluidResidenceTime As Double, NoOfTanksInSeries As Long, HydraulicRetentTimeOfEaTank As Double, VolumeOfEachTank As Double, TotalVolumeOfAllTanks As Double, WaterFlowRate As Double, TankParametersCode As Long)
''''Declare Sub Fortran_TrueKLa Lib "dlls\asapbub.dll" Alias "TrueKLa" (OxygenMassTransferCoeffOperatTemp As Double, AppOxygenMassTransferCoeff20Deg As Double, ParameterUsedInKla As Double, AirFlowRate As Double, WaterVolPerTankL As Double, BarometricPressure As Double, WeightDensityWater As Double, TrueOxygenMassTransferCoef20 As Double, EffectiveSaturatonDepth As Double, OperatingTemp As Double)
Declare Sub Fortran_TrueKLa Lib "dlls\asapbub.dll" Alias "TRUEKLA" (OxygenMassTransferCoeffOperatTemp As Double, AppOxygenMassTransferCoeff20Deg As Double, ParameterUsedInKla As Double, AirFlowRate As Double, WaterVolPerTankL As Double, BarometricPressure As Double, WeightDensityWater As Double, TrueOxygenMassTransferCoef20 As Double, EffectiveSaturatonDepth As Double, OperatingTemp As Double)
Declare Sub Fortran_VOLBUB Lib "dlls\asapbub.dll" Alias "VOLBUB" (TankVolume As Double, HenrysConstant As Double, AirFlowRate As Double, ContaminantMassTransferCoeff As Double, Influent As Double, TreatmentObjective As Double, NumberOfTanks As Long, WaterFlowRate As Double, ErrorFlag As Long)
Declare Sub Fortran_VQBUB Lib "dlls\asapbub.dll" Alias "VQBUB" (AirToWaterRatio As Double, AirFlowRateToEachTank As Double, WaterFlowRate As Double)
Declare Sub Fortran_VQMINBUB Lib "dlls\asapbub.dll" Alias "VQMINBUB" (MinAirToWaterRatio As Double, Influent As Double, TreatmentObjective As Double, HenrysConstant As Double, NumberOfTanks As Long)
' .......... ALIASED DLL ROUTINES:
'
' (THERE ARE NO ALIASED DLL ROUTINES IN THE DLLS\ASAPBUB.DLL FILE.)


'
' DLL DECLARATIONS FOR SURFACE AERATION (DLLS\ASAPSURF.DLL).
' NOTE: (*) INDICATES THIS ROUTINE IS PRESENT IN THE DLL BUT IS NOT CALLED BY THE ASAP VISUAL BASIC CODE.
'
' .......... NON-ALIASED DLL ROUTINES:
'DIFO2 -- (*)
Declare Sub Fortran_KLAO2SUR Lib "dlls\asapsurf.dll" Alias "KLAO2SUR" (OxygenMTCoeff As Double, PowerInput_PoverV As Double)
Declare Sub Fortran_KLASURF Lib "dlls\asapsurf.dll" Alias "KLASURF" (ContaminantMassTransferCoeff As Double, OxygenMassTransferCoeff As Double, ContaminantLiquidDiffusivity As Double, OxygenLiquidDiffusivity As Double, N_forFindingKLa As Double, kgOVERkl_forFindingKLa As Double, HenrysConstant As Double)
Declare Sub Fortran_PCALCSUR Lib "dlls\asapsurf.dll" Alias "PCALCSUR" (TotalPower As Double, PowerPerTank As Double, PowerInput_PoverV As Double, TotalVolumeAllTanks As Double, NumberOfTanks As Long, AeratorMotorEfficiency As Double)
Declare Sub Fortran_SEFFL Lib "dlls\asapsurf.dll" Alias "SEFFL" (EffluentConcentrations As Double, AchievedRemovalEfficiency As Double, Influent As Double, ContaminantMassTransferCoeff As Double, TankResidenceTime As Double, NumberOfTanks As Long)
Declare Sub Fortran_SURFEFF Lib "dlls\asapsurf.dll" Alias "SURFEFF" (RemovalEfficiency As Double, Influent As Double, EffluentOrTreatmentObjective As Double)
Declare Sub Fortran_TAUISURF Lib "dlls\asapsurf.dll" Alias "TAUISURF" (TankResidenceTime As Double, Influent As Double, TreatmentObjective As Double, NumberOfTanks As Long, ContaminantMassTransferCoeff As Double)
' .......... ALIASED DLL ROUTINES:
'
' (THERE ARE NO ALIASED DLL ROUTINES IN THE DLLS\ASAPBUB.DLL FILE.)




