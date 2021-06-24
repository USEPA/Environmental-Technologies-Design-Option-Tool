Attribute VB_Name = "DLL_Decl_Mod"
Option Explicit

'
' This module contains the DLL declarations needed to make
' calls to FORTRAN routines within the program
'


'
' DLL DECLARATIONS FOR MAIN STEPP FORTRAN CODE (DLLS\STEPP.DLL).
' THESE ROUTINES PERFORM PROPERTY CALCULATIONS.
' NOTE: (*) INDICATES THIS ROUTINE IS PRESENT IN THE DLL BUT IS NOT CALLED BY THE StEPP VISUAL BASIC CODE.
'
' .......... NON-ALIASED DLL ROUTINES:
'Alias "GETSAF"
Declare Sub ACCALL Lib "dlls\stepp.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTEMP As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)
Declare Sub AIRDENS Lib "dlls\stepp.dll" (AirDensity As Double, OperatingTemp As Double, OperatingPressure As Double, AirDensityError As Long, AirDensityShortSource As Long, AirDensityLongSource As Long, AirDensityTemperature As Double)
Declare Sub AIRVISC Lib "dlls\stepp.dll" (AirViscosity As Double, OperatingTemp As Double, AirViscosityError As Long, AirViscosityShortSource As Long, AirViscosityLongSource As Long, AirViscosityTemperature As Double)
Declare Sub AQSCALL Lib "dlls\stepp.dll" (AqueousSolubilty As Double, AqueousSolubiltyShortSource As Long, AqueousSolubiltyLongSource As Long, AqueousSolubiltyError As Long, AqueousSolubiltyTemp As Double, CalculationTemperature As Double, MaximumUnifacGroups As Long, MS As Long, XMW As Double, BinaryInteractionParameterDatabase As Long)
Declare Sub AQSFIT Lib "dlls\stepp.dll" (AqueousSolubilityFit As Double, AqSolFitShortSource As Long, AqSolFitLongSource As Long, AqSolFitError As Long, AqSolFitTemp As Double, AqSolUnifacDBT As Double, AQSolUnifacDBTTmp As Double, AqSolUnifacOpT As Double, AqSolDatabase As Double, AqSolDatabaseTemp As Double, OperatingTemp As Double)
Declare Sub DIFGWL Lib "dlls\stepp.dll" (DIFGWilkeLee As Double, MolecularWeight As Double, MolarVolumeNBP As Double, BoilingPointTemp As Double, OperatingTemp As Double, OperatingPressure As Double, DIFGWilkeLeeError As Long, DIFGWilkeLeeShortSource As Long, DIFGWilkeLeeLongSource As Long, DIFGWilkeLeeTemp As Double)
Declare Sub DIFLHL Lib "dlls\stepp.dll" (DIFLHaydukLaudie As Double, MolarVolumeNBP As Double, OperatingTemp As Double, MolecularWeight As Double, DIFLHaydukLaudieError As Long, DIFLHaydukLaudieSourceShort As Long, DIFLHaydukLaudieSourceLong As Long, DIFLHaydukLaudieTemp As Double)
Declare Sub DIFLPOL Lib "dlls\stepp.dll" (DIFLPolson As Double, MolecularWeight As Double, DIFLPolsonError As Long, DIFLPolsonShortSource As Long, DIFLPolsonLongSource As Long, DIFLPolsonTemp As Double, OperatingTemp As Double)
Declare Sub DIFLWC Lib "dlls\stepp.dll" (DIFLWilkeChang As Double, MolarVolumeNBP As Double, OperatingTemperature As Double, DIFLWilkeChangError As Long, DIFLWilkeChangShortSource As Long, DIFLWilkeChangLongSource As Long, DIFLWilkeChangTemp As Double)
Declare Sub H2OST Lib "dlls\stepp.dll" (WaterSurfaceTension As Double, OperatingTemp As Double, WaterSurfaceTensionError As Long, WaterSurfaceTensionShortSource As Long, WaterSurfaceTensionLongSource As Long, WaterSurfaceTensionTemp As Double)
Declare Sub HC1CALL Lib "dlls\stepp.dll" (HenryCUNIFAC As Double, HCUnifacShortSource As Long, HCUnifacLongSource As Long, HCUnifacError As Long, HCUnifacTemp As Double, OperatingTemp As Double, ActivityCoefficient As Double, VaporPressure As Double)
Declare Sub HC2CALL Lib "dlls\stepp.dll" (HenryCRegress As Double, HCRegressShortSource As Long, HCRegressLongSource As Long, HCRegressError As Long, HCRegressTemp As Double, HCDatabase As Double, HCDatabaseTemp As Double, OperatingTemp As Double, NumberOfDatabaseHCs As Long)
Declare Sub HCDBCONV Lib "dlls\stepp.dll" (HenrysConstantDatabase As Double, HenrysConstantDatabaseTemp As Double, NumberOfDatabaseHenrysConstants As Long, HCShortSource As Long)
Declare Sub HENFIT Lib "dlls\stepp.dll" (HenrysConstantFit As Double, HCFitShortSource As Long, HCFitLongSource As Long, HCFitError As Long, HCFitTemperature As Double, HCDatabase As Double, HCDatabaseTemp As Double, HCUnifacOpT As Double, HCUnifacDBTs As Double, HCUnifacDBTsErrors As Long, OperatingTemp As Double, NumberDBHenryCs As Long)
Declare Sub KOWCALL Lib "dlls\stepp.dll" (Kow As Double, KowShortSource As Long, KowLongSource As Long, KowError As Long, KowTemp As Double, CalculationTemperature As Double, FGRPErrorFlag As Long, MaximumUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)
Declare Sub LDDBCALL Lib "dlls\stepp.dll" (LiquidDensityDB As Double, LiquidDensityDBShortSource As Long, LiquidDensityDBLongSource As Long, LiquidDensityDBError As Long, LiquidDensityDBEqn As Long, LiquidDensityDBTemp As Double, LiquidDensityDBMinimumT As Double, LiquidDensityDBMaximumT As Double, LiquidDensityDBCoeffA As Double, LiquidDensityDBCoeffB As Double, LiquidDensityDBCoeffC As Double, LiquidDensityDBCoeffD As Double, MolecularWeight As Double, OperatingTemperature As Double)
Declare Sub LDGCCALL Lib "dlls\stepp.dll" (LiquidDensityGC As Double, LiquidDensityGCShortSource As Long, LiquidDensityGCLongSource As Long, LiquidDensityGCError As Long, LiquidDensityGCTemp As Double, MolecularWeight As Double, MolarVolumeNBP As Double, OperatingTemperature As Double)
Declare Sub MWTCALL Lib "dlls\stepp.dll" (MWtUnifac As Double, MWtUnifacShortSource As Long, MWtUnifacLongSource As Long, MWtUnifacError As Long, MaximumUnifacGroups As Long, MS As Long, XMW As Double)
Declare Sub VBBPCALL Lib "dlls\stepp.dll" (MolarVolumeNBP As Double, MolarVolumeNBPShortSource As Long, MolarVolumeNBPLongSource As Long, MolarVolumeNBPError As Long, MolarVolumeNBPTemp As Double, BoilingPointTemp As Double, MaximumUnifacGroups As Long, MS As Long, NumberOfRings As Long)
Declare Sub VBMATT Lib "dlls\stepp.dll" (MolarVolumeOpT As Double, LiquidDensity As Double, MolecularWeight As Double)
Declare Sub VPRCALL Lib "dlls\stepp.dll" (VaporPressure As Double, VPShortSource As Long, VPLongSource As Long, VPError As Long, VPEquation As Long, VPTempDB As Double, VPMinimumT As Double, VPMaximumT As Double, VPCoeffA As Double, VPCoeffB As Double, VPCoeffC As Double, VPCoeffD As Double, VPCoeffE As Double, VPSuperfund As Double, VPSuperfundTemp As Double, OperatingTemp As Double)
' .......... ALIASED DLL ROUTINES:
'AQSOL -- (*)
'BINPAR -- (*)
'DBDENS -- (*)
'ERROR -- (*)
'FGRP -- (*)
'FGRPCALL -- (*)
'GETGAM -- (*)
Declare Sub H2ODENS Lib "dlls\stepp.dll" Alias "_H2ODENS@24" (WaterDensity As Double, OperatingTemperature As Double, WaterDensityError As Long, WaterDensityShortSource As Long, WaterDensityLongSource As Long, WaterDensityTemp As Double)
Declare Sub H2OVISC Lib "dlls\stepp.dll" Alias "_H2OVISC@24" (WaterViscosity As Double, OperatingTemp As Double, WaterViscosityError As Long, WaterViscosityShortSource As Long, WaterViscosityLongSource As Long, WaterViscosityTemp As Double)
'HENRY -- (*)
'INITVS -- (*)
'MOLWT -- (*)
'NEWTON -- (*)
'ORGDENS -- (*)
'PARMS -- (*)
'PARTC -- (*)
'REGRESS -- (*)
'UNIMOD -- (*)
'VAPORP -- (*)
'VBMSCH -- (*)



'
' DLL DECLARATIONS FOR UNIT CONVERSION FORTRAN CODE (DLLS\STEPPCONV.DLL).
' THESE ROUTINES PERFORM CONVERSIONS FROM SI-TO-ENGLISH AND FROM ENGLISH-TO-SI.
' NOTE: (*) INDICATES THIS ROUTINE IS PRESENT IN THE DLL BUT IS NOT CALLED BY THE StEPP VISUAL BASIC CODE.
'
' ROUTINES NAMED *CONV OR *CNV ARE SI-TO-ENGLISH CONVERTORS.
' ROUTINES NAMED *SI ARE ENGLISH-TO-SI CONVERTORS.
'
' .......... NON-ALIASED DLL ROUTINES:
Declare Sub ACCONV Lib "dlls\stepconv.dll" (ActivityCoefficientEnglishUnits As Double, ActivityCoefficientSIUnits As Double)
Declare Sub ACENSI Lib "dlls\stepconv.dll" (ActivityCoefficientSIUnits As Double, ActivityCoefficientEnglishUnits As Double)
Declare Sub ADENENSI Lib "dlls\stepconv.dll" (AirDensitySIUnits As Double, AirDensityEnglishUnits As Double)
Declare Sub ADENSCNV Lib "dlls\stepconv.dll" (AirDensityEnglishUnits As Double, AirDensitySIUnits As Double)
Declare Sub AQSCONV Lib "dlls\stepconv.dll" (AqueousSolubilityEnglishUnits As Double, AqueousSolubilitySIUnits As Double)
Declare Sub AQSENSI Lib "dlls\stepconv.dll" (AqueousSolubilitySIUnits As Double, AqueousSolubilityEnglishUnits As Double)
Declare Sub AVISCCNV Lib "dlls\stepconv.dll" (AirViscosityEnglishUnits As Double, AirViscositySIUnits As Double)
Declare Sub AVISENSI Lib "dlls\stepconv.dll" (AirViscositySIUnits As Double, AirViscosityEnglishUnits As Double)
Declare Sub GDIFENSI Lib "dlls\stepconv.dll" (GasDiffusivitySIUnits As Double, GasDiffusivityEnglishUnits As Double)
Declare Sub GDIFFCNV Lib "dlls\stepconv.dll" (GasDiffusivityEnglishUnits As Double, GasDiffusivitySIUnits As Double)
Declare Sub H2OSTCNV Lib "dlls\stepconv.dll" (WaterSurfTensionEnglishUnits As Double, WaterSurfTensionSIUnits As Double)
Declare Sub HCCONV Lib "dlls\stepconv.dll" (HenrysConstantEnglishUnits As Double, HenrysConstantSIUnits As Double)
Declare Sub HCENSI Lib "dlls\stepconv.dll" (HenrysConstantSIUnits As Double, HenrysConstantEnglishUnits As Double)
Declare Sub KOWCONV Lib "dlls\stepconv.dll" (OctWaterPartCoeffEnglishUnits As Double, OctWaterPartCoeffSIUnits As Double)
Declare Sub KOWENSI Lib "dlls\stepconv.dll" (OctWaterPartCoeffSIUnits As Double, OctWaterPartCoeffEnglishUnits As Double)
Declare Sub LDENENSI Lib "dlls\stepconv.dll" (LiquidDensitySIUnits As Double, LiquidDensityEnglishUnits As Double)
Declare Sub LDENSCNV Lib "dlls\stepconv.dll" (LiquidDensityEnglishUnits As Double, LiquidDensitySIUnits As Double)
Declare Sub LDIFENSI Lib "dlls\stepconv.dll" (LiquidDiffusivitySIUnits As Double, LiquidDiffusivityEnglishUnits As Double)
Declare Sub LDIFFCNV Lib "dlls\stepconv.dll" (LiquidDiffusivityEnglishUnits As Double, LiquidDiffusivitySIUnits As Double)
Declare Sub MVBPENSI Lib "dlls\stepconv.dll" (MolarVolumeNBPSIUnits As Double, MolarVolumeNBPEnglishUnits As Double)
Declare Sub MVNBPCNV Lib "dlls\stepconv.dll" (MolarVolumeNBPEnglishUnits As Double, MolarVolumeNBPSIUnits As Double)
Declare Sub MVOPTCNV Lib "dlls\stepconv.dll" (MolarVolumeOpTEnglishUnits As Double, MolarVolumeOpTSIUnits As Double)
Declare Sub MVOTENSI Lib "dlls\stepconv.dll" (MolarVolumeOpTSIUnits As Double, MolarVolumeOpTEnglishUnits As Double)
Declare Sub MWCONV Lib "dlls\stepconv.dll" (MolecularWeightEnglishUnits As Double, MolecularWeightSIUnits As Double)
Declare Sub MWENSI Lib "dlls\stepconv.dll" (MolecularWeightSIUnits As Double, MolecularWeightEnglishUnits As Double)
Declare Sub NBPCONV Lib "dlls\stepconv.dll" (BoilingPointEnglishUnits As Double, BoilingPointSIUnits As Double)
Declare Sub NBPENSI Lib "dlls\stepconv.dll" (BoilingPointSIUnits As Double, BoilingPointEnglishUnits As Double)
Declare Sub PRESENSI Lib "dlls\stepconv.dll" (PressureSIUnits As Double, PressureEnglishUnits As Double)
Declare Sub PRESSCNV Lib "dlls\stepconv.dll" (PressureEnglishUnits As Double, PressureSIUnits As Double)
Declare Sub RICONV Lib "dlls\stepconv.dll" (RefractiveIndexEnglishUnits As Double, RefractiveIndexSIUnits As Double)
Declare Sub RIENSI Lib "dlls\stepconv.dll" (RefractiveIndexSIUnits As Double, RefractiveIndexEnglishUnits As Double)
Declare Sub TEMPCNV Lib "dlls\stepconv.dll" (TemperatureEnglishUnits As Double, TemperatureSIUnits As Double)
Declare Sub TEMPENSI Lib "dlls\stepconv.dll" (TemperatureSIUnits As Double, TemperatureEnglishUnits As Double)
Declare Sub VPCONV Lib "dlls\stepconv.dll" (VaporPressureEnglishUnits As Double, VaporPressureSIUnits As Double)
Declare Sub VPENSI Lib "dlls\stepconv.dll" (VaporPressureSIUnits As Double, VaporPressureEnglishUnits As Double)
Declare Sub WDENENSI Lib "dlls\stepconv.dll" (WaterDensitySIUnits As Double, WaterDensityEnglishUnits As Double)
Declare Sub WDENSCNV Lib "dlls\stepconv.dll" (WaterDensityEnglishUnits As Double, WaterDensitySIUnits As Double)
Declare Sub WSTENSI Lib "dlls\stepconv.dll" (WaterSurfTensionSIUnits As Double, WaterSurfTensionEnglishUnits As Double)
Declare Sub WVISCCNV Lib "dlls\stepconv.dll" (WaterViscosityEnglishUnits As Double, WaterViscositySIUnits As Double)
Declare Sub WVISENSI Lib "dlls\stepconv.dll" (WaterViscositySIUnits As Double, WaterViscosityEnglishUnits As Double)
' .......... ALIASED DLL ROUTINES:
' (NO ALIASED DLL ROUTINES PRESENT.)






'
'
' OLD DLL DECLARATION LIST. =======================
'
'
'
''DLL Declarations from StEPP.dll
''   DLL Declarations needed to calculate the properties
'
'Declare Sub VPRCALL Lib "dlls\stepp.dll" (VaporPressure As Double, VPShortSource As Long, VPLongSource As Long, VPError As Long, VPEquation As Long, VPTempDB As Double, VPMinimumT As Double, VPMaximumT As Double, VPCoeffA As Double, VPCoeffB As Double, VPCoeffC As Double, VPCoeffD As Double, VPCoeffE As Double, VPSuperfund As Double, VPSuperfundTemp As Double, OperatingTemp As Double)
'Declare Sub ACCALL Lib "dlls\stepp.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTEMP As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)
'Declare Sub HC1CALL Lib "dlls\stepp.dll" (HenryCUNIFAC As Double, HCUnifacShortSource As Long, HCUnifacLongSource As Long, HCUnifacError As Long, HCUnifacTemp As Double, OperatingTemp As Double, ActivityCoefficient As Double, VaporPressure As Double)
'Declare Sub HC2CALL Lib "dlls\stepp.dll" (HenryCRegress As Double, HCRegressShortSource As Long, HCRegressLongSource As Long, HCRegressError As Long, HCRegressTemp As Double, HCDatabase As Double, HCDatabaseTemp As Double, OperatingTemp As Double, NumberOfDatabaseHCs As Long)
'Declare Sub MWTCALL Lib "dlls\stepp.dll" (MWtUnifac As Double, MWtUnifacShortSource As Long, MWtUnifacLongSource As Long, MWtUnifacError As Long, MaximumUnifacGroups As Long, MS As Long, XMW As Double)
'Declare Sub VBBPCALL Lib "dlls\stepp.dll" (MolarVolumeNBP As Double, MolarVolumeNBPShortSource As Long, MolarVolumeNBPLongSource As Long, MolarVolumeNBPError As Long, MolarVolumeNBPTemp As Double, BoilingPointTemp As Double, MaximumUnifacGroups As Long, MS As Long, NumberOfRings As Long)
'Declare Sub LDDBCALL Lib "dlls\stepp.dll" (LiquidDensityDB As Double, LiquidDensityDBShortSource As Long, LiquidDensityDBLongSource As Long, LiquidDensityDBError As Long, LiquidDensityDBEqn As Long, LiquidDensityDBTemp As Double, LiquidDensityDBMinimumT As Double, LiquidDensityDBMaximumT As Double, LiquidDensityDBCoeffA As Double, LiquidDensityDBCoeffB As Double, LiquidDensityDBCoeffC As Double, LiquidDensityDBCoeffD As Double, MolecularWeight As Double, OperatingTemperature As Double)
'Declare Sub LDGCCALL Lib "dlls\stepp.dll" (LiquidDensityGC As Double, LiquidDensityGCShortSource As Long, LiquidDensityGCLongSource As Long, LiquidDensityGCError As Long, LiquidDensityGCTemp As Double, MolecularWeight As Double, MolarVolumeNBP As Double, OperatingTemperature As Double)
'Declare Sub VBMATT Lib "dlls\stepp.dll" (MolarVolumeOpT As Double, LiquidDensity As Double, MolecularWeight As Double)
'Declare Sub AQSCALL Lib "dlls\stepp.dll" (AqueousSolubilty As Double, AqueousSolubiltyShortSource As Long, AqueousSolubiltyLongSource As Long, AqueousSolubiltyError As Long, AqueousSolubiltyTemp As Double, CalculationTemperature As Double, MaximumUnifacGroups As Long, MS As Long, XMW As Double, BinaryInteractionParameterDatabase As Long)
'Declare Sub KOWCALL Lib "dlls\stepp.dll" (Kow As Double, KowShortSource As Long, KowLongSource As Long, KowError As Long, KowTemp As Double, CalculationTemperature As Double, FGRPErrorFlag As Long, MaximumUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)
'Declare Sub DIFLHL Lib "dlls\stepp.dll" (DIFLHaydukLaudie As Double, MolarVolumeNBP As Double, OperatingTemp As Double, MolecularWeight As Double, DIFLHaydukLaudieError As Long, DIFLHaydukLaudieSourceShort As Long, DIFLHaydukLaudieSourceLong As Long, DIFLHaydukLaudieTemp As Double)
'Declare Sub DIFLPOL Lib "dlls\stepp.dll" (DIFLPolson As Double, MolecularWeight As Double, DIFLPolsonError As Long, DIFLPolsonShortSource As Long, DIFLPolsonLongSource As Long, DIFLPolsonTemp As Double, OperatingTemp As Double)
'Declare Sub DIFLWC Lib "dlls\stepp.dll" (DIFLWilkeChang As Double, MolarVolumeNBP As Double, OperatingTemperature As Double, DIFLWilkeChangError As Long, DIFLWilkeChangShortSource As Long, DIFLWilkeChangLongSource As Long, DIFLWilkeChangTemp As Double)
'Declare Sub DIFGWL Lib "dlls\stepp.dll" (DIFGWilkeLee As Double, MolecularWeight As Double, MolarVolumeNBP As Double, BoilingPointTemp As Double, OperatingTemp As Double, OperatingPressure As Double, DIFGWilkeLeeError As Long, DIFGWilkeLeeShortSource As Long, DIFGWilkeLeeLongSource As Long, DIFGWilkeLeeTemp As Double)
'Declare Sub H2ODENS Lib "dlls\stepp.dll" (WaterDensity As Double, OperatingTemperature As Double, WaterDensityError As Long, WaterDensityShortSource As Long, WaterDensityLongSource As Long, WaterDensityTemp As Double)
'Declare Sub H2OVISC Lib "dlls\stepp.dll" (WaterViscosity As Double, OperatingTemp As Double, WaterViscosityError As Long, WaterViscosityShortSource As Long, WaterViscosityLongSource As Long, WaterViscosityTemp As Double)
'Declare Sub H2OST Lib "dlls\stepp.dll" (WaterSurfaceTension As Double, OperatingTemp As Double, WaterSurfaceTensionError As Long, WaterSurfaceTensionShortSource As Long, WaterSurfaceTensionLongSource As Long, WaterSurfaceTensionTemp As Double)
'Declare Sub AIRDENS Lib "dlls\stepp.dll" (AirDensity As Double, OperatingTemp As Double, OperatingPressure As Double, AirDensityError As Long, AirDensityShortSource As Long, AirDensityLongSource As Long, AirDensityTemperature As Double)
'Declare Sub AIRVISC Lib "dlls\stepp.dll" (AirViscosity As Double, OperatingTemp As Double, AirViscosityError As Long, AirViscosityShortSource As Long, AirViscosityLongSource As Long, AirViscosityTemperature As Double)
'Declare Sub HENFIT Lib "dlls\stepp.dll" (HenrysConstantFit As Double, HCFitShortSource As Long, HCFitLongSource As Long, HCFitError As Long, HCFitTemperature As Double, HCDatabase As Double, HCDatabaseTemp As Double, HCUnifacOpT As Double, HCUnifacDBTs As Double, HCUnifacDBTsErrors As Long, OperatingTemp As Double, NumberDBHenryCs As Long)
'Declare Sub AQSFIT Lib "dlls\stepp.dll" (AqueousSolubilityFit As Double, AqSolFitShortSource As Long, AqSolFitLongSource As Long, AqSolFitError As Long, AqSolFitTemp As Double, AqSolUnifacDBT As Double, AQSolUnifacDBTTmp As Double, AqSolUnifacOpT As Double, AqSolDatabase As Double, AqSolDatabaseTemp As Double, OperatingTemp As Double)
'Declare Sub HCDBCONV Lib "dlls\stepp.dll" (HenrysConstantDatabase As Double, HenrysConstantDatabaseTemp As Double, NumberOfDatabaseHenrysConstants As Long, HCShortSource As Long)
'
'
''DLL Declarations from STEPPCONV.dll
''   DLL Declarations needed to perform unit conversions
'
''SI to English Units
'Declare Sub VPCONV Lib "dlls\stepconv.dll" (VaporPressureEnglishUnits As Double, VaporPressureSIUnits As Double)
'Declare Sub ACCONV Lib "dlls\stepconv.dll" (ActivityCoefficientEnglishUnits As Double, ActivityCoefficientSIUnits As Double)
'Declare Sub HCCONV Lib "dlls\stepconv.dll" (HenrysConstantEnglishUnits As Double, HenrysConstantSIUnits As Double)
'Declare Sub MWCONV Lib "dlls\stepconv.dll" (MolecularWeightEnglishUnits As Double, MolecularWeightSIUnits As Double)
'Declare Sub LDENSCNV Lib "dlls\stepconv.dll" (LiquidDensityEnglishUnits As Double, LiquidDensitySIUnits As Double)
'Declare Sub MVOPTCNV Lib "dlls\stepconv.dll" (MolarVolumeOpTEnglishUnits As Double, MolarVolumeOpTSIUnits As Double)
'Declare Sub MVNBPCNV Lib "dlls\stepconv.dll" (MolarVolumeNBPEnglishUnits As Double, MolarVolumeNBPSIUnits As Double)
'Declare Sub NBPCONV Lib "dlls\stepconv.dll" (BoilingPointEnglishUnits As Double, BoilingPointSIUnits As Double)
'Declare Sub RICONV Lib "dlls\stepconv.dll" (RefractiveIndexEnglishUnits As Double, RefractiveIndexSIUnits As Double)
'Declare Sub AQSCONV Lib "dlls\stepconv.dll" (AqueousSolubilityEnglishUnits As Double, AqueousSolubilitySIUnits As Double)
'Declare Sub KOWCONV Lib "dlls\stepconv.dll" (OctWaterPartCoeffEnglishUnits As Double, OctWaterPartCoeffSIUnits As Double)
'Declare Sub LDIFFCNV Lib "dlls\stepconv.dll" (LiquidDiffusivityEnglishUnits As Double, LiquidDiffusivitySIUnits As Double)
'Declare Sub GDIFFCNV Lib "dlls\stepconv.dll" (GasDiffusivityEnglishUnits As Double, GasDiffusivitySIUnits As Double)
'Declare Sub WDENSCNV Lib "dlls\stepconv.dll" (WaterDensityEnglishUnits As Double, WaterDensitySIUnits As Double)
'Declare Sub WVISCCNV Lib "dlls\stepconv.dll" (WaterViscosityEnglishUnits As Double, WaterViscositySIUnits As Double)
'Declare Sub H2OSTCNV Lib "dlls\stepconv.dll" (WaterSurfTensionEnglishUnits As Double, WaterSurfTensionSIUnits As Double)
'Declare Sub ADENSCNV Lib "dlls\stepconv.dll" (AirDensityEnglishUnits As Double, AirDensitySIUnits As Double)
'Declare Sub AVISCCNV Lib "dlls\stepconv.dll" (AirViscosityEnglishUnits As Double, AirViscositySIUnits As Double)
'Declare Sub PRESSCNV Lib "dlls\stepconv.dll" (PressureEnglishUnits As Double, PressureSIUnits As Double)
'Declare Sub TEMPCNV Lib "dlls\stepconv.dll" (TemperatureEnglishUnits As Double, TemperatureSIUnits As Double)
'
''English to SI Units
'Declare Sub VPENSI Lib "dlls\stepconv.dll" (VaporPressureSIUnits As Double, VaporPressureEnglishUnits As Double)
'Declare Sub ACENSI Lib "dlls\stepconv.dll" (ActivityCoefficientSIUnits As Double, ActivityCoefficientEnglishUnits As Double)
'Declare Sub HCENSI Lib "dlls\stepconv.dll" (HenrysConstantSIUnits As Double, HenrysConstantEnglishUnits As Double)
'Declare Sub MWENSI Lib "dlls\stepconv.dll" (MolecularWeightSIUnits As Double, MolecularWeightEnglishUnits As Double)
'Declare Sub LDENENSI Lib "dlls\stepconv.dll" (LiquidDensitySIUnits As Double, LiquidDensityEnglishUnits As Double)
'Declare Sub MVOTENSI Lib "dlls\stepconv.dll" (MolarVolumeOpTSIUnits As Double, MolarVolumeOpTEnglishUnits As Double)
'Declare Sub MVBPENSI Lib "dlls\stepconv.dll" (MolarVolumeNBPSIUnits As Double, MolarVolumeNBPEnglishUnits As Double)
'Declare Sub NBPENSI Lib "dlls\stepconv.dll" (BoilingPointSIUnits As Double, BoilingPointEnglishUnits As Double)
'Declare Sub RIENSI Lib "dlls\stepconv.dll" (RefractiveIndexSIUnits As Double, RefractiveIndexEnglishUnits As Double)
'Declare Sub AQSENSI Lib "dlls\stepconv.dll" (AqueousSolubilitySIUnits As Double, AqueousSolubilityEnglishUnits As Double)
'Declare Sub KOWENSI Lib "dlls\stepconv.dll" (OctWaterPartCoeffSIUnits As Double, OctWaterPartCoeffEnglishUnits As Double)
'Declare Sub LDIFENSI Lib "dlls\stepconv.dll" (LiquidDiffusivitySIUnits As Double, LiquidDiffusivityEnglishUnits As Double)
'Declare Sub GDIFENSI Lib "dlls\stepconv.dll" (GasDiffusivitySIUnits As Double, GasDiffusivityEnglishUnits As Double)
'Declare Sub WDENENSI Lib "dlls\stepconv.dll" (WaterDensitySIUnits As Double, WaterDensityEnglishUnits As Double)
'Declare Sub WVISENSI Lib "dlls\stepconv.dll" (WaterViscositySIUnits As Double, WaterViscosityEnglishUnits As Double)
'Declare Sub WSTENSI Lib "dlls\stepconv.dll" (WaterSurfTensionSIUnits As Double, WaterSurfTensionEnglishUnits As Double)
'Declare Sub ADENENSI Lib "dlls\stepconv.dll" (AirDensitySIUnits As Double, AirDensityEnglishUnits As Double)
'Declare Sub AVISENSI Lib "dlls\stepconv.dll" (AirViscositySIUnits As Double, AirViscosityEnglishUnits As Double)
'Declare Sub PRESENSI Lib "dlls\stepconv.dll" (PressureSIUnits As Double, PressureEnglishUnits As Double)
'Declare Sub TEMPENSI Lib "dlls\stepconv.dll" (TemperatureSIUnits As Double, TemperatureEnglishUnits As Double)

