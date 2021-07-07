Attribute VB_Name = "WriteMod"
Option Explicit
' THIS MODULE HAS ALL OF THE FUNCTIONS IN IT THAT RELATE
' TO READING AND WRITING OF THE MAIN DEFAULT FILE.

' CURRENT METHOD..   0 = Packed Tower   1 = Bubble  2 = Surface
Global CurrMethod%

' CURRENT MODE        0 = DESIGN   1 = RATING
Global CurrMode%

Sub WriteBubbleType(fnum As Integer, buf As BubbleType)
Dim i%

Call WriteBubInfoType(fnum, buf.OperatingPressure)
Call WriteBubInfoType(fnum, buf.operatingtemperature)
Call WriteBubInfoType(fnum, buf.WaterDensity)
Call WriteBubInfoType(fnum, buf.WaterViscosity)
Call WriteBubInfoType(fnum, buf.N_for_Finding_KLa)
Call WriteBubInfoType(fnum, buf.kgOVERkl_for_Finding_KLa)
Call WriteBubInfoType(fnum, buf.ContaminantMassTransferCoefficient)
Call WriteBubInfoType(fnum, buf.WaterFlowRate)
Call WriteBubInfoType(fnum, buf.MinimumAirToWaterRatio)
Call WriteBubInfoType(fnum, buf.AirToWaterRatio)
Call WriteBubInfoType(fnum, buf.AirFlowRate)
Call WriteBubInfoType(fnum, buf.TankHydraulicRetentionTime)
Call WriteBubInfoType(fnum, buf.TotalHydraulicRetentionTime)
Call WriteBubInfoType(fnum, buf.TankVolume)
Call WriteBubInfoType(fnum, buf.TotalTankVolume)
Call WriteBubInfoType(fnum, buf.StantonNumber)

Put #fnum, , buf.NumberOfTanks.value
Put #fnum, , buf.NumberOfTanks.UserInput
Put #fnum, , buf.NumberOfTanks.ValChanged

Put #fnum, , buf.CodeForTausAndTankVolumes
Put #fnum, , buf.DesiredPercentRemoval
Put #fnum, , buf.AchievedPercentRemoval
Put #fnum, , buf.ID_OptimalDesignContaminant

Put #fnum, , buf.Power.BlowerBrakePower
Put #fnum, , buf.Power.TotalBrakePower
Put #fnum, , buf.Power.InletAirTemperature
Put #fnum, , buf.Power.BlowerEfficiency
Put #fnum, , buf.Power.TankWaterDepth
Put #fnum, , buf.Power.NumberOfBlowersinEachTank

Put #fnum, , buf.Chemical

Put #fnum, , buf.NumChemical
For i% = 1 To buf.NumChemical
    Call WriteBubContamProperty(fnum, buf.Contaminant(i%))
Next i%
Call WriteBubContamProperty(fnum, buf.DesignContaminant)

Put #fnum, , buf.Oxygen.KLaMethod
Call WriteBubInfoType(fnum, buf.Oxygen.LiquidDiffusivity)
Call WriteBubInfoType(fnum, buf.Oxygen.MassTransferCoefficient)
Call WriteBubInfoType(fnum, buf.Oxygen.CWO2TestData.SOTR)
Call WriteBubInfoType(fnum, buf.Oxygen.CWO2TestData.SOTE)
Call WriteBubInfoType(fnum, buf.Oxygen.CWO2TestData.AirFlowRate_QAIR)
Call WriteBubInfoType(fnum, buf.Oxygen.CWO2TestData.BarometricPressure_PB)
Call WriteBubInfoType(fnum, buf.Oxygen.CWO2TestData.WaterDepth_DEPTHW)
Call WriteBubInfoType(fnum, buf.Oxygen.CWO2TestData.WaterVolumePerTank_VM3)
Put #fnum, , buf.Oxygen.CWO2TestData.DOSaturationConc_CSTR20
Put #fnum, , buf.Oxygen.CWO2TestData.WeightDensityOfWater_GAMMAW
Put #fnum, , buf.Oxygen.CWO2TestData.EffectiveSaturationDepth_DEFF
Put #fnum, , buf.Oxygen.CWO2TestData.ApparentOxygenMTCoeff_KLA20
Put #fnum, , buf.Oxygen.CWO2TestData.WaterVolumePerTankLiters_V
Put #fnum, , buf.Oxygen.CWO2TestData.TrueKLaAt20DegC_KLAT20
Put #fnum, , buf.Oxygen.CWO2TestData.Phi
Put #fnum, , buf.Oxygen.CWO2TestData.TrueOxygenMTCoeffOperatingT_KLAO2

End Sub

Sub WriteBubContamProperty(fnum As Integer, buf As BubbleContaminantPropertyType)
    Dim strsize%, i%

strsize% = Len(buf.Name)
Put #fnum, , strsize%
Put #fnum, , buf.Name
Put #fnum, , buf.Pressure
Put #fnum, , buf.Temperature

For i% = 0 To MAXIMUM_TANKS
    Put #fnum, , buf.Effluent(i%)
Next i%

For i% = 1 To MAXIMUM_TANKS
    Put #fnum, , buf.GasEffluent(i%)
Next i%

Call WriteBubInfoType(fnum, buf.MolecularWeight)
Call WriteBubInfoType(fnum, buf.HenrysConstant)
Call WriteBubInfoType(fnum, buf.MolarVolume)
Call WriteBubInfoType(fnum, buf.LiquidDiffusivity)
Call WriteBubInfoType(fnum, buf.Influent)
Call WriteBubInfoType(fnum, buf.TreatmentObjective)
End Sub

Sub WriteBubInfoType(fnum As Integer, buf As BubbleInformationType)

Put #fnum, , buf.value
Put #fnum, , buf.UserInput
Put #fnum, , buf.ValChanged

End Sub

Sub WriteOndaMassTransCoef(fnum As Integer, buf As OndaMassTransferCoefficientType)

Put #fnum, , buf.ReynoldsNumber
Put #fnum, , buf.FroudeNumber
Put #fnum, , buf.WeberNumber
Put #fnum, , buf.LiquidPhaseMassTransferResistance
Put #fnum, , buf.GasPhaseMassTransferResistance
Put #fnum, , buf.TotalMassTransferResistance
Put #fnum, , buf.LiquidPhaseMassTransferCoefficient
Put #fnum, , buf.GasPhaseMassTransferCoefficient
Put #fnum, , buf.OverallMassTransferCoefficient
Put #fnum, , buf.ValChanged

End Sub

Sub WriteProgramState(Filename$)
    Dim StateFileID$
    Dim strsize%


If Filename$ = "" Then Exit Sub
StateFileID$ = "Stepp Data File"


Open Filename$ For Binary As #1
strsize% = Len(StateFileID$)
Put #1, 1, strsize%
Put #1, 1, StateFileID$

Select Case CurrMethod%
    Case 0 ' ***** SAVE PACKED TOWER DESIGN MODE ******
        Call WriteSCRType(1, scr1)
        Call WriteSCRType(1, Scr2)

    Case 1 ' ***** SAVE BUBBLE TOWER RATING MODE ******
        Call WriteBubbleType(1, bub)
'        Call WriteBubbleType(1, bub(CurrMode%))

    Case 2 ' ***** SAVE SURFACE TOWER RATING MODE ******
        Call WriteSurfaceType(1, sur)
'        Call WriteSurfaceType(1, sur(CurrMode%))

End Select


Close #1

End Sub

Sub WritePTADInfoType(fnum As Integer, buf As PTADInformationType)

Put #fnum, , buf.value
Put #fnum, , buf.ValChanged
Put #fnum, , buf.UserInput

End Sub

Sub WritePTContamProperty(fnum As Integer, buf As ContaminantPropertyType)
    Dim strsize%

strsize% = Len(buf.Name)
Put #fnum, , strsize%
Put #fnum, , buf.Name
Put #fnum, , buf.Pressure
Put #fnum, , buf.Temperature
Put #fnum, , buf.AirWaterInterfaceConcentration
Call WritePTADInfoType(fnum, buf.MolecularWeight)
Call WritePTADInfoType(fnum, buf.HenrysConstant)
Call WritePTADInfoType(fnum, buf.MolarVolume)
Call WritePTADInfoType(fnum, buf.NormalBoilingPoint)
Call WritePTADInfoType(fnum, buf.LiquidDiffusivity)
Call WritePTADInfoType(fnum, buf.GasDiffusivity)
Call WritePTADInfoType(fnum, buf.Influent)
Call WritePTADInfoType(fnum, buf.TreatmentObjective)
Call WritePTADInfoType(fnum, buf.Effluent)

End Sub

Sub WriteSCRType(fnum As Integer, buf As SCR)
  Dim i%

Call WritePTADInfoType(fnum, buf.OperatingPressure)
Call WritePTADInfoType(fnum, buf.operatingtemperature)
Call WritePTADInfoType(fnum, buf.WaterFlowRate)
Call WritePTADInfoType(fnum, buf.WaterDensity)
Call WritePTADInfoType(fnum, buf.WaterViscosity)
Call WritePTADInfoType(fnum, buf.WaterSurfaceTension)
Call WritePTADInfoType(fnum, buf.WaterLoadingRate)
Call WritePTADInfoType(fnum, buf.AirDensity)
Call WritePTADInfoType(fnum, buf.AirViscosity)
Call WritePTADInfoType(fnum, buf.AirToWaterRatio)
Call WritePTADInfoType(fnum, buf.AirFlowRate)
Call WritePTADInfoType(fnum, buf.AirPressureDrop)
Call WritePTADInfoType(fnum, buf.AirLoadingRate)
Call WritePTADInfoType(fnum, buf.MinimumAirToWaterRatio)
Call WritePTADInfoType(fnum, buf.MultipleOfMinimumAirToWaterRatio)
Call WritePTADInfoType(fnum, buf.KLaSafetyFactor)
Call WritePTADInfoType(fnum, buf.DesignMassTransferCoefficient)
Call WritePTADInfoType(fnum, buf.TowerArea)
Call WritePTADInfoType(fnum, buf.TowerDiameter)
Call WritePTADInfoType(fnum, buf.TowerHeight)
Call WritePTADInfoType(fnum, buf.TowerVolume)
Call WritePTADInfoType(fnum, buf.SpecifiedTowerDiameter)
Call WritePTADInfoType(fnum, buf.SpecifiedTowerHeight)

Call WritePackingDataType(fnum, buf.Packing)

Put #fnum, , buf.NumChemical
For i% = 1 To buf.NumChemical
    Call WritePTContamProperty(fnum, buf.Contaminant(i%))
Next i%

Put #fnum, , buf.ID_OptimalDesignContaminant
Call WritePTContamProperty(fnum, buf.DesignContaminant)

Call WriteOndaMassTransCoef(fnum, buf.Onda)

Put #fnum, , buf.Power.BlowerBrakePower
Put #fnum, , buf.Power.PumpBrakePower
Put #fnum, , buf.Power.TotalBrakePower
Put #fnum, , buf.Power.InletAirTemperature
Put #fnum, , buf.Power.BlowerEfficiency
Put #fnum, , buf.Power.PumpEfficiency

Put #fnum, , buf.TransferUnitHeight
Put #fnum, , buf.NumberOfTransferUnits
Put #fnum, , buf.Chemical

End Sub

Sub WriteSurfaceType(fnum As Integer, buf As SurfaceType)
    Dim i%


Call WriteSurfInfoType(fnum, buf.OperatingPressure)
Call WriteSurfInfoType(fnum, buf.operatingtemperature)
Call WriteSurfInfoType(fnum, buf.WaterDensity)
Call WriteSurfInfoType(fnum, buf.WaterViscosity)
Call WriteSurfInfoType(fnum, buf.PowerInput_PoverV)
Call WriteSurfInfoType(fnum, buf.N_for_Finding_KLa)
Call WriteSurfInfoType(fnum, buf.kgOVERkl_for_Finding_KLa)
Call WriteSurfInfoType(fnum, buf.ContaminantMassTransferCoefficient)
Call WriteSurfInfoType(fnum, buf.WaterFlowRate)
Call WriteSurfInfoType(fnum, buf.TankHydraulicRetentionTime)
Call WriteSurfInfoType(fnum, buf.TotalHydraulicRetentionTime)
Call WriteSurfInfoType(fnum, buf.TankVolume)
Call WriteSurfInfoType(fnum, buf.TotalTankVolume)

Put #fnum, , buf.NumberOfTanks.value
Put #fnum, , buf.NumberOfTanks.UserInput
Put #fnum, , buf.NumberOfTanks.ValChanged

Put #fnum, , buf.CodeForTausAndTankVolumes
Put #fnum, , buf.DesiredPercentRemoval
Put #fnum, , buf.AchievedPercentRemoval

Put #fnum, , buf.Power.AeratorMotorEfficiency
Put #fnum, , buf.Power.PowerForEachTank
Put #fnum, , buf.Power.TotalPowerForAllTanks

Put #fnum, , buf.Oxygen.KLaMethod
Call WriteSurfInfoType(fnum, buf.Oxygen.LiquidDiffusivity)
Call WriteSurfInfoType(fnum, buf.Oxygen.MassTransferCoefficient)

Put #fnum, , buf.NumChemical
Put #fnum, , buf.Chemical
For i% = 1 To buf.NumChemical
    Call WriteSurfContamProperty(fnum, buf.Contaminant(i%))
Next i%
Call WriteSurfContamProperty(fnum, buf.DesignContaminant)


End Sub

Sub WriteSurfContamProperty(fnum As Integer, buf As SurfaceContaminantPropertyType)
    Dim strsize%
    Dim i%

strsize% = Len(buf.Name)
Put #fnum, , strsize%
Put #fnum, , buf.Name
Put #fnum, , buf.Pressure
Put #fnum, , buf.Temperature
Call WriteSurfInfoType(fnum, buf.MolecularWeight)
Call WriteSurfInfoType(fnum, buf.HenrysConstant)
Call WriteSurfInfoType(fnum, buf.MolarVolume)
Call WriteSurfInfoType(fnum, buf.LiquidDiffusivity)
Call WriteSurfInfoType(fnum, buf.Influent)
Call WriteSurfInfoType(fnum, buf.TreatmentObjective)

For i% = 0 To MAXIMUM_TANKS
    Put #fnum, , buf.Effluent(i%)
Next i%
End Sub

Sub WriteSurfInfoType(fnum As Integer, buf As SurfaceInformationType)

Put #fnum, , buf.value
Put #fnum, , buf.UserInput
Put #fnum, , buf.ValChanged

End Sub

