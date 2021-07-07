Attribute VB_Name = "ScrnTypeMod"
Option Explicit
Type PTADInformationType
     value As Double
     ValChanged As Integer
     UserInput As Integer
End Type

Type ContaminantPropertyType
     Name As String
     Pressure As Double
     Temperature As Double
     AirWaterInterfaceConcentration As Double
     MolecularWeight As PTADInformationType
     HenrysConstant As PTADInformationType
     MolarVolume As PTADInformationType
     NormalBoilingPoint As PTADInformationType
     LiquidDiffusivity As PTADInformationType
     GasDiffusivity As PTADInformationType
     Influent As PTADInformationType
     TreatmentObjective As PTADInformationType
     Effluent As PTADInformationType
End Type

Type OndaMassTransferCoefficientType
     ReynoldsNumber As Double
     FroudeNumber As Double
     WeberNumber As Double
     LiquidPhaseMassTransferResistance As Double
     GasPhaseMassTransferResistance As Double
     TotalMassTransferResistance As Double
     LiquidPhaseMassTransferCoefficient As Double
     GasPhaseMassTransferCoefficient As Double
     OverallMassTransferCoefficient As Double
     ValChanged As Integer
End Type

Type PowerType
     BlowerBrakePower As Double
     PumpBrakePower As Double
     TotalBrakePower As Double
     InletAirTemperature As Double
     BlowerEfficiency As Double
     PumpEfficiency As Double
End Type

Type SCR
     Packing As PackingDataType
     
     NumChemical As Long
     Contaminant(0 To MAXCHEMICAL) As ContaminantPropertyType
     DesignContaminant As ContaminantPropertyType
     
     Onda As OndaMassTransferCoefficientType
     
     ID_OptimalDesignContaminant As Long
     
     Power As PowerType

     TransferUnitHeight As Double
     NumberOfTransferUnits As Double
     Chemical As Integer
     
     OperatingPressure As PTADInformationType       'kPa
     operatingtemperature As PTADInformationType    'K
     WaterFlowRate As PTADInformationType           'm^3/s
     WaterDensity As PTADInformationType            '
     WaterViscosity As PTADInformationType          '
     WaterSurfaceTension As PTADInformationType     '
     WaterLoadingRate As PTADInformationType        'kg/m^2-s
     AirDensity As PTADInformationType              '
     AirViscosity As PTADInformationType            '
     AirToWaterRatio As PTADInformationType         '(-)
     AirFlowRate As PTADInformationType             'm^3/s
     AirPressureDrop As PTADInformationType         'Pa/m
     AirLoadingRate As PTADInformationType          'kg/m^2-s
     MinimumAirToWaterRatio As PTADInformationType  '
     MultipleOfMinimumAirToWaterRatio As PTADInformationType
     KLaSafetyFactor As PTADInformationType         '
     DesignMassTransferCoefficient As PTADInformationType
     TowerArea As PTADInformationType               'm^2
     TowerDiameter As PTADInformationType           'm
     TowerHeight As PTADInformationType             'm
     TowerVolume As PTADInformationType             'm^3
     SpecifiedTowerDiameter As PTADInformationType  'm
     SpecifiedTowerHeight As PTADInformationType    'm
End Type

Global SaveAndLoadPath As String

Global ShownScreen1Previously As Integer

Function GetTheFormat(value As Double) As String
   Dim AbsValue As Double

   AbsValue = Abs(value)

   If AbsValue < 0.001 Then
      GetTheFormat = "0.00E+00"
   ElseIf AbsValue < 0.01 Then
      GetTheFormat = "0.00E+00"
   ElseIf AbsValue < 0.1 Then
      GetTheFormat = "0.0000"
   ElseIf AbsValue < 1 Then
      GetTheFormat = "0.000"
   ElseIf AbsValue < 10 Then
      GetTheFormat = "0.00"
   ElseIf AbsValue < 100 Then
      GetTheFormat = "0.0"
   ElseIf AbsValue < 1000 Then
      GetTheFormat = "0"
   Else
      GetTheFormat = "0.00E+00"
   End If

End Function

