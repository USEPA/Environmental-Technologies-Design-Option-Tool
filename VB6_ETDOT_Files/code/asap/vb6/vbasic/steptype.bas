Attribute VB_Name = "StepType"
Option Explicit


Global Const MAXCHEMICAL = 10    '/* Maximum no. of occurrences of any chemical
                          '                               in the database */
Global Const MAXNAME = 40        '/* Maximum length of a chemical name */
Global Const MAXFORMULA = 14     '/* Maximum length of a chemical formula */

Global designtype As Integer  'ie (surface, bubble)=0,(ptad1)=1 ,(ptad2)=2 for import reasons

Type sourceType
        short As Integer
        long As Integer
End Type

Type temperatureType
        Temperature As Double
End Type

Type temperatureRangeType
        minimumT As Double
        maximumT As Double
End Type

Type VPsuperfundType
        value As Double
        Temperature As Double
End Type

Type informationType
        value As Double
        source As sourceType
        error As Integer
        equation As Integer
        Temperature As Double
End Type

Type databaseType
        database As informationType
End Type

Type unifacType
        unifac As informationType
End Type

Type inputType
        input As informationType
End Type

Type databaseUnifacInputType
        database As informationType
        unifac As informationType
        input As informationType
End Type

Type unifacInputType
        unifac As informationType
        input As informationType
End Type

Type databaseInputType
        database As informationType
        input As informationType
End Type

Type VPintermediary
        value As Double
        source As sourceType
        error As Integer
        equation As Integer
        Temperature As Double
        minimumT As Double
        maximumT As Double
        antoineA As Double
        antoineB As Double
        antoineC As Double
        antoineD As Double
        antoineE As Double
        superfund As VPsuperfundType
End Type

Type vaporPressureType
        database As VPintermediary
        input As informationType
End Type

Type activityCoefficientType
        unifac As informationType
        input As informationType
End Type

Type henrysConstantType
        RTI As informationType
        operatingT As unifacType
        regress As informationType
        fit As unifacType
        database(MAXCHEMICAL) As informationType
        unifac(MAXCHEMICAL) As informationType
        input As informationType
End Type

Type molecularWeightType
        database As informationType
        unifac As informationType
        input As informationType
End Type

Type boilingPointType
        database As informationType
        input As informationType
End Type

Type liquidDensityType
        database As informationType
        unifac As informationType
        input As informationType
End Type

Type molarVolumeType
        operatingT As databaseUnifacInputType
        BoilingPoint As unifacInputType
End Type

Type refractiveIndexType
        database As informationType
        input As informationType
End Type

Type aqueousSolubilityType
        fit As unifacType
        operatingT As unifacType
        database As informationType
        unifac As informationType
        input As informationType
End Type

Type octWaterPartCoeffType
        database As informationType
        unifac As informationType
        input As informationType
End Type

Type liquidDiffusivityType
        polson As informationType
        haydukLaudie As informationType
        wilkeChang As informationType
        input As informationType
End Type

Type gasDiffusivityType
        wilkeLee As informationType
        input As informationType
End Type

Type waterDensityType
        correlation As informationType
        input As informationType
End Type

Type waterViscosityType
        correlation As informationType
        input As informationType
End Type

Type waterSurfaceTensionType
        correlation As informationType
        input As informationType
End Type

Type airDensityType
        correlation As informationType
        input As informationType
End Type

Type airViscosityType
        correlation As informationType
        input As informationType
End Type
                                                                                                 
Type PHPR    'PHPR --> PHysical PRoperties:  structure to hold physical properties
        OperatingPressure As Double
        operatingtemperature As Double
        BinaryInteractionParameterDatabaseChoice As Integer
        VaporPressure As vaporPressureType
        ActivityCoefficient As activityCoefficientType
        HenrysConstant As henrysConstantType
        MolecularWeight As molecularWeightType
        BoilingPoint As boilingPointType
        LiquidDensity As liquidDensityType
        MolarVolume As molarVolumeType
        RefractiveIndex As refractiveIndexType
        AqueousSolubility As aqueousSolubilityType
        OctWaterPartCoeff As octWaterPartCoeffType
        LiquidDiffusivity As liquidDiffusivityType
        GasDiffusivity As gasDiffusivityType
        WaterDensity As waterDensityType
        WaterViscosity As waterViscosityType
        WaterSurfaceTension As waterSurfaceTensionType
        AirDensity As airDensityType
        AirViscosity As airViscosityType
End Type

Type INP     'INP --> structure to read values from database into
        CASnumber As Integer
        '/* place for contaminant name:  left out for now */
        '/* place for chemical formula:  left out for now */
        MolecularWeight As Double
        '/* field for whether molecular weights have been double checked goes here */
        HenrysConstant(MAXCHEMICAL) As Double
        HenrysConstantTemperature(MAXCHEMICAL) As Double
        HenrysConstantSource As Integer
        VaporPressureSuperfund As Double
        VaporPressureSuperfundTemperature As Double
        LiquidDensityEquation As Integer
        LiquidDensityNumberCoefficients As Integer
        LiquidDensityCoefficientA As Double
        LiquidDensityCoefficientB As Double
        LiquidDensityCoefficientC As Double
        LiquidDensityCoefficientD As Double
        LiquidDensityMinimumT As Double
        LiquidDensityMaximumT As Double
        LiquidDensitySource As Integer
        VaporPressureDatabaseEquation As Integer
        VaporPressureNumberCoefficients As Integer
        VaporPressureAntoineA As Double
        VaporPressureAntoineB As Double
        VaporPressureAntoineC As Double
        VaporPressureAntoineD As Double
        VaporPressureAntoineE As Double
        VaporPressureMinimumT As Double
        VaporPressureMaximumT As Double
        VaporPressureSource As Integer
        NumberofRingsinCompound As Integer
        MaximumUnifacGroups As Integer
        MS(10, 10, 2) As Double
        AqueousSolubility As Double
        AqueousSolubilityTemperature As Double
        AqueousSolubilitySource As Integer
        OctWaterPartCoeff As Double
        OctWaterPartCoeffTemperature As Double
        OctWaterPartCoeffSource As Integer
        BoilingPoint As Double
        BoilingPointSource As Integer
        RefractiveIndex As Double
        RefractiveIndexSource As Integer
        operatingtemperature As Double
        NumberofDatabaseHenrysConstants As Integer

End Type
      

Type StrippingContaminantProperties
     Name As String
     MolecularWeight As Double
     HenrysConstant As Double
     MolarVolume As Double
     NormalBoilingPoint As Double
     LiquidDiffusivity As Double
     GasDiffusivity As Double
End Type

