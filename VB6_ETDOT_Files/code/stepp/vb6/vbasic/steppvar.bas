Attribute VB_Name = "SteppVarMod"
'This module contains the declaration of the variables and structures needed
'in the StEPP program.  It also contains type declarations for the user-defined types.
'It also contains CONSTANTS for the PROPERTIES AVAILABLE that are used with the
'PROPAVAILABLE and HAVEPROPERTY arrays.


Global Const NC = 2
Global Database_Path  As String

Global Const NDCONSTANT = 10
Global Const Maxchemical = 20    '/* Maximum no. of occurrences of any chemical
                          '                               in the database */
Global Const MAXNAME = 40        '/* Maximum length of a chemical name */
Global Const MAXFORMULA = 14     '/* Maximum length of a chemical formula */

Global Const ND = 10
Global Const NUMBER_OF_PROPERTIES_AVAILABLE = 75      'Corresponds to PROPAVAILABLE array and PROPERTIES AVAILABLE AND SOURCES set of constants
Global Const NUMBER_OF_PROPERTIES = 20                'Corresponds to HAVEPROPERTY array and PROPERTIES AVAILABLE set of constants
Global Const MAXSELECTEDCHEMICALS = 10 '* Maximum no. of chemicals user is allowed to select

'****************************************************
'       Set of Constants:  PROPERTIES AVAILABLE AND SOURCES
'                          Corresponds to PROPAVAILABLE array
'
Global Const OPERATING_PRESSURE = 1
Global Const OPERATING_TEMPERATURE = 2
Global Const VAPOR_PRESSURE_DATABASE = 3
Global Const VAPOR_PRESSURE_INPUT = 4
Global Const ACTIVITY_COEFFICIENT_UNIFAC = 5
Global Const ACTIVITY_COEFFICIENT_INPUT = 6
Global Const HENRYS_CONSTANT_REGRESS = 7
Global Const HENRYS_CONSTANT_FIT = 8
Global Const HENRYS_CONSTANT_OPT_UNIFAC = 9
Global Const HENRYS_CONSTANT_DATABASE = 10
Global Const HENRYS_CONSTANT_UNIFAC = 11
Global Const HENRYS_CONSTANT_INPUT = 12
Global Const MOLECULAR_WEIGHT_DATABASE = 13
Global Const MOLECULAR_WEIGHT_UNIFAC = 14
Global Const MOLECULAR_WEIGHT_INPUT = 15
Global Const BOILING_POINT_DATABASE = 16
Global Const BOILING_POINT_INPUT = 17
Global Const LIQUID_DENSITY_DATABASE = 18
Global Const LIQUID_DENSITY_UNIFAC = 19
Global Const LIQUID_DENSITY_INPUT = 20
Global Const MOLAR_VOLUME_NBP_UNIFAC = 21
Global Const MOLAR_VOLUME_NBP_INPUT = 22
Global Const MOLAR_VOLUME_OPT_DATABASE = 23
Global Const MOLAR_VOLUME_OPT_UNIFAC = 24
Global Const MOLAR_VOLUME_OPT_INPUT = 25
Global Const REFRACTIVE_INDEX_DATABASE = 26
Global Const REFRACTIVE_INDEX_INPUT = 27
Global Const AQUEOUS_SOLUBILITY_FIT = 28
Global Const AQUEOUS_SOLUBILITY_OPT_UNIFAC = 29
Global Const AQUEOUS_SOLUBILITY_DATABASE = 30
Global Const AQUEOUS_SOLUBILITY_DBT_UNIFAC = 31
Global Const AQUEOUS_SOLUBILITY_INPUT = 32
Global Const OCT_WATER_PART_COEFF_DB = 33
Global Const OCT_WATER_PART_COEFF_DBT_UNIFAC = 34
Global Const OCT_WATER_PART_COEFF_OPT_UNIFAC = 35
Global Const OCT_WATER_PART_COEFF_INPUT = 36
Global Const LIQUID_DIFFUSIVITY_POLSON = 37
Global Const LIQUID_DIFFUSIVITY_HAYDUKLAUDIE = 38
Global Const LIQUID_DIFFUSIVITY_WILKECHANG = 39
Global Const LIQUID_DIFFUSIVITY_INPUT = 40
Global Const GAS_DIFFUSIVITY_WILKELEE = 41
Global Const GAS_DIFFUSIVITY_INPUT = 42
Global Const WATER_DENSITY_CORRELATION = 43
Global Const WATER_DENSITY_INPUT = 44
Global Const WATER_VISCOSITY_CORRELATION = 45
Global Const WATER_VISCOSITY_INPUT = 46
Global Const WATER_SURF_TENSION_CORRELATION = 47
Global Const WATER_SURF_TENSION_INPUT = 48
Global Const AIR_DENSITY_CORRELATION = 49
Global Const AIR_DENSITY_INPUT = 50
Global Const AIR_VISCOSITY_CORRELATION = 51
Global Const AIR_VISCOSITY_INPUT = 52

'    Set of Constants:  PROPERTIES AVAILABLE
'                       Corresponds to HAVEPROPERTY array

Global Const VAPOR_PRESSURE = 3
Global Const ACTIVITY_COEFFICIENT = 4
Global Const HENRYS_CONSTANT = 5
Global Const MOLECULAR_WEIGHT = 6
Global Const BOILING_POINT = 7
Global Const LIQUID_DENSITY = 8
Global Const MOLAR_VOLUME_BOILING_POINT = 9
Global Const MOLAR_VOLUME_OPT = 10
Global Const REFRACTIVE_INDEX = 11
Global Const AQUEOUS_SOLUBILITY = 12
Global Const OCT_WATER_PART_COEFF = 13
Global Const LIQUID_DIFFUSIVITY = 14
Global Const GAS_DIFFUSIVITY = 15
Global Const WATER_DENSITY = 16
Global Const WATER_VISCOSITY = 17
Global Const WATER_SURFACE_TENSION = 18
Global Const AIR_DENSITY = 19
Global Const AIR_VISCOSITY = 20

'****************************************************

Global Const BIP_DB_ORIGINAL_UNIFAC_VLE = 1     'Corresponds to MDL = 1 in FORTRAN code
Global Const BIP_DB_UNIFAC_LLE = 2              'Corresponds to MDL = 2 in FORTRAN code
Global Const BIP_DB_ENVIRONMENTAL = 3           'Corresponds to MDL = 3 in FORTRAN code


Type CurrentSelectionType
   choice As Long     'Choice corresponding to properties available and sources
   source As Long     'Source corresponding to short list of sources
   Value As Double
End Type

Type sourceType
        short As Long
        long As Long
End Type

Type temperatureType
        temperature As Double
End Type

Type temperatureRangeType
        minimumT As Double
        maximumT As Double
End Type

Type VPsuperfundType
        Value As Double
        temperature As Double
End Type

Type informationType
        Value As Double
        source As sourceType
        error As Long
        equation As Long
        temperature As Double
End Type

Type databaseType
        database As informationType
End Type

Type unifacType
        UNIFAC As informationType
End Type

Type inputType
        input As informationType
End Type

Type databaseUnifacInputType
        database As informationType
        UNIFAC As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type unifacInputType
        UNIFAC As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type databaseInputType
        database As informationType
        input As informationType
End Type

Type VPintermediary
        Value As Double
        ncoeffs As Long
        source As sourceType
        error As Long
        equation As Long
        temperature As Double
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
        antoineA As Double
        antoineB As Double
        antoineC As Double
        antoineD As Double
        antoineE As Double
        minimumT As Double
        maximumT As Double
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type activityCoefficientType
        UNIFAC As informationType
        input As informationType
        BinaryInteractionParameterDBAvailable(1 To 3) As Long  'Array storing whether a particular UNIFAC parameter set is a valid choice for a compound.  Initialized to True and then set to False if this particular choice is unavailable.  Indexing corresponds to hierarchy.
        PreviousBinaryInteractionParameterDB As Long
        BinaryInteractionParameterDatabase As Long
        CurrentSelection As CurrentSelectionType
End Type

Type henrysConstantType
        RTI As informationType
        operatingT As unifacType
        regress As informationType
        fit As unifacType
        NumberOfDatabaseHenrysConstants As Long
        database(1 To Maxchemical) As informationType
        chosenDatabaseIndex As Long
        UNIFAC(1 To Maxchemical) As informationType
        chosenUNIFACIndex As Long
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type molecularWeightType
        database As informationType
        UNIFAC As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type boilingPointType
        database As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type liquidDensityType
        dbase_n_coeffs As Long
        dbase_coeffA As Double
        dbase_coeffB As Double
        dbase_coeffC As Double
        dbase_coeffD As Double
        dbase_minT As Double
        dbase_maxT As Double
        database As informationType
        UNIFAC As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type molarVolumeType
        operatingT As databaseUnifacInputType
        BoilingPoint As unifacInputType
End Type

Type refractiveIndexType
        database As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type aqueousSolubilityType
        fit As unifacType
        operatingT As unifacType
        database As informationType
        UNIFAC As informationType
        input As informationType
        BinaryInteractionParameterDBAvailable(1 To 3) As Long  'Array storing whether a particular UNIFAC parameter set is a valid choice for a compound.  Initialized to True and then set to False if this particular choice is unavailable.  Indexing corresponds to hierarchy.
        PreviousBinaryInteractionParameterDB As Long
        BinaryInteractionParameterDatabase As Long
        CurrentSelection As CurrentSelectionType
End Type

Type octWaterPartCoeffType
        database As informationType
        operatingT As unifacType
        databaseT As unifacType
        input As informationType
        BinaryInteractionParameterDBAvailable(1 To 3) As Long  'Array storing whether a particular UNIFAC parameter set is a valid choice for a compound.  Initialized to True and then set to False if this particular choice is unavailable.  Indexing corresponds to hierarchy.
        PreviousBinaryInteractionParameterDB As Long
        BinaryInteractionParameterDatabase As Long
        CurrentSelection As CurrentSelectionType
End Type

Type liquidDiffusivityType
        polson As informationType
        haydukLaudie As informationType
        wilkeChang As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type gasDiffusivityType
        wilkeLee As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type waterDensityType
        correlation As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type waterViscosityType
        correlation As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type waterSurfaceTensionType
        correlation As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type airDensityType
        correlation As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type

Type airViscosityType
        correlation As informationType
        input As informationType
        CurrentSelection As CurrentSelectionType
End Type
                                                                                                 
Type phpr    'PHPR --> PHysical PRoperties:  structure to hold physical properties
        CASNumber As Long
        Name As String * 42
        formula As String * 14
        OperatingPressure As Double
        OperatingTemperature As Double
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
        NumberofRingsinCompound As Long
        MaximumUnifacGroups As Long
        MS(1 To 10, 1 To 10, 1 To 2) As Long
        XMW(1 To ND) As Double
        HaveProperty(1 To NUMBER_OF_PROPERTIES) As Long
        PROPAVAILABLE(1 To NUMBER_OF_PROPERTIES_AVAILABLE)  As Long
End Type



Global FGRPErrorFlag As Long     'Error flag corresponding to FORTRAN routine called FGRPCALL

Type inp     'INP --> structure to read values from database into
        CASNumber As Long
        Name As String * 42
        formula As String * 14
        MolecularWeight As Double
        '/* field for whether molecular weights have been double checked goes here */
        HenrysConstant(1 To Maxchemical) As Double
        HenrysConstantTemperature(1 To Maxchemical) As Double
        HenrysConstantSource As Long
        VaporPressureSuperfund As Double
        VaporPressureSuperfundTemperature As Double
        LiquidDensityEquation As Long
        LiquidDensityNumberCoefficients As Long
        LiquidDensityCoefficientA As Double
        LiquidDensityCoefficientB As Double
        LiquidDensityCoefficientC As Double
        LiquidDensityCoefficientD As Double
        LiquidDensityMinimumT As Double
        LiquidDensityMaximumT As Double
        LiquidDensitySource As Long
        VaporPressureDatabaseEquation As Long
        VaporPressureNumberCoefficients As Long
        VaporPressureAntoineA As Double
        VaporPressureAntoineB As Double
        VaporPressureAntoineC As Double
        VaporPressureAntoineD As Double
        VaporPressureAntoineE As Double
        VaporPressureMinimumT As Double
        VaporPressureMaximumT As Double
        VaporPressureSource As Long
        NumberofRingsinCompound As Long
        MaximumUnifacGroups As Long
        MS(1 To 10, 1 To 10, 1 To 2) As Long
        AqueousSolubility As Double
        AqueousSolubilityTemperature As Double
        AqueousSolubilitySource As Long
        OctWaterPartCoeff As Double
        OctWaterPartCoeffTemperature As Double
        OctWaterPartCoeffSource As Long
        BoilingPoint As Double
        BoilingPointSource As Long
        RefractiveIndex As Double
        RefractiveIndexSource As Long
        OperatingTemperature As Double
        NumberOfDatabaseHenrysConstants As Long
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


'********  Set up hierarchy structure and selected structure for properties  ********

Type HierarchyType
     hierarchy As Long
     source As String
End Type


'****  This structure will contain the hierarchy for properties
'****  and also the currently selected properties for StEPP.
'****  Note that the values assigned to these variables correspond
'****  to the global constants given above
'****     e.g. ACTIVITY_COEFFICIENT_UNIFAC = 5
'****          MOLAR_VOLUME_OPT_UNIFAC = 24  etc.

Type HierarchyChoices
     VaporPressure(1 To 2) As HierarchyType
     ActivityCoefficient(1 To 2) As HierarchyType
     HenrysConstant(1 To 6) As HierarchyType
     MolecularWeight(1 To 3) As HierarchyType
     BoilingPoint(1 To 2) As HierarchyType
     LiquidDensity(1 To 3) As HierarchyType
     MolarVolumeBoilingPoint(1 To 2)  As HierarchyType
     MolarVolumeOperatingT(1 To 3) As HierarchyType
     RefractiveIndex(1 To 2) As HierarchyType
     AqueousSolubility(1 To 5) As HierarchyType
     OctWaterPartCoeff(1 To 4) As HierarchyType
     LiquidDiffusivityMWTlt1000(1 To 4) As HierarchyType
     LiquidDiffusivityMWTgt1000(1 To 4) As HierarchyType
     GasDiffusivity(1 To 2) As HierarchyType
     WaterDensity(1 To 2) As HierarchyType
     WaterViscosity(1 To 2) As HierarchyType
     WaterSurfaceTension(1 To 2) As HierarchyType
     AirDensity(1 To 2) As HierarchyType
     AirViscosity(1 To 2) As HierarchyType
End Type

'*** This structure will store the hierarchy for properties relating
'*** to choice of UNIFAC Binary interaction parameter database

Type BIP_DB_Hierarchy_Type
     ActivityCoefficient(1 To 3) As Long
     AqueousSolubility(1 To 3) As Long
     OctWaterPartCoeff(1 To 2) As Long
End Type


'*** This structure will be used to store the indexes for the
'*** previously selected values on the property forms for use
'*** when a user clicks on a new property

Type HighlightingSelectedValue
     PreviousIndex As Integer
End Type

Type HighlightProperties
     VaporPressure As HighlightingSelectedValue
     ActivityCoefficient As HighlightingSelectedValue
     HenrysConstant As HighlightingSelectedValue
     MolecularWeight As HighlightingSelectedValue
     BoilingPoint As HighlightingSelectedValue
     LiquidDensity As HighlightingSelectedValue
     MolarVolumeBoilingPoint As HighlightingSelectedValue
     MolarVolumeOperatingT As HighlightingSelectedValue
     RefractiveIndex As HighlightingSelectedValue
     AqueousSolubility As HighlightingSelectedValue
     OctWaterPartCoeff As HighlightingSelectedValue
     LiquidDiffusivity As HighlightingSelectedValue
     GasDiffusivity As HighlightingSelectedValue
     WaterDensity As HighlightingSelectedValue
     WaterViscosity As HighlightingSelectedValue
     WaterSurfaceTension As HighlightingSelectedValue
     AirDensity As HighlightingSelectedValue
     AirViscosity As HighlightingSelectedValue
End Type



Global dbinput As inp   'input structure for chemical to store properties as they are read from the database



Global phprop As phpr '/* physical properties structure */
Global hie As HierarchyChoices '* default hierarchy structure *

Global hilight As HighlightProperties '* current highlighted value on each form *

Global HaveProperty(1 To NUMBER_OF_PROPERTIES)   As Long
Global PROPAVAILABLE(1 To NUMBER_OF_PROPERTIES_AVAILABLE)  As Long

Global Find_String As String

Global PropContaminant(1 To MAXSELECTEDCHEMICALS) As phpr
Global NumSelectedChemicals As Integer

Global PreviouslySelectedIndex  'Index of item in cboSelectContaminant selected previously

Global steppPath As String      'Path where StEPP is stored
Global SaveAndLoadPath As String   'Path where user has specified for saving and loading

Global BIP_dbHierarchy As BIP_DB_Hierarchy_Type  'Hierarchy for UNIFAC parameter set
Global UserSelectedTheUnifacBIPDBAqSol As Integer      'Boolean variable telling if user has selected this or if it is to be done by the program according to hierarchy
Global UserSelectedTheUnifacBIPDBActCoeff As Integer   'Boolean variable telling if user has selected this or if it is to be done by the program according to hierarchy
Global UserSelectedTheUnifacBIPDBKow As Integer        'Boolean variable telling if user has selected this or if it is to be done by the program according to hierarchy

