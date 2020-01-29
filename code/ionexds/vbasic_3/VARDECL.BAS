Option Explicit
Option Base 1

'Global Constants Used in Type Definitions

    Global Const MAX_NERNST_HASKELL_DB_IONS = 100
    Global Const MAX_CHEMICAL = 10
    Global Const MAX_AXIAL_COLLOCATION_POINTS = 18
    Global Const MAX_RADIAL_COLLOCATION_POINTS = 14
    Global Const Number_Data_Points_Max = 100

    Global Const Number_Points_Max = 400    'Maximum number of output points from PFPDM dll
    Global Const NVersion = .5
    Global Const Number_Compo_Max_PFPDM = 6

    'Correlations available to calculate Ionic Transport Coefficient, kf
    Global Const IONIC_TRANSPORT_COEFFICIENT_1 = "Wildhagen"
    Global Const IONIC_TRANSPORT_COEFFICIENT_2 = "Gnielinski"
    Global Const Number_Max_Influent_Points = 400
    Global Const Maximum_Beds_In_Series = 200
    Global Const EPS_ERROR_CRITERIA = .0005
'    Global Const EPS_ERROR_CRITERIA = .000001

'Type Definitions

Type User_Input_Or_Calculation_Type
   Value As Double
   UserInput As Integer
End Type

Type OperatingConditionsType
     Pressure As Double                         'N/m2
     Temperature As Double                      'K
     LiquidDensity As Double                    'g/cm3
     LiquidViscosity As Double                  'g/cm/s
End Type

Type BedPropertyType
    Length As Double                            'm
    Diameter As Double                          'm
    Weight As Double                            'kg
    Flowrate As User_Input_Or_Calculation_Type  'm3/s
    EBCT As User_Input_Or_Calculation_Type      'min
    Area As Double                              'm2
    Volume As Double                            'm3
    Density As Double                           'g/cm3
    Porosity As Double                          '(-)
    SuperficialVelocity As Double               'cm/s
    InterstitialVelocity As Double              'cm/s
    EffectiveContactTime As Double              's
    NumberOfBeds As Integer                        'dimensionless
End Type

Type AdsorbentPropertyType
    Name As String
    ParticlePorosity As Double     'Dimensionless
    ApparentDensity As Double      'g/cm3  Apparent density
    ParticleRadius As Double       'm
    ParticleDiameter As Double     'm
    Tortuosity As Double           'Dimensionless
    TotalCapacity As Double        'meq/g dry resin
    PresaturantPercentage(1 To MAX_CHEMICAL) As Double
End Type

Type NernstHaskellDatabaseType
    Ion_Name As String
    Valence As Double
    LimitingIonicConductance As Double   '(A/cm2)(V/cm)(g-equiv/cm3)
End Type

Type NernstHaskellIonType
    'Variables for Nernst-Haskell Database
    Anion(1 To MAX_NERNST_HASKELL_DB_IONS) As NernstHaskellDatabaseType
    Cation(1 To MAX_NERNST_HASKELL_DB_IONS) As NernstHaskellDatabaseType
    NumberOfCationsInDB As Integer
    NumberOfAnionsInDB As Integer
    'Variables for Use in Nernst-Haskell Calculation
    SelectedAnion As NernstHaskellDatabaseType
    SelectedCation As NernstHaskellDatabaseType
    DefaultAnion As NernstHaskellDatabaseType
    DefaultCation As NernstHaskellDatabaseType
    FaradaysConstant As Double       'cal/g/equiv
    GasConstant As Double            'J/mol/K
    LiquidDiffusivity As Double
End Type

    Global IonicTransportCoeffCorrName As String     'Name of Correlation used to calculate ionic transport coefficient

Type KineticParameterType
   NernstHaskellCation As NernstHaskellDatabaseType
   NernstHaskellAnion As NernstHaskellDatabaseType
   LiquidDiffusivity As User_Input_Or_Calculation_Type          'cm2/sec
   LiquidDiffusivityCorrelation As Double                       'cm2/sec
   LiquidDiffusivityUserInput As Double                         'cm2/sec
   SchmidtNumber As Double          '(-)
   ReynoldsNumber As Double         '(-)
   IonicTransportCoefficient As User_Input_Or_Calculation_Type  'cm/sec
   IonicTransportCoeffCorrelation As Double                     'cm/sec
   IonicTransportCoeffUserInput As Double                       'cm/sec
   PoreDiffusivity As User_Input_Or_Calculation_Type            'cm2/sec
   PoreDiffusivityCorrelation As Double                         'cm2/sec
   PoreDiffusivityUserInput As Double                           'cm2/sec
End Type

Type DimensionlessGroupsType
   SurfaceDistributionParameter As Double                       '(-)
   PoreDistributionParameter As Double                          '(-)
   TotalDistributionParameter As Double                         '(-)
   PoreDiffusionModulus As Double                               '(-)
   StantonNumber As Double                                      '(-)
   PoreBiotNumber As Double                                     '(-)
End Type


Type ComponentPropertyType
   Name As String
   Valence As Double
   SeparationFactor As Double
   MolecularWeight As Double                                      'mg/mmol
   InitialConcentration As Double                                 'mg/L
   EquivalentInitialConcentration As Double                       'meq/L
   Kinetic As KineticParameterType
   Dimensionless As DimensionlessGroupsType
End Type


Type separationFactorType
   Row As Integer    'True =>  Separation factors are being input across a row
                     'False => Separation Factors are being input down a column
   Value As Integer  'The number of the ion for which the Separation Factors are being input (i.e. the row or column)
End Type

Type AvailableIonsType
     Available As Integer
End Type

'Variable Declarations
    Global Bed As BedPropertyType
    Global Resin As AdsorbentPropertyType
    Global Operating As OperatingConditionsType
    Global NernstHaskell As NernstHaskellIonType
    Global Anion(0 To MAX_CHEMICAL) As ComponentPropertyType
    Global Cation(0 To MAX_CHEMICAL) As ComponentPropertyType
    Global Ion(0 To MAX_CHEMICAL) As ComponentPropertyType
    Global Ion_Array(0 To MAX_CHEMICAL) As ComponentPropertyType
    Global DefaultAnion As ComponentPropertyType
    Global DefaultCation As ComponentPropertyType
    Global ChangedIon As ComponentPropertyType
    Global NumberOfAnions As Integer
    Global NumberOfCations As Integer
    Global NumberOfIons As Integer
    Global PresaturantCation As Integer  'The Number of the Presaturant Cation in the List of Cations
    Global PresaturantAnion As Integer   'The Number of the Presaturant Anion in the List of Anions
    Global SeparationFactorInput As separationFactorType  'Cation for which separation factors are being input
    Global CationSeparationFactorInput As separationFactorType
    Global AnionSeparationFactorInput As separationFactorType
    Global AddingCation As Integer
    Global AddingAnion As Integer
    Global EditingCation As Integer
    Global EditingAnion As Integer
    Global NumberOfIonToEdit As Integer

    Global OneDimSeparationFactors(0 To MAX_CHEMICAL) As Double
    Global TwoDimSeparationFactors(0 To MAX_CHEMICAL, 0 To MAX_CHEMICAL) As Double
    Global OldOneDimSeparationFactors(0 To MAX_CHEMICAL) As Double
    Global OldOptionButtonSeparationFactors As Double

    Global OldCationKineticParameters(1 To MAX_CHEMICAL) As KineticParameterType   'Variable for Storing Old Cation Kinetic Parameters Before Modifying Kinetic Parameters on frmInputKineticParameters
    Global OldAnionKineticParameters(1 To MAX_CHEMICAL) As KineticParameterType   'Variable for Storing Old Anion Kinetic Parameters Before Modifying Kinetic Parameters on frmInputKineticParameters
    Global ClickGeneratedFromcboIon As Integer   'Whether the click event on opt button in Kinetic form was generated from cboIon

    Global SumAnionInitialEquivalents As Double   'meq/L  --> Sum of the time-averaged initial influent concentrations for anions
    Global SumCationInitialEquivalents As Double  'meq/L  --> Sum of the time-averaged initial influent concentrations for cations
    
    Global ViewingKineticParametersForm As Integer  'Parameter telling whether frmInputKineticParameters is showing

    Global Cations As AvailableIonsType   'Variable telling whether cations are available for the current resin
    Global Anions As AvailableIonsType    'Variable telling whether anions are available for the current resin

    'Variable telling whether it is possible to calculate dimensionless groups (i.e. if SumInitialEquivalentConcs > 0)
    Global OKToGetCationDimensionless As Integer
    Global OKToGetAnionDimensionless As Integer

    Global NumSelectedAnions As Integer
    Global NumSelectedCations As Integer
    Global NumSelectedComponents_PFPDM As Integer   'Number Of Components currently selected by the user
    Global Number_Influent_Points As Integer        'Number of variable influent data points
    Global Number_Component As Integer              'Number of components currently selected
    Global Total_NumberOfComponents As Integer

    Global Anions_Selected(1 To MAX_CHEMICAL) As Integer   'Anions selected for a simulation
    Global Cations_Selected(1 To MAX_CHEMICAL) As Integer  'Cations selected for a simulation
    Global Component_Index_PFPDM(1 To MAX_CHEMICAL) As Integer  'Ions selected by the user, either anions or cations

    Global AlphaInput(1 To MAX_CHEMICAL) As Double    'Array of separation factors sent to PFPDM DLL

Type TimeParametersType
   InitialTime As Double   'min
   FinalTime As Double     'min
   TimeStep As Double      'min
End Type

    Global NumAxialCollocationPoints As Integer
    Global NumRadialCollocationPoints As Integer

'------Begin Modification Hokanson: 11-Aug2000
    Global EPS_ErrorCriteriaForDGEARIntegrator As Double
    Global DH0_InitialTimeStepForDGEARIntegrator As Double   'min
'------End Modification Hokanson: 11-Aug2000

    Global TimeParameters As TimeParametersType

    Global Application_Name As String

Type ThroughPut
    T As Double
    c As Double
End Type

Type ResultsType
    NComponent As Integer
    NPoints As Integer
    CP(MAX_CHEMICAL, Number_Points_Max) As Double
    T(Number_Points_Max) As Double
    ThroughPut_05(MAX_CHEMICAL) As ThroughPut
    ThroughPut_50(MAX_CHEMICAL) As ThroughPut
    ThroughPut_95(MAX_CHEMICAL) As ThroughPut
    Component(MAX_CHEMICAL) As ComponentPropertyType
    Bed As BedPropertyType
    Resin As AdsorbentPropertyType
    Use_Tortuosity_Correlation As Integer
    Constant_Tortuosity As Integer
End Type

    Global Results As ResultsType
    Global Use_Tortuosity_Correlation As Integer
    Global Constant_Tortuosity As Integer

'Variable used to transfer data to excel
Global Excel_4 As Integer

'Arrays to store the experimental data points for comparison to a PFPSDM simulation
Global T_Data_Points(Number_Data_Points_Max)   As Double
Global C_Data_Points(Number_Compo_Max_PFPDM, Number_Data_Points_Max) As Double
Global NData_Points As Integer
Global NComponents As Integer

Global TimeUnitsOnGraphs As Integer

'Variables for Time Variable Influent Concentration

Global C_Influent(MAX_CHEMICAL, Number_Max_Influent_Points) As Double
Global T_Influent(Number_Max_Influent_Points) As Double

Global VarInfluentFileCation As String
Global VarInfluentFileAnion As String

Global Const VAR_INFLUENT_CATION_FILEID = "Variable Influent File - Cations"
Global Const VAR_INFLUENT_ANION_FILEID = "Variable Influent File - Anions"

