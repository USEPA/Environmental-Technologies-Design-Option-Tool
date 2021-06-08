Attribute VB_Name = "Structs"
Option Explicit

Global Const Latest_DataVersion_Major = 1
Global Const Latest_DataVersion_Minor = 0

Global Const PRESSURE_PA = 0
Global Const PRESSURE_KPA = 1
Global Const PRESSURE_BARS = 2
Global Const PRESSURE_ATM = 3
Global Const PRESSURE_PSI = 4
Global Const PRESSURE_MMHG = 5
Global Const PRESSURE_MH2O = 6
Global Const PRESSURE_FTH2O = 7
Global Const PRESSURE_INHG = 8

Global Const TEMPERATURE_K = 0
Global Const TEMPERATURE_C = 1
Global Const TEMPERATURE_R = 2
Global Const TEMPERATURE_F = 3

Global Const LENGTH_M = 0
Global Const LENGTH_CM = 1
Global Const LENGTH_FT = 2
Global Const LENGTH_IN = 3

Global Const MASS_KG = 0
Global Const MASS_G = 1
Global Const MASS_LB = 2

Global Const FLOW_M3_per_S = 0
Global Const FLOW_M3_per_D = 1
Global Const FLOW_CM3_per_S = 2
Global Const FLOW_ML_per_MIN = 3
Global Const FLOW_FT3_per_S = 4
Global Const FLOW__FT3_per_D = 5
Global Const FLOW_GPM = 6
Global Const FLOW_GPD = 7
Global Const FLOW_MGD = 8
    
Global Const TIME_MIN = 1
Global Const TIME_S = 0
Global Const TIME_HR = 2
Global Const TIME_D = 3

Global Const APPARENT_DENSITY_G_per_ML = 0
Global Const APPARENT_DENSITY_KG_per_M3 = 1
Global Const APPARENT_DENSITY_LB_per_FT3 = 2
Global Const APPARENT_DENSITY_LB_per_GAL = 3

Global Const RESIN_CAPACITY_MEQ_per_G = 0
Global Const RESIN_CAPACITY_MEQ_per_MLbed = 1
Global Const RESIN_CAPACITY_MEQ_per_MLresin = 2

Global Const MOLECULAR_WEIGHT_MG_per_MMOL = 0
Global Const MOLECULAR_WEIGHT_UG_per_UMOL = 1
Global Const MOLECULAR_WEIGHT_G_per_GMOL = 2
Global Const MOLECULAR_WEIGHT_KG_per_KMOL = 3

Global Const CONCENTRATION_MG_per_L = 0
Global Const CONCENTRATION_UG_per_L = 1
Global Const CONCENTRATION_G_per_L = 2
Global Const CONCENTRATION_MEQ_per_L = 3
Global Const CONCENTRATION_EQ_per_L = 4
Global Const CONCENTRATION_MMOL_per_L = 5
Global Const CONCENTRATION_UMOL_per_L = 6
Global Const CONCENTRATION_GMOL_per_L = 7

Global Const DIFFUSIVITY_CM2_per_S = 0
Global Const DIFFUSIVITY_CM2_per_MIN = 1
Global Const DIFFUSIVITY_M2_per_S = 2
Global Const DIFFUSIVITY_M2_per_MIN = 3
Global Const DIFFUSIVITY_M2_per_HR = 4
Global Const DIFFUSIVITY_M2_per_D = 5
Global Const DIFFUSIVITY_FT2_per_S = 6
Global Const DIFFUSIVITY_FT2_per_MIN = 7
Global Const DIFFUSIVITY_FT2_per_HR = 8
Global Const DIFFUSIVITY_FT2_per_D = 9

Global Const VELOCITY_CM_per_S = 0
Global Const VELOCITY_CM_per_MIN = 1
Global Const VELOCITY_M_per_S = 2
Global Const VELOCITY_M_per_MIN = 3
Global Const VELOCITY_M_per_HR = 4
Global Const VELOCITY_M_per_D = 5
Global Const VELOCITY_FT_per_S = 6
Global Const VELOCITY_FT_per_MIN = 7
Global Const VELOCITY_FT_per_HR = 8
Global Const VELOCITY_FT_per_D = 9

'''''Option Explicit
'''''Option Base 1

'Global Constants Used in Type Definitions
Global Const MAX_NERNST_HASKELL_DB_IONS = 100
Global Const MAX_CHEMICAL = 10
Global Const MAX_AXIAL_COLLOCATION_POINTS = 18
Global Const MAX_RADIAL_COLLOCATION_POINTS = 14
Global Const Number_Data_Points_Max = 100
Global Const Number_Points_Max = 400    'Maximum number of output points from PFPDM dll
Global Const NVersion = 0.5
Global Const Number_Compo_Max_PFPDM = 6

'Correlations available to calculate Ionic Transport Coefficient, kf
Global Const IONIC_TRANSPORT_COEFFICIENT_1 = "Wildhagen"
Global Const IONIC_TRANSPORT_COEFFICIENT_2 = "Gnielinski"
Global Const Number_Max_Influent_Points = 400
Global Const Maximum_Beds_In_Series = 200
Global Const EPS_ERROR_CRITERIA = 0.0005
'Global Const EPS_ERROR_CRITERIA = .000001

Global Const VAR_INFLUENT_CATION_FILEID = "Variable Influent File - Cations"
Global Const VAR_INFLUENT_ANION_FILEID = "Variable Influent File - Anions"

'Type Definitions
Type User_Input_Or_Calculation_Type
   Value As Double
   UserInput As Double
End Type

Type OperatingConditionsType
     Pressure As Double                         'N/m2
     Temperature As Double                      'K
     LiquidDensity As Double                    'g/cm3
     LiquidViscosity As Double                  'g/cm/s
End Type

Type BedPropertyType
    length As Double                            'm
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
    NumberOfBeds As Double                 'dimensionless
End Type

Type AdsorbentPropertyType
    Name As String
    ParticlePorosity As Double     'Dimensionless
    ApparentDensity As Double      'g/cm3  Apparent density
    ParticleRadius As Double       'm
    ParticleDiameter As Double     'm
    Tortuosity As Double           'Dimensionless
    TotalCapacity As Double        'meq/g dry resin
    PresaturantPercentage(1 To MAX_CHEMICAL) As Double   '1 To MAX_CHEMICAL
End Type

Type NernstHaskellDatabaseType
    Ion_Name As String
    Valence As Double
    LimitingIonicConductance As Double   '(A/cm2)(V/cm)(g-equiv/cm3)
End Type

Type NernstHaskellIonType
    'Variables for Nernst-Haskell Database
    Anion() As NernstHaskellDatabaseType    '1 To MAX_NERNST_HASKELL_DB_IONS
    Cation() As NernstHaskellDatabaseType   '1 To MAX_NERNST_HASKELL_DB_IONS
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
   Row As Double    'True =>  Separation factors are being input across a row
                     'False => Separation Factors are being input down a column
   Value As Double  'The number of the ion for which the Separation Factors are being input (i.e. the row or column)
End Type

Type AvailableIonsType
     Available As Double
End Type
   
Type TimeParametersType
   InitialTime As Double   'min
   FinalTime As Double     'min
   TimeStep As Double      'min
End Type

Type ThroughPut
    T As Double
    c As Double
End Type

Type ResultsType
    NComponent As Integer
    NPoints As Integer
    CP(MAX_CHEMICAL, Number_Points_Max) As Double  '
    T(Number_Points_Max) As Double   '
    ThroughPut_05(MAX_CHEMICAL) As ThroughPut   '
    ThroughPut_50(MAX_CHEMICAL) As ThroughPut   '
    ThroughPut_95(MAX_CHEMICAL) As ThroughPut   '
    Component(MAX_CHEMICAL) As ComponentPropertyType    '
    Bed As BedPropertyType
    Resin As AdsorbentPropertyType
    Use_Tortuosity_Correlation As Integer
    Constant_Tortuosity As Integer
End Type

'Declare DLL's for water density and water viscosity
'DLLs for Air & Water Properties
'H2oDens Returns D as kg/m3 and Input T as K
Declare Sub H2ODens Lib "H2oDens.DLL" (D As Double, T As Double)

'H2oVisc Returns V as kg/m/sec and Input T as K
Declare Sub H2OVisc Lib "H2oVisc.DLL" (v As Double, T As Double)

'DLL for PFPDM - Version without variable influent or beds in series
'Declare Sub PFPDM04 Lib "pdm04dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, CIN As Double, Size As Long, FLAG As Integer)

'DLL for PFPDM - Version with variable influent or beds in series
'Declare Sub PFPDM05 Lib "pdm05dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, Size As Long, FLAG As Integer)

'DLL for PFPDM - Version with exiting dll when reached 1 % of influent concentration
'Declare Sub PFPDM06 Lib "pdm06dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, Size As Long, FLAG As Integer)

'DLL for PFPDM - Version for setting CPORE to 0 if it is <= 0 and Calculating QH as QT - (Sum of Other Ions)
'Declare Sub PFPDM07 Lib "pdm07dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, CT_Average As Double, Size As Long, FLAG As Integer)

'------Begin Modification Hokanson: 11-Aug2000
'DLL for PFPDM - Version for setting CPORE to 0 if it is <= 0 and Calculating QH as QT - (Sum of Other Ions)
'Declare Sub PFPDM08 Lib "pdm08dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, CT_Average As Double, Size As Long, FLAG As Integer)

'------Begin Modification Hokanson: 12-Aug2000
'Declare Sub PFPDM09 Lib "pdm09dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, CT_Average As Double, Size As Long, FLAG As Integer, EPS As Double, DH0 As Double)
'------End Modification Hokanson: 11-Aug2000

Declare Sub PFPDM10 Lib "pdm10dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, M As Integer, Nin As Integer, Tin As Double, Cin As Double, CT_Average As Double, Size As Long, FLAG As Integer, EPS As Double, DH0 As Double)
'------End Modification Hokanson: 12-Aug2000

'Fit Data subroutines
Declare Sub OBJFUN Lib "OBJFUN.DLL" (NCOMP As Long, NDATA As Long, NP As Long, TP As Double, CP As Double, TD As Double, CD As Double, Cin As Double, FMIN As Double)

Type Project_Type
  filename As String
  dirty As Integer
  
  FileID As String
  Operating As OperatingConditionsType
  Bed As BedPropertyType
  Resin As AdsorbentPropertyType
  TimeParameters As TimeParametersType
  EPS_ErrorCriteriaForDGEARIntegrator As String
  DH0_InitialTimeStepForDGEARIntegrator As Double
  NumAxialCollocationPoints As Double
  NumRadialCollocationPoints As Double
  IonicTransportCoeffCorrName As String     'Name of Correlation used to calculate ionic transport coefficient
  
  NumberOfCations As Double
  PresaturantCation As Double  'The Number of the Presaturant Cation in the List of Cations
  SumCationInitialEquivalents As Double  'meq/L  --> Sum of the time-averaged initial influent concentrations for cations
  OKToGetCationDimensionless As Double
  CationSeparationFactorInput As separationFactorType
  Cation(0 To MAX_CHEMICAL) As ComponentPropertyType
  
  NumberOfAnions As Double
  PresaturantAnion As Double   'The Number of the Presaturant Anion in the List of Anions
  SumAnionInitialEquivalents As Double   'meq/L  --> Sum of the time-averaged initial influent concentrations for anions
  OKToGetAnionDimensionless As Double
  AnionSeparationFactorInput As separationFactorType
  Anion(0 To MAX_CHEMICAL) As ComponentPropertyType
  
  VarInfluentFileCation As String
  VarInfluentFileAnion As String

  Application_Name As String
  
End Type

Global Results As ResultsType

'Variable used to transfer data to excel
Global Excel_4 As Integer

'Arrays to store the experimental data points for comparison to a PFPSDM simulation
Global T_Data_Points(Number_Data_Points_Max)   As Double    '
Global C_Data_Points(Number_Compo_Max_PFPDM, Number_Data_Points_Max) As Double      '
Global NData_Points As Integer
Global NComponents As Integer
Global TimeUnitsOnGraphs As Integer
Global C_Influent(MAX_CHEMICAL, Number_Max_Influent_Points) As Double  '
Global T_Influent(Number_Max_Influent_Points) As Double         '
Global NernstHaskell As NernstHaskellIonType
Global Ion(0 To MAX_CHEMICAL) As ComponentPropertyType   '
Global Ion_Array(0 To MAX_CHEMICAL) As ComponentPropertyType   '
Global DefaultAnion As ComponentPropertyType
Global DefaultCation As ComponentPropertyType
Global ChangedIon As ComponentPropertyType
Global SeparationFactorInput As separationFactorType  'Cation for which separation factors are being input
Global OldCationKineticParameters(1 To MAX_CHEMICAL) As KineticParameterType   ', Variable for Storing Old Cation Kinetic Parameters Before Modifying Kinetic Parameters on frmInputKineticParameters
Global OldAnionKineticParameters(1 To MAX_CHEMICAL) As KineticParameterType   ', Variable for Storing Old Anion Kinetic Parameters Before Modifying Kinetic Parameters on frmInputKineticParameters
Global NumberOfIons As Integer
Global Current_Filename As String
Global AddingCation As Integer
Global AddingAnion As Integer
Global EditingCation As Integer
Global EditingAnion As Integer
Global NumberOfIonToEdit As Integer
Global OneDimSeparationFactors(0 To MAX_CHEMICAL) As Double    '
Global TwoDimSeparationFactors(0 To MAX_CHEMICAL, 0 To MAX_CHEMICAL) As Double    '
Global OldOneDimSeparationFactors(0 To MAX_CHEMICAL) As Double   '
Global OldOptionButtonSeparationFactors As Double
Global ViewingKineticParametersForm As Integer  'Parameter telling whether frmInputKineticParameters is showing
Global Cations As AvailableIonsType   'Variable telling whether cations are available for the current resin
Global Anions As AvailableIonsType    'Variable telling whether anions are available for the current resin
Global ClickGeneratedFromcboIon As Integer   'Whether the click event on opt button in Kinetic form was generated from cboIon
Global NumSelectedAnions As Integer
Global NumSelectedCations As Integer
Global NumSelectedComponents_PFPDM As Integer   'Number Of Components currently selected by the user
Global Number_Influent_Points As Integer        'Number of variable influent data points
Global Number_Component As Integer              'Number of components currently selected
Global Total_NumberOfComponents As Integer
Global Anions_Selected(1 To MAX_CHEMICAL) As Integer   ', Anions selected for a simulation
Global Cations_Selected(1 To MAX_CHEMICAL) As Integer  ', Cations selected for a simulation
Global Component_Index_PFPDM(1 To MAX_CHEMICAL) As Integer  ', Ions selected by the user, either anions or cations
Global AlphaInput(1 To MAX_CHEMICAL) As Double    ', Array of separation factors sent to PFPDM DLL
Global EPS_ErrorCriteriaForDGEARIntegrator As Double
Global DH0_InitialTimeStepForDGEARIntegrator As Double
Global Is_Dirty As Boolean
Global kmtest As Boolean

Global NowProj As Project_Type


