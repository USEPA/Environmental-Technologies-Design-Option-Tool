Attribute VB_Name = "Structs"
Option Explicit


Global Const USE_GASPHASE_WAKAO_AND_FUNAZUKRI = False  'False 'True

''''Global Const Application_Name = "AdXsorption Design Software"
''''Global Const name_app = "AdXDesignS"
''''Global Const name_app_long = "AdXsorption Design Software"

'Constants
'Global Const NVersion = 1.3
Global Const NVersion = 1.4
Global Const Latest_DataVersion_Major = 1
Global Const Latest_DataVersion_Minor = 60
Global Const Number_Compo_Max = 10
Global Const MAXCHEMICAL = Number_Compo_Max
Global Const Number_Compo_Max_PFPSDM = 6
Global Const Number_Compo_Max_ECM = 9
Global Const Number_Compo_Max_CPM = 1
Global Const Number_Points_Max = 400
Global Const Number_Data_Points_Max = 400
Global Const Number_Max_Influent_Points = 400
Global Const Max_Number_Correlation_Compo = 25
Global Const Max_Number_Water_Correlations = 25
Global Const CPM_Max_Points = 100
Global Const PI = 3.14159265359

  'Modified Hokanson 2/8/97
  'Global Const Max_Radial_Collocation = 6
Global Const Max_Radial_Collocation = 18
Global Const Max_Equations_DGEAR = 750
  'end Modified Hokanson 2/8/97

Global Const Max_Axial_Collocation = 18
Global Const Max_Number_Fouling_Iterations = 100
Global Const Maximum_Beds_In_Series = 200
Global Const EPS_ERROR_CRITERIA = 0.0005

Type Tempo_Data
    MW As Double
    Solubility As Double
    Density As Double
    R_Index As Double
    T As Double
    Pvap As Double
End Type

Type IPES_Input
    Adsorbent As String
    BB As Double
    W0 As Double
    RH As Double    'Relative Humidity
    C As Double     'Concentration
    Phase As String * 6
    IMOD As Integer
    GM As Double
    OMAG As Double
    NL As Integer
End Type

Type IPES_Output
    XN As Double
    XK1 As Double
    XK2 As Double
    CSAV As Double
    QSAV As Double
    CBEG As Double
    CEND As Double
    RSQD As Double
    RMSE As Double
    Error_Matrix(30) As Integer
    CorrelationPoints_lnC(1 To 200) As Double
    CorrelationPoints_lnQ(1 To 200) As Double
    CorrelationPoints_NumPoints As Integer
    QCAP(1 To 200) As Double
    ADSP(1 To 200) As Double
    PI(1 To 200) As Double
End Type

Type IPES_Variable
    Input As IPES_Input
    Output As IPES_Output
End Type

Type Correlation_Compound_Type
    Name As String
    Coeff(2) As Double
End Type

Type Correlation_Water_Type
    Name As String
    Coeff(4) As Double
End Type

Global Const KNSOURCE_ISOTHERMDB = 1
Global Const KNSOURCE_IPES = 2
Global Const KNSOURCE_USERINPUT = 3


Global Const IPESMETHOD_LIQ_3PARAM = 1
Global Const IPESMETHOD_LIQ_DRUNIFORM = 2
Global Const IPESMETHOD_LIQ_DRNONUNIFORM = 3
Global Const IPESMETHOD_GAS_DRZERORH = 101
Global Const IPESMETHOD_GAS_CALGONBPL = 102
Global Const IPESMETHOD_GAS_DRSPREADINGP = 103


Type ComponentPropertyType
'Isotherm Freundlich q=k*C^OneOverN
    
    '***** Properties: *****
    Name As String * 50
    CAS As Long
    MW As Double                    'g/mol
    MolarVolume As Double           'cm3/mol
    BP As Double                    'degrees Celcius
    InitialConcentration As Double  'mg/l
    Use_K As Double                 '(mg/g)*(L/mg)^OneoverN
    Use_OneOverN As Double          '(-)
    Source_KandOneOverN As Integer  'Isotherm DB / IPES / User-Input
    UserEntered_K As Double
    UserEntered_OneOverN As Double
    Treatment_Objective As Double
    '***** IPES RELATED: *****
    IPES_OrderOfMagnitude As Double
    IPES_NumRegressionPts As Integer
    IPES_RelativeHumidity As Double
    IPES_EstimationMethod As Integer
    Liquid_Density As Double        'g/cm^3
    Aqueous_Solubility As Double    'mg/L
    Vapor_Pressure As Double        'Pa
    Refractive_Index As Double      '(-)
    IPESResult_K As Double
    IPESResult_OneOverN As Double
    
    '***** Isotherm Database: *****
    IsothermDB_Component_Name As String * 70
      '(note: includes CAS and Name exactly as it appears on the Freundlich Isotherm Parameters form)
    IsothermDB_Range_Num As Integer
      '(this is the list index [1-n] of the selected range.)
    IsothermDB_K As Double
    IsothermDB_OneOverN As Double

    '***** Kinetics *****
    SPDFR As Double
    SPDFR_Low_Concentration As Integer
    Use_SPDFR_Correlation As Integer
    Corr(3) As Integer
    kf As Double                    'cm/s
    Ds As Double                    'cm2/s
    Dp As Double                    'cm2/s
    KP_User_Input(3) As Double
    Tortuosity As Double
    Use_Tortuosity_Correlation As Integer
    Constant_Tortuosity As Integer

    '***** ECM *****
    K_Reduction As Integer          'Boolean
    Correlation As Correlation_Compound_Type
                                    'Correlation to calculate K reduction

    Is_Selected_On_List As Boolean
        'TEMPORARY INTERNAL VARIABLE: NOT SAVED.
End Type

Type PropertyUnitsType
    MW As String
    MolarVolume As String
    BP As String
    InitialConcentration As String
    Liquid_Density As String
    Aqueous_Solubility As String
    Vapor_Pressure As String
    Refractive_Index As String
    k As String
    BedTemperature As String
    BedPressure As String
    BedFluidDensity As String
    BedFluidViscosity As String
End Type
Global PropertyUnits As PropertyUnitsType

Type BedPropertyType
    length As Double            'm
    Diameter As Double          'm
    Weight As Double            'kg
    Flowrate As Double          'm3/s
    Density As Double           'g/cm3
    SuperficialVelocity As Double  'm/s
    Porosity As Double          '(-)
    InterstitialVelocity As Double 'm/s
    Area As Double              'm2
    Volume As Double            'm3
    TAU As Double               'min (packed bed contact time)
    NumberOfBeds As Integer     '(-)
    WaterDensity As Double      'g/cm3
    WaterViscosity As Double    'g/cm.s
    Temperature As Double       'C
    Pressure As Double          'Atm
    Phase As Integer            '=0 -> liquid, =1 -> gas
    Water_Correlation As Correlation_Water_Type   'Correlation to calcualte K reduc

    UnitsLength As Integer
    UnitsDiameter As Integer
    UnitsWeight As Integer
    UnitsFlowrate As Integer
    UnitsEBCT As Integer
    UnitsFluidDensity As Integer
    UnitsFluidViscosity As Integer
    UnitsFluidTemperature As Integer
    UnitsFluidPressure As Integer
End Type

Type CarbonPropertyType
    'Manu As String
    Name As String
    Porosity As Double         ' -
    Density As Double          'g/cm3  Apparent density
    ParticleRadius As Double   'm
    Tortuosity As Double       ' -            'UNUSED!!!
' Need to add W0, BB, Polanyi Exponent here!
' It appears safe to add these variables.
    W0 As Double
    BB As Double
    PolanyiExponent As Double

    '---- Added by EJO on 11/1/96 for kf (external mass xfer coefficient)
    ShapeFactor As Double       ' -
End Type

Type DataCarbon
    Index As Integer
    CAS As Long
    Name As String * 50
    NameC As String * 20
    pHmin As Double
    pHmax As Double
    k As Double
    N As Double 'Actually, this is 1/n
    Cmin As Double
    Cmax As Double
    Temperature As Double
    Phase As String * 50
    Source As String * 50
    Comments As String * 50
End Type

Type Para_Int
    Init As Double
    End As Double
    np As Integer
    Step As Double
End Type

Type Throughput
    T As Double
    C As Double
End Type

Type ResultsType
    NComponent As Integer
    npoints As Integer
    CP(Number_Compo_Max, Number_Points_Max) As Double
    T(Number_Points_Max) As Double
    ThroughPut_05(Number_Compo_Max) As Throughput
    ThroughPut_50(Number_Compo_Max) As Throughput
    ThroughPut_95(Number_Compo_Max) As Throughput
    Component(Number_Compo_Max) As ComponentPropertyType
    Bed As BedPropertyType
    Carbon As CarbonPropertyType
    Use_Tortuosity_Correlation As Integer
    Constant_Tortuosity As Integer
    NumPoints_Before_ThroughPut_100(Number_Compo_Max) As Integer       'Used by PSDM to cut off display when C/C0 >= 1
    is_psdm_in_room_model As Integer
    int_Which_PSDMR_Model As Integer
    psdmroom_Crss(1 To Number_Compo_Max) As Double       'ug/L
    AnyCrCloseToZero As Integer
End Type

Type PSDMInputsType
    VARS1(1 To 15) As Double
    VARS2(1 To Number_Compo_Max, 1 To 19) As Double
End Type

Type ECM_Data
    Index As Integer
    Bed_Volume_Fed As Double
    Wave_Velocity As Double
    Dimensionless_Bed_Length As Double
    SS_Treatment_Capacity As Double
    Solid_Concentration(Number_Compo_Max) As Double
    Liquid_Concentration(Number_Compo_Max)  As Double
    C_Over_C0(Number_Compo_Max) As Double
    Carbon_Usage_Rate As Double
End Type

Type ECM_MASSBAL
    MASSBAL_C0_e_Vf(Number_Compo_Max) As Double
    MASSBAL_TERM_SUM(Number_Compo_Max) As Double
    MASSBAL_PERCENT_ERR(Number_Compo_Max) As Double
End Type
Global Output_ECM_MASSBAL As ECM_MASSBAL

Type CPM_Data
    T(CPM_Max_Points) As Double
    C_Over_C0(CPM_Max_Points)  As Double
    Par(7) As Double
    ThroughPut_05 As Throughput
    ThroughPut_50 As Throughput
    ThroughPut_95 As Throughput
    Component As ComponentPropertyType
    Bed As BedPropertyType
    Carbon As CarbonPropertyType
    Use_Tortuosity_Correlation As Integer
    Constant_Tortuosity As Integer
End Type

Type Isotherm_Data
    number As Integer
    Selected As Long
    Record As DataCarbon
End Type

Type Isotherm_Data_Save
    New_CAS As Integer
    New_Name As Integer
    Record As DataCarbon
End Type

Type Isotherm_Chemical
    CAS As Long
    Name As String * 50
    Update_Name As Integer
    Update_CAS As Integer
End Type

Type Carbon_Data
    Name As String * 50
    Density As Double
    ParticleRadius As Double
    Porosity As Double
    ShapeFactor As Double
    Tortuosity As Double                  'UNUSED!
    Phase As Integer '1=Liquid, 2 =gas
    Type As String * 50
    W0 As Double
    b As Double
    PolanyiExponent As Double
End Type

'Variables for Help System
Global HelpFile As String

'Variable for security
Global Open_Database As Integer

'Variables
Global PFPSDM_Path As String, Error_In_Kinetic_Calculation As Integer
Global Database_Path As String
Global Exe_Path As String     'NEW 9/2/98.
Global Flag_Openfile As Integer
Global Update_Value_From_Carbon As Integer, Update_Value_From_IPES As Integer
Global Temp_Text As String 'Temporary string to store former value of a text box
Global AddFlag As Integer 'Flag - True = Add a chemical to the list - False = Edit chemical properties
Global Component_Number_Selected As Integer
Global State_Check_Water(2) As Integer  'Flag to know whether correlations are used for water properties
Global Filename As String, Previous_FileName As String
Global Results As ResultsType
Global PSDM_Inputs As PSDMInputsType
Global Print_To_Printer As Integer 'Flag - True = print to Printer - False = Print to file
Global NL As String
Global batchrun As Integer
Global Treatment_Objective(Number_Compo_Max_PFPSDM) As Throughput

'Variable for Tortuosity as a function of Time
Global Use_Tortuosity_Correlation As Integer, Constant_Tortuosity As Integer

'Variable for the search function
Global Start_Search As Integer
Global Find_String As String, Index_NameT As Integer, Index_Find As Integer

'Variables for K reduction
Global Number_Correlations_Compounds As Integer
Global Number_Water_Correlations As Integer
Global Correlations_For_Classes(Max_Number_Correlation_Compo) As Correlation_Compound_Type
Global Correlations_For_Water(Max_Number_Water_Correlations) As Correlation_Water_Type

'Variable used to transfer data to excel
Global Excel_4 As Integer
Global PFPSDM_Excel As Integer

'Variables used to edit the isotherm database
Global Mode_Chemical As Integer, Mode_Isotherm As Integer
Global Iso_Data As Isotherm_Data
Global Iso_Data_Save As Isotherm_Data_Save
Global Iso_Chemical As Isotherm_Chemical

'Variables used to edit the carbon database
Global Mode_Manu As Integer
Global Mode_Carbon As Integer
Global Name_Manufacturer_In As String
Global Name_Manufacturer_Out As String
Global Carbon_Data_In As Carbon_Data
Global Carbon_Data_Out As Carbon_Data

'Variables for ECM model
Global Output_ECM(Number_Compo_Max) As ECM_Data
Global Number_Component_ECM As Integer
Global Component_Index_ECM(Number_Compo_Max_ECM)

'Variables for the Contant Pattern Model
Global CPHSDM_Excel As Integer
Global Number_Component_CPM As Integer
Global Component_Index_CPM As Integer
Global CPM_Results As CPM_Data

'Variables for the PFPSDM FORTRAN program
Global C_Influent(Number_Compo_Max, Number_Max_Influent_Points) As Double, T_Influent(Number_Max_Influent_Points) As Double
Global Number_Influent_Points As Integer
Global TimeP As Para_Int
Global MC As Integer, NC As Integer
Global Number_Component As Integer
Global Bed As BedPropertyType
Global Component(0 To Number_Compo_Max) As ComponentPropertyType 'Component(0) is a component used for temporary storage. Component(1) to Component(N) are the components in the listbox
Global Carbon As CarbonPropertyType

Global Component_Index_PFPSDM(Number_Compo_Max_PFPSDM) As Integer
Global Number_Component_PFPSDM As Integer

Global IPES_Data As IPES_Variable
Global Properties_For_IPES As Tempo_Data

'Flags to avoid problems when loading a file...
Global Use_Update_Display_Kinetic As Integer

'Flags to check whether or not the user wants the data to be saved before exiting
' in the QueryUnload event from the main window
Global ReallyQuit  As Integer

'Arrays to store the experimental data points for comparison to a PSDM simulation
Global T_Data_Points(Number_Data_Points_Max)   As Double
Global C_Data_Points(Number_Compo_Max, Number_Data_Points_Max) As Double
Global NData_Points As Integer
'Global NComponents As Integer

'Variable to store isotherm parameters to plot it
Global IsothermProperties As DataCarbon

'Variable to tell frmBatch its default model to simulate
Global BatchSimulation_DefaultModel As Integer

'Variables to tell frmStEPPImport what to do.
Global Const STEPPIMPORT_ADDCOMPONENTS = 1
Global Const STEPPIMPORT_IPESCOMPONENT = 2
Global StEPPImportDestination As Integer
Global StEPPImportSuccess As Integer
Type StEPP_to_IPES_Properties_type
  Name As String
  MW As Double
  MolarVolume As Double
  BP As Double
End Type
Global StEPP_to_IPES_Properties As StEPP_to_IPES_Properties_type

