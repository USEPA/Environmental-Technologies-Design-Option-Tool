Attribute VB_Name = "modeclare"
'Global variables used in PEARLS

'Current version number
Global Const VersionNumber = "1.20"

'Number of available properties and methods
'Global Const NumProperties = 40 '(there are actually 41)

'mrt- NumProperties is not a very good thing to have in the code
'       all references to this should eventually be changed!!!

'14 more props now that tabs 7 and 8 were added
Global Const NumProperties = 54 '(there are actually 55)

Global Const NumMethods = 12

'A boolean value used to keep track of the file currently in use
Global WorkModified As Boolean

'Variables used for database file access
Global DBJetMaster As Database
Global DBJetUser As Database
Global DIPPR801 As Boolean
Global DIPPR911 As Boolean
Global UserDBName As String
Global MasterDBName As String
Global SaveFileName As String
Global default_master_name As String

'paths needed (initialized in ModMain.bas: Initialize_File_Stuff)
Global PathUser As String   ' constant at app.path & "\dbuser.mdb"
Global PathSave As String
Global PathDemo As String   ' constant at app.path & "\demo.prl"
Global PathMaster As String
Global Path801 As String    ' always = PathMaster
Global Path911 As String    ' always = PathMaster
Global PathReport As String ' constant at app.path
Global PathBlock5 As String

' the number of files needed
Global Const numrptfiles = 3
Global Const numuserfiles = 1
Global Const numsavefiles = 1
Global Const numdemofiles = 1
Global Const nummasterfiles = 1
Global Const numdb801files = 1
Global Const numdb911files = 1
Global Const numdbblock5files = 1

' variables to hold the files needed

Global rptfile(numrptfiles) As String
Global userfile(numuserfiles) As String
Global savefile(numsavefiles) As String
Global demofile(numdemofiles) As String
Global masterfile(nummasterfiles) As String
Global db801file(numdb801files) As String
Global db911file(numdb911files) As String
Global dbblock5file(numdbblock5files) As String
Global deffile As String

'Property Codes
Global Const OptTemp = -1   'Operating tempreture
Global Const OptPress = -2  'Operating Pressure

Global Const MW = 0         'Molecular Weight
Global Const LD25 = 1       'Liquid Density @ 25C
Global Const LD = 2         'Liquid Density as f(T)
Global Const mp = 3         'Melting Point
Global Const NBP = 4        'Normal Boiling Point
Global Const VP25 = 5       'Vapor Pressure @ 25C
Global Const VP = 6         'Vapor Pressure as f(T)
Global Const hfor = 7       'Heat of Formation
Global Const LHC = 8        'Liquid Heat Capacity as f(T)
Global Const VHC = 9        'Vapor Heat Capacity as f(T)
Global Const Hvap25 = 10    'Heat of Vaporization @ 25C
Global Const HvapNBP = 11   'Heat of Vaporization @ NBP
Global Const Hvap = 12      'Heat of Vaporization as f(T)
Global Const CT = 13        'Critical Temperature
Global Const CP = 14        'Critical Pressure
Global Const CV = 38        'Critical Volume
Global Const Dwater = 15    'Diffusivity in Water
Global Const Dair = 16      'Diffusivity in Air
Global Const ST25 = 17      'Surface Tension @ 25C
Global Const ST = 18        'Surface Tension as f(T)
Global Const VV = 19        'Vapor Viscosity as f(T)
Global Const LV = 20        'Liquid Viscosity as f(T)
Global Const LTC = 21       'Liquid Thermal Conductivity as f(T)
Global Const VTC = 22       'Vapor Thermal Conductivity as f(T)
Global Const UFL = 23       'Upper Flammability Limit
Global Const LFL = 24       'Lower Flammability Limit
Global Const FP = 25        'Flash Point
Global Const AIT = 26       'Autoignition Temperature
Global Const Hcomb = 27     'Heat of Combustion
Global Const ThODcarb = 28  'Carbonaceous ThOD
Global Const ThODcomb = 29  'Combined ThOD
Global Const COD = 30       'Chemical Oxygen Demand
Global Const BOD = 31       'Biochemical Oxygen Demand
Global Const ACwater = 32   'Infinite Dilution Activity Coefficient of Water in Chemical
Global Const HC = 33        'Henry's Constant
Global Const ACchem = 34    'Infinite Dilution Activity Coefficient of Chemical in Water
Global Const logKow = 35    'log Kow
Global Const logKoc = 36    'log Koc
Global Const BCF = 37       'Bioconcentration Factor
Global Const Schem = 39     'Solubility Limit of Chemical in Water
Global Const Swater = 40    'Solubility Limit of Water in Chemical
    '41 to 54 is the new toxicity props added approx 8/14/98 JEM
Global Const Fat48E = 41    'Fathead Minnow, 48h, EC50
Global Const Fat96E = 42    'Fathead Minnow, 96h, EC50
Global Const Fat24L = 43    'Fathead Minnow, 24h, LC50
Global Const Fat48L = 44    'Fathead Minnow, 48h, LC50
Global Const Fat96L = 45    'Fathead Minnow, 96h, LC50
Global Const Sal24L = 46    'Salmonidae, 24h, LC50
Global Const Sal48L = 47    'Salmonidae, 48h, LC50
Global Const Sal96L = 48    'Salmonidae, 96h, LC50
Global Const Daph24E = 49    'Daphnia magna, 24h, EC50
Global Const Daph48E = 50    'Daphnia magna, 48h, EC50
Global Const Daph24L = 51    'Daphnia magna, 24h, LC50
Global Const Daph48L = 52    'Daphnia magna, 48h, LC50
Global Const Mysid96L = 53    'Mysid, 96h, LC50
Global Const AltSpecies = 54    'Alternate species

' mrt 3/24/99
'int for antoine coefficients, for use mostly with printing
Global Const ANT = 55

' some standard values
Global Const STANDARD_K_TEMP = 298.15
Global Const ERROR_FLAG = -99999.9
Global Const STANDARD_ATM_PRESSURE = 1#
Global Const WATER_MW = 18.015

' other properties not stored but calculated
' for ASPEN:
Global Const MV As Integer = 45     ' Molar Volume at Tb
Global Const omega As Integer = 46  ' Acentric Factor
Global Const CPIGa As Integer = 47  ' Ideal Gas Heat Capacity
Global Const PLXANTb As Integer = 48 ' Antoine Eqn Parameters
Global Const ZC As Integer = 49     ' Critical Compressibility Factor

'This defines information for each method
Type MethodInfoType
    
    'Method information
    CurMethod As Integer
    MethodName(NumMethods) As String * 25
    value(NumMethods) As Double
    Enabled(NumMethods) As Boolean
    Unit As String * 15
        
    'f(T) data for each method
    EqNum(NumMethods) As Integer
    Coeff(NumMethods, 5) As Double
    MinT(NumMethods) As Double
    MaxT(NumMethods) As Double
            
    'Temperature and units for f(T) methods
    TFT As Double
    TFTUnit As String * 5
    
End Type

'This defines information on the current chemical
Type CurInfoType
    
    'Basic chemical information
    CAS As Long
    name As String * 50
    source As String * 10
    Formula As String * 20
    SMILES As String * 20
    Family As String * 5
    
    'Fragment Information
    NumRings As Integer
    MaxGroups  As Integer
 'paul
    Grp(99) As Long
    NumGrp(99) As Long
    
    'Operating conditions
    OpT As Double
    OpTUnit As String * 5
    OpP As Double
    OpPUnit As String * 10
    
End Type

'mrt- this defines information for antoine property (used mainly for printing)
Type AntoineInfoType
    
    'type information
    AntCalc As Boolean 'indicates weather or not antoine has been calculated
    AntType As Boolean 'true = regular coefficients, false = regressed coeffs
    
    'common data
    A As String * 10
    B As String * 10
    C As String * 10
    D As String * 10
    E As String * 10
    EqNum As String * 10
    TFTUnit As String * 10
    TMin As String * 10
    TMax As String * 10
    MethodName As String * 25
    
    'data for regression only
    value As String * 10
    Unit As String * 10
    TFT As String * 10
End Type

'Declare cur_info variables for storing chemical data
    'holds all data in ONLY internal standard
Global Cur_Info As CurInfoType

'mrt- Declare antoine_info to hold antoine info for printing. This is only
'       updated when user calculates ant coefficients individually. This info is
'       deleted if the user switches chemicals (i.e. picks a new chem and hits
'       the calculate button on the main page)
Global Antoine_Info As AntoineInfoType
'mrt- Declare print_antoine to tell the print functions weather or not to print
'       the antoine stuff. This is again a special case.
Global print_antoine As Boolean

'Define storage for each method
Global InfoMethod(NumProperties) As MethodInfoType

'Variable for current property selected
Global CurProp As Integer



'Variables used for user preferences
Global DefaultUnit(NumProperties) As String
Global DefaultTFTUnit As String
Global TFTConvert As Boolean
Global GraphConvert As Boolean
Global SetDefaultUnit As Boolean
Global FormatGT1000 As String
Global FormatLT001 As String
Global FormatGeneral As String

'Variables for TFT cancel
Global TempTFT As Double
Global TempTFTUnit As String * 5

'Variables for UNIFAC Calculations
Global MGSG(1 To 116) As Long               'Main group number
Global RI(1 To 116) As Double               'Volume parameter
Global QI(1 To 116) As Double               'Area parameter
Global MWS(1 To 116)  As Double             'Molecular Weight array for UNIFAC groups
Global MVS(1 To 116) As Double              'Molar Volume array for UNIFAC groups

'Variables for BIP hierarchy currently selected BIPs
Global BIP(4, 1 To 58, 1 To 58) As Double   'Binary interaction parameter matrix
Global BIPHierarchy(3, 4) As Integer        'BIP calculation hierarchy
Global BIPIndex(3) As Integer               'Index to currently selected BIP database

'Variables for block 5 preferences
Global B5Preference(4, NumMethods) As Integer

'Variables for unit conversions
Global Unit1(200) As String
Global Unit2(200) As String
Global AddProp(200) As String
Global Mult(200) As Double
Global AddConst(200) As Double
Global AddPropOpFlag(200) As String

'Variables for calculation hierarchy
Global CalcHierarchy(NumProperties) As Integer

'Variables for method screen
Global PrevProp(10) As Integer
Global ScreenNum As Integer

'Variables for sort option
Global SortChemListAsc As Boolean
Global SortUserListAsc As Boolean

'Variables for environment preferences
Global PrefStartup As Boolean


