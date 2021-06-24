Attribute VB_Name = "moddeclare"
Option Explicit

' the standard operating temperature in case user doesn't enter one
Global Const STANDARD_TEMPERATURE = 25#
' for option buttons
Global Const UNCHECKED = 0
Global Const CHECKED = 1

Global Const MAX_DB = 100

Global Const MAX_PROGRAMS = 10
Global Const MAX_FORMATS = 10
Global Const MAX_AVAILABLE = 10
Global Const MAX_DEF_LINES = 100
' status for the global database
Global Const STATUS_CLOSED = -1
Global Const STATUS_OPEN = 0
Global Const STATUS_CHANGED = 1
' the number of groups and elements a compound can have
Global Const MAX_GROUPS_PER_CHEM = 16
Global Const MAX_GROUPS = 150
Global Const MAX_ELEMENTS = 9
Global Const MAX_IMPORT_CHEMICALS = 20
' the databases (used for database access objects)
Global Const MASTER = 0
Global Const BLOCK5 = 1
' number of properties, methods for each prop, and input for each method we can have
Global Const MAX_PROPERTIES = 50
Global Const MAX_DISPLAY_PROPERTIES = 40    ' the properties we'll actually display (pearls properties)
Global Const MAX_INPUTS = 50    ' really the same thing as MAX_PROPERTIES
Global Const MAX_METHODS_EACH = 6
Global Const MAX_INPUTS_EACH = 6

' input and property constants (actual PPMS properties are 10 through 49
Global Const CONST_U_GROUPS = 41
Global Const CONST_P_GROUPS = 42
Global Const CONST_B_GROUPS = 43
Global Const CONST_L_GROUPS = 44
Global Const CONST_HM_GROUPS = 45
Global Const CONST_ELEMENTS = 46
Global Const CONST_NUM_RINGS = 47
Global Const CONST_REF_CHEM = 48
Global Const CONST_TEMP = 49
Global Const MW = 0         'Molecular Weight
Global Const LD25 = 1       'Liquid Density @ 25C
Global Const LD = 2         'Liquid Density as f(T)
Global Const mp = 3         'Melting Point
Global Const NBP = 4        'Normal Boiling Point
Global Const VP25 = 5       'Vapor Pressure @ 25C
Global Const Vp = 6         'Vapor Pressure as f(T)
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

' the global database (the browser will use)
Global chembrowsedb As Database
Global chembrtable As Recordset
Global dbpath As String
'Global dllpath As String
' the currently selected chemical info
Global selected_name As String
Global selected_cas As Long
Global selected_smiles As String
Global selected_structure As String
Global selected_family As String
Global selected_rings As Integer
Global selected_temperature As Double
Global selected_temp_units As String

' global settings
Global BIPCode(4) As Integer    ' AGLB = 1, AVLE = 2, AENV = 3, ALLE = 4

' the names of inputs and properties as strings for printing purposes
Global input_name(MAX_INPUTS) As String
Global input_enabled(MAX_INPUTS) As Boolean
' the arrays that hold the properties, their methods available, and the inputs each method needs
Global wiz_methods(MAX_PROPERTIES, MAX_METHODS_EACH) As String
Global wiz_inputs(MAX_PROPERTIES, MAX_METHODS_EACH, MAX_INPUTS_EACH) As Integer

Global global_cur_property As Integer ' the current property

Global import_flag As Boolean

Global group_smiles(150) As String

Global dbstatus As Integer  ' the status of the database (use status flags above)
Global def_modified As Boolean

'Global curapp As String
'Global curformat As String
Global curname As String
Global curpath As String

' dbman_(apps, 0) = *.mdb ; dbman_(apps, 1) = path
Global dbman_(MAX_DB, 1) As String
Global dbman_apps As Integer

' the information about persistant tables in master and block 5, should be
' included with data in the copy process
'Dim constant_master_tables(10) As String
'Dim constant_block5_tables(10) As String

Global global_grouptype As String
Global global_groupfile As String

Global global_method As String
Global global_method_file As String
Global cur_chem_groups(21) As Integer
Global num_cur_chem_groups(21) As Integer
' the arrays passed to the shredder dll
'Global intSF_ID() As Long, intSF_Quant() As Long
'Global intMF_ID() As Long, intMF_Quant() As Long

'Global ArrayID() As Long
'Global ArrayQuant() As Long
' the array that will hold import chemical cas numbers
Global import_chemicals(20) As Long

'paul old mosdap
'Declare Sub SubIsomorph Lib "mosdap.dll" (ByVal WordA As String, _
ByRef OneArray As Integer, Array1() As String, ByVal IntVar1 As Integer, _
ByRef IntVar2 As Integer, ByRef TwoArray As Integer, ByRef ThreeArray As Integer)
'Declare Sub SubIsomorph Lib "d:\work\pearls\mosdap.dll" (ByVal WordA As String, ByVal WordB As String, ByVal IntVar1 As Integer, ByRef IntVar2 As Integer, ByRef OneArray As Integer, ByRef TwoArray As Integer)
' the fortran dll routines

'Declare Sub ACCALL Lib "pearls.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTEMP As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups As Long, MS As Long, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)
'Declare Sub ACCALL2 Lib "Pearls.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTemp As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups2 As Long, MS As Long, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)
'Declare Sub AQSCALL Lib "Pearls.dll" (AqueousSolubilty As Double, AqueousSolubiltyShortSource As Long, AqueousSolubiltyLongSource As Long, AqueousSolubiltyError As Long, AqueousSolubiltyTemp As Double, CalculationTemperature As Double, MaximumUnifacGroups As Long, MS As Long, XMW As Double, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)
'Declare Sub AQSCALL2 Lib "pearls.dll" (AqueousSolubilty As Double, AqueousSolubiltyShortSource As Long, AqueousSolubiltyLongSource As Long, AqueousSolubiltyError As Long, AqueousSolubiltyTemp As Double, CalculationTemperature As Double, MaximumUnifacGroups As Long, MS As Long, XMW As Double, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)
'Declare Sub KOWCALL Lib "Pearls.dll" (Kow As Double, KowShortSource As Long, KowLongSource As Long, KowError As Long, KowTemp As Double, CalculationTemperature As Double, FGRPErrorFlag As Long, MaximumUnifacGroups As Long, MS As Long, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)

Declare Sub ACCALL Lib "Environ.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTEMP As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)
Declare Sub AQSCALL Lib "Environ.dll" (AqueousSolubilty As Double, AqueousSolubiltyShortSource As Long, AqueousSolubiltyLongSource As Long, AqueousSolubiltyError As Long, AqueousSolubiltyTemp As Double, CalculationTemperature As Double, MaximumUnifacGroups As Long, MS As Long, XMW As Double, BinaryInteractionParameterDatabase As Long)
Declare Sub HC1CALL Lib "Environ.dll" (HenryCUNIFAC As Double, HCUnifacShortSource As Long, HCUnifacLongSource As Long, HCUnifacError As Long, HCUnifacTemp As Double, OperatingTemp As Double, ActivityCoefficient As Double, VaporPressure As Double)
Declare Sub KOWCALL Lib "Environ.dll" (Kow As Double, KowShortSource As Long, KowLongSource As Long, KowError As Long, KowTemp As Double, CalculationTemperature As Double, FGRPErrorFlag As Long, MaximumUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)

'------------------------------------------------------------------------------------------
'
' Declare: Sub MOSDAP Lib "dlls\Mosdap2.dll"
'
'    The arrays intSF_ID[] and intSF_Quant[] are dimensioned to 100, and the arrays intMF_ID[]
'           and intMF_Quant[] are dimensioned to 21.  To pass them to the .dll, you pass the
'           first element in the array (e.e., intSF_ID[0] which is passed by reference).
'
'    query string  - input of smile string or file name for input file of smile strings
'    Querytype - 0 for string input 1 for file input
'    Subfile - is the name of substructure file (ie unifac.dat)
'    Outfile - is were to output file if querytype is type 1 is writen, file is delimiated (ie. tab)
'    Searchtype -
'           0: Sequential, Non-Truncating
'           1: Sequential, Truncating
'           2: Combinatorial, Truncating
'    Searchresult - flag - fail, pass, or patrial falure
'           0: Unable to disassemble the given smiles string
'               or Error occured in funtion
'           1: Successfully disassembled
'           2: Partially disassembled
'    sf_id array.. subfragment id .. multiple groups seperated by -1 intialized to 0
'    sf_quant arry.. subfragment quantity .. multiple groups seperated by -1 intialized to 0
'    mf_id arry.. molecular feture ... intialized to 0
'    mf_quant arry.. molecular feture ... intialized to 0
'------------------------------------------------------------------------------------------

Declare Sub MOSDAP Lib "Mosdap32.dll" (ByVal Query As String, _
    ByVal QueryType As Byte, ByVal Subfile As String, ByVal Outfile As String, _
    ByVal SearchType As Byte, ByRef SearchResult As Byte, ByRef SF_ID As Long, _
    ByRef SF_Quant As Long, ByRef MF_ID As Long, ByRef MF_Quant As Long)


