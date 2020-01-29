Attribute VB_Name = "modDLLDeclare"
Option Explicit
'This module contains the DLL declarations needed
'to make calls to FORTRAN routines within the
'program

Declare Sub ACCALL Lib "pearls.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTEMP As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups As Long, MS As Long, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)
Declare Sub ACCALL2 Lib "pearls.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTEMP As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups2 As Long, MS As Long, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)
Declare Sub AQSCALL Lib "pearls.dll" (AqueousSolubilty As Double, AqueousSolubiltyShortSource As Long, AqueousSolubiltyLongSource As Long, AqueousSolubiltyError As Long, AqueousSolubiltyTemp As Double, CalculationTemperature As Double, MaximumUnifacGroups As Long, MS As Long, XMW As Double, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)
Declare Sub AQSCALL2 Lib "pearls.dll" (AqueousSolubilty As Double, AqueousSolubiltyShortSource As Long, AqueousSolubiltyLongSource As Long, AqueousSolubiltyError As Long, AqueousSolubiltyTemp As Double, CalculationTemperature As Double, MaximumUnifacGroups As Long, MS As Long, XMW As Double, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)
Declare Sub KOWCALL Lib "pearls.dll" (Kow As Double, KowShortSource As Long, KowLongSource As Long, KowError As Long, KowTemp As Double, CalculationTemperature As Double, FGRPErrorFlag As Long, MaximumUnifacGroups As Long, MS As Long, MainGroup As Long, AIM As Double, RIM As Double, QIM As Double, MWM As Double, MVM As Double)

'paul: old mosdap
'Declare Sub SubIsomorph Lib "a:\mosdap.dll" (ByVal WordA As String, ByVal WordB As String, ByVal IntVar1 As Integer, ByRef IntVar2 As Integer, ByRef OneArray As Integer, ByRef TwoArray As Integer)
'Declare Sub SubIsomorph Lib "a:\crap.dll" (ByVal WordA As String, ByRef OneArray As Integer, Array1() As String, ByVal IntVar1 As Integer, ByRef IntVar2 As Integer, ByRef TwoArray As Integer, ByRef ThreeArray As Integer)
'Declare Sub SubIsomorph Lib "mosdap.dll" (ByVal WordA As String, ByRef OneArray As Integer, Array1() As String, ByVal IntVar1 As Integer, ByRef IntVar2 As Integer, ByRef TwoArray As Integer, ByRef ThreeArray As Integer)

' ---- New dll calls from StEPP ----
'Declare Sub ACCALL Lib "dlls\Environ.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTEMP As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)
'Declare Sub AQSCALL Lib "dlls\Environ.dll" (AqueousSolubilty As Double, AqueousSolubiltyShortSource As Long, AqueousSolubiltyLongSource As Long, AqueousSolubiltyError As Long, AqueousSolubiltyTemp As Double, CalculationTemperature As Double, MaximumUnifacGroups As Long, MS As Long, XMW As Double, BinaryInteractionParameterDatabase As Long)
'Declare Sub HC1CALL Lib "dlls\Environ.dll" (HenryCUNIFAC As Double, HCUnifacShortSource As Long, HCUnifacLongSource As Long, HCUnifacError As Long, HCUnifacTemp As Double, OperatingTemp As Double, ActivityCoefficient As Double, VaporPressure As Double)
'Declare Sub KOWCALL Lib "dlls\Environ.dll" (Kow As Double, KowShortSource As Long, KowLongSource As Long, KowError As Long, KowTemp As Double, CalculationTemperature As Double, FGRPErrorFlag As Long, MaximumUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)

'Mosdap
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
