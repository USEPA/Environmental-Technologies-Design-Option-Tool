Attribute VB_Name = "modblock5"
Option Explicit

Global Block5DB As Database
Global databasename As String
Dim inorganic As Boolean
Global Const NULLCODE = -1                  ' this is because we can't use 0 as an empty code
Dim ntel, ktel, ktgp, ntgp As Integer
Dim NGPT(100) As Integer
Dim NET(120) As Integer
    
Dim FMLDI, FMLGP, FMLCR, FMLFP As Double    ' lfl values needed by flashpoint calcs
    
    ' hard-coded in right now (in initialize function)
Dim EL(120) As String           ' The elements: initialized in Initialize function
Dim CNU(120) As Double          ' initialized in Initialize function

Dim QB(21) As String            ' The possible quality codes: initialized in Initialize function
    
    ' globals, used by all four functions - from database
Dim elr(7) As String        ' chemical symbol for each element in compound
Dim nelr(7) As Integer      ' number of atoms of each element in compound
Dim IGPCOD(16) As Integer    ' group number for each chemical group in compound
Dim NGP(16) As Integer       ' number of each group in compound

Dim bl(10, 16) As String    ' quality code
Dim cat(2) As String        ' category code

Dim ICODE As Long      ' [Dippr Code] in db
Dim CHEM_NAME As String
    ' values from fire and explosion data file
Dim FML As Double  ' lower flammability limit
Dim FLP As Double   ' flashpoint
Dim FMU As Double  ' upper flammability limit
Dim AITemp As Double  ' autoignition temp
Dim hcom, hfor, gfor, CMW As Double

    ' These are globals that hold data from the main table
    ' (would have been from lines 4-10 in the original data file)
Dim Pc As Double            ' critical pressure of compound
Dim neqvp As Integer        ' equation number for vapor pressure
Dim vpc(5) As Double        ' regression coeff for vapor pressure

Dim Tc As Double            ' critical temperature of compound
Dim neqhv As Integer        ' equation number for heat of vaporization
Dim hvapc(5) As Double      ' regression coeff for heat of vaporization

Dim TBP As Double           ' normal boiling point of compound
Dim neqcpg As Integer       ' equation number for heat capacity of gas
Dim cpgc(5) As Double       ' regression coeff for gas phase heat capacity

Dim mp As Double            ' melting point of compound (was TMP)
Dim neqcpl As Integer       ' equation number for heat capacity of liquid
Dim cplc(5) As Double       ' regression coeff for liquid phase heat capacity

Dim HFUS As Double          ' heat of fusion
Dim neqcps As Integer       ' equation number for heat capacity of solid
Dim cpsc(5) As Double       ' regression coeff for solid phase heat capacity

Dim VLIQ As Double          ' molar volume of liquid
Dim neqdnl As Integer       ' equation number for liquid density
Dim denlc(5) As Double      ' regression coeff for liquid density

Dim neqdns As Integer       ' equation number for solid density
Dim densc(5) As Double      ' regression coeff for solid density
Dim QC(5) As String
'error codes
Dim PC_ECODE As String
Dim TC_ECODE As String
Dim TBP_ECODE As String
Dim MV_ECODE As String
Dim mp_ecode As String
Dim HFUS_ECODE As String
Dim VLIQ_ECODE As String
Public Sub Initialize()

    '  3/23/97 : DMW  fixed array  indexes so that arrays start at 0
    ' A function called by read_database load to initialize some globals
    
Dim i, J As Integer
    ' This is the global that is used as a delimiter in the data file????
For i = 0 To 10
    For J = 0 To 15
        bl(i, J) = " "
    Next J
Next i

For i = 0 To 6
    nelr(i) = 0
    elr(i) = " "
Next i
For i = 0 To 15
    NGP(i) = 0
    IGPCOD(i) = 0
Next i
cat(0) = " "
cat(1) = " "
QB(0) = "A"
QB(1) = "B"
QB(2) = "C"
QB(3) = "D"
QB(4) = "E"
QB(5) = "F"
QB(6) = "G"
QB(7) = "H"
QB(8) = "I"
QB(9) = "X"
QB(10) = "J"
QB(11) = "K"
QB(12) = "L"
QB(13) = "M"
QB(14) = "N"
QB(15) = "P"
QB(16) = "Q"
QB(17) = "R"
QB(18) = "S"
QB(19) = "U"
QB(20) = "V"
EL(0) = "A"
EL(1) = "Ac"
CNU(1) = 0.75
EL(2) = "Ag"
CNU(2) = 0.25
EL(3) = "Al"
CNU(3) = 0.75
EL(4) = "Am"
CNU(4) = 0.75
EL(5) = "As"
CNU(5) = 0.75
EL(6) = "At"
EL(7) = "Au"
CNU(7) = 0.75
EL(8) = "B"
CNU(8) = 0.75
EL(9) = "Ba"
CNU(9) = 0.5
EL(10) = "Be"
CNU(10) = 0.5
EL(11) = "Bi"
CNU(11) = 0.75
EL(12) = "Bk"
EL(13) = "Br"
EL(14) = "C"
CNU(14) = 1#
EL(15) = "Ca"
CNU(15) = 0.5
EL(16) = "Cd"
CNU(16) = 0.5
EL(17) = "Ce"
CNU(17) = 1#
EL(18) = "Cf"
EL(19) = "Cl"
EL(20) = "Cm"
EL(21) = "Co"
CNU(21) = 0.75
EL(22) = "Cr"
CNU(22) = 1.5
EL(23) = "Cs"
CNU(23) = 0.25
EL(24) = "Cu"
CNU(24) = 0.5
EL(25) = "D"
CNU(25) = 0.25
EL(26) = "Dy"
EL(27) = "Er"
EL(28) = "Eu"
EL(29) = "F"
EL(30) = "Fe"
CNU(30) = 0.75
EL(31) = "Fr"
CNU(31) = 0.25
EL(32) = "Ga"
CNU(32) = 0.75
EL(33) = "Gd"
EL(34) = "Ge"
CNU(34) = 1#
EL(35) = "H"
CNU(35) = 0.25
EL(36) = "He"
EL(37) = "Hf"
CNU(37) = 1#
EL(38) = "Hg"
CNU(38) = 0.5
EL(39) = "Ho"
EL(40) = "I"
EL(41) = "In"
CNU(41) = 0.75
EL(42) = "Ir"
CNU(42) = 0.75
EL(43) = "K"
CNU(43) = 0.25
EL(44) = "Kr"
EL(45) = "La"
CNU(45) = 0.75
EL(46) = "Li"
CNU(46) = 0.25
EL(47) = "Lu"
EL(48) = "Mg"
CNU(48) = 0.5
EL(49) = "Mn"
CNU(49) = 1#
EL(50) = "Mo"
CNU(50) = 1.5
EL(51) = "Mv"
EL(52) = "N"
CNU(52) = 0.5
EL(53) = "Na"
CNU(53) = 0.25
EL(54) = "Nb"
CNU(54) = 1.25
EL(55) = "Nd"
CNU(55) = 0.75
EL(56) = "Ne"
EL(57) = "Ni"
CNU(57) = 0.5
EL(58) = "Np"
CNU(58) = 1.25
EL(59) = "O"
CNU(59) = -0.5
EL(60) = "Os"
CNU(60) = 1#
EL(61) = "P"
CNU(61) = 1.25
EL(62) = "Pa"
CNU(62) = 1.25
EL(63) = "Pb"
CNU(63) = 0.5
EL(64) = "Pd"
CNU(64) = 0.5
EL(65) = "Pm"
EL(66) = "Po"
CNU(66) = 1#
EL(67) = "Pr"
EL(68) = "Pt"
CNU(68) = 0.5
EL(69) = "Pu"
CNU(69) = 0.5
EL(70) = "Ra"
CNU(70) = 0.5
EL(71) = "Rb"
CNU(71) = 0.25
EL(72) = "Re"
CNU(72) = 1#
EL(73) = "Rh"
CNU(73) = 0.75
EL(74) = "Rn"
EL(75) = "Ru"
EL(76) = "S"
CNU(76) = 1#
EL(77) = "Sb"
CNU(77) = 0.75
EL(78) = "Sc"
CNU(78) = 0.75
EL(79) = "Se"
CNU(79) = 1.5
EL(80) = "Si"
CNU(80) = 1#
EL(81) = "Sm"
CNU(81) = 0.75
EL(82) = "Sn"
CNU(82) = 1#
EL(83) = "Sr"
CNU(83) = 0.5
EL(84) = "Ta"
CNU(84) = 1.25
EL(85) = "Tb"
EL(86) = "Tc"
CNU(86) = 1#
EL(87) = "Te"
CNU(87) = 1#
EL(88) = "Th"
CNU(88) = 1#
EL(89) = "Ti"
CNU(89) = 1#
EL(90) = "Tl"
CNU(90) = 0.75
EL(91) = "Tm"
EL(92) = "U"
CNU(92) = 1#
EL(93) = "V"
CNU(93) = 1.25
EL(94) = "W"
CNU(94) = 1.5
EL(95) = "Xe"
EL(96) = "Y"
CNU(96) = 0.75
EL(97) = "Yb"
EL(98) = "Zn"
CNU(98) = 0.5
EL(99) = "Zr"
CNU(99) = 1#

End Sub








Public Sub calc_flpt(CNT As Integer)
    
    ' calc_flpt:  a function that calculates the flashpoint and puts the
    '           values in the text boxes on the main form
    '           arrays are indexed from 0 instead of 1 (as in the fortran code)
        
    ' 3/23/97 :  DMW  fixed array indexes to start from 0
    ' 5/30/97 :  DMW  CNT is the position in the method array to put the first result
    ' 6/4/97  :  DMW  fixed sig figs (should have 0 # after decimal)
    ' 6/9/97  :  DMW  added a check for 0 here so format function wouldn't error.  Assuming no FP of 0 K

Dim FPTDI, FPTGP, FPTCR, FPTDA As Double    ' flashpt values
Dim dummyvalue As Double
Dim HLFLDI(100) As Double   ' this is also in lower, make it a global???
Dim HLFLGP(100) As Double
Dim CLFL As Double
Dim FPTI As Double
Dim DFPTDA As Double
Dim DFPTDI As Double
Dim DFPTGP As Double
Dim DFPTCR As Double
Dim FNORM As Double
Dim SNUHDI, SNUHGP As Double
Dim CFP, DA, DI, GP, CR, AIR, CAIR As Double

Dim Block5Table As Recordset
Dim Block5DBName As String
Dim string1, string2, string3, string4, string5 As String

Dim NETC(120) As Integer
Dim NERR, NCODR As Integer
Dim JF, JL, ICODEP As Integer
Dim i, J, curindex As Integer

On Error GoTo error_in_flpt
dummyvalue = 100000000000000#
FPTDI = dummyvalue
FPTGP = dummyvalue
FPTCR = dummyvalue
FPTDA = dummyvalue
CAIR = 0.512
CFP = 0#
ICODEP = 0
NERR = 0
NCODR = 0

For i = 0 To ntgp - 1
    HLFLGP(i) = 0#
    HLFLDI(i) = 0
Next i
    ' get this from data file??
HLFLDI(14) = 9.1
HLFLDI(19) = -4.38
HLFLDI(25) = 2.17
HLFLDI(35) = 2.17
HLFLDI(40) = 17.5
HLFLDI(52) = 1.38
HLFLDI(59) = -2.68
HLFLDI(61) = 9.6
HLFLDI(76) = 10.9
HLFLDI(80) = 1.3

    ' get from LFLHORG.DAT
    Set Block5Table = Block5DB.OpenRecordset("LFLHORG", dbOpenTable)
    While Not Block5Table.EOF
        curindex = Block5Table("groupindex")
        HLFLGP(curindex - 1) = Block5Table("data")
        Block5Table.MoveNext
    Wend
    Block5Table.Close
If inorganic = True Then
    Set Block5Table = Block5DB.OpenRecordset("LFLHINO", dbOpenTable)
    While Not Block5Table.EOF
        curindex = Block5Table("groupindex")
        If Block5Table("data") <> 0 Then
            HLFLGP(curindex - 1) = Block5Table("data")
        End If
        Block5Table.MoveNext
    Wend
    Block5Table.Close
End If
    
For J = 0 To ntel - 1
    NETC(J) = 0
Next J

      NETC(3) = NGPT(57)
      NETC(8) = NGPT(58)
      NETC(13) = NGPT(26) + 2 * NGPT(27) + 3 * NGPT(28)
      NETC(14) = NGPT(0) + NGPT(1) + NGPT(2) + NGPT(3) + NGPT(8) + NGPT(10) + NGPT(11) + NGPT(35) + 6 * NGPT(13) + NGPT(37) + NGPT(36) + 6 * (NGPT(14) + NGPT(15) + NGPT(16) + NGPT(17)) + NGPT(9) + NGPT(38) + NGPT(46) + 2 * NGPT(12)
      NETC(19) = NGPT(20) + 2 * NGPT(21) + 3 * NGPT(22)
      NETC(22) = NGPT(59)
      NETC(25) = NET(25)
      NETC(29) = NGPT(23) + 2 * NGPT(24) + 3 * NGPT(25)
      NETC(35) = 3 * NGPT(0) + 2 * NGPT(1) + NGPT(2) + NGPT(4) + NGPT(5) + NGPT(10) + 2 * NGPT(32) + 5 * NGPT(13) + NGPT(40) + 2 * NGPT(38) + 4 * (NGPT(14) + NGPT(15) + NGPT(16)) + NGPT(33) + NGPT(9)
      NETC(40) = NGPT(29) + 2 * NGPT(30) + 3 * NGPT(31)
      NETC(52) = NGPT(32) + NGPT(35) + NGPT(34) + NGPT(39) + NGPT(37) + NGPT(36) + NGPT(33) + NGPT(38)
      NETC(53) = NGPT(60)
      NETC(60) = NGPT(5) + NGPT(6) + NGPT(8) + 2 * (NGPT(10) + NGPT(11) + NGPT(39) + NGPT(43)) + 3 * (NGPT(44) + NGPT(49) + NGPT(46)) + 4 * (NGPT(45) + NGPT(50)) + NGPT(37) + 2 * NGPT(7) + NGPT(42) + NGPT(48) + NGPT(9) + NGPT(38) + 3 * NGPT(12)
      NETC(61) = NGPT(47) + NGPT(50) + NGPT(49) + NGPT(48)
      NETC(76) = NGPT(40) + NGPT(43) + NGPT(44) + NGPT(45) + NGPT(42) + NGPT(41)
      NETC(80) = (NGPT(51) + NGPT(52) + NGPT(53) + NGPT(54) + NGPT(55) + 2 * NGPT(56)) / 4

    ' check for errors in contribution from each group
For J = 0 To ntel - 1
    If NETC(J) <> NET(J) Then
            'Debug.Print string4; EL(J); string5; ICODEP
            NERR = NERR + 1
    End If
Next J

If NET(35) > NET(29) Or NET(35) = NET(29) Then
    HLFLDI(29) = -4.18
Else
    HLFLDI(29) = -2.55
End If

        ' JF will hold the position in QB that corresponds to BL(2,0)
        ' QB is possible values for the quality code (QC)
    For i = 0 To 20
        JF = i
        If bl(2, 0) = QB(i) Then
            GoTo continue_JF_calc
        End If
    Next i
    If bl(2, 0) <> " " Then
        NERR = NERR + 1
    End If
continue_JF_calc:

        ' JL will hold the position in QB that corresponds to BL(2,1)
        
    For i = 0 To 20
        JL = i
        If bl(2, 1) = QB(i) Then
            GoTo continue_JL_calc
        End If
    Next i
    If bl(2, 1) <> " " Then
        NERR = NERR + 1
    End If
continue_JL_calc:

If FLP > 10000# Then FLP = 1E+15
If NGPT(75) <> 0 Then FLP = 1E+15
If FML >= 100# Then FLP = 1E+15
DA = -1E+15
DI = -1E+15
GP = -1E+15
CR = -1E+15
If FML >= 100# Then DA = 1E+15
If FML >= 100# Then DI = 1E+15
If FML >= 100# Then GP = 1E+15
If FML >= 100# Then CR = 1E+15
If neqvp = 0 Then GoTo 403

For J = 1 To 4
      If vpc(J) <> 0# Then GoTo 402
Next J
GoTo 403
402 If NGPT(74) <> 0 Then FML = 100#
If FML >= 100# Then
    FPTDA = 1E+15
ElseIf FML <= 0# Then
    FPTDA = -1E+15
ElseIf JL > 10 Then
    FPTDA = -1E+15
Else
    CLFL = FML
    FPTI = 550#
    If TBP > 0# Or TBP = 0# Then FPTI = TBP
    ' for these purposes DNEQNFP will always use the NEQVP equation
    Call DNEQNFP(0.00001, 1, 100, FPTI, DFPTDA, FNORM, CLFL)
    FPTDA = DFPTDA + CFP
End If


SNUHDI = NGPT(19) * 14.07
SNUHGP = 0#
For J = 0 To ntgp - 1
    SNUHGP = SNUHGP + NGPT(J) * HLFLGP(J)
Next J
    
For J = 0 To ntel - 1
    SNUHDI = SNUHDI + CDbl(NET(J)) * HLFLDI(J)
Next J

If SNUHDI <= 1# Then FMLDI = 100#
If SNUHDI > 1# Then FMLDI = 100# / SNUHDI
If SNUHGP <= 1# Then FMLGP = 100#
If SNUHGP > 1# Then FMLGP = 100# / SNUHGP
If NGPT(74) <> 0 Then FMLDI = 100#
If NGPT(74) <> 0 Then FMLGP = 100#
If FMLDI <= 0# Then
    FPTDI = -1E+15
ElseIf FMLDI > 100# Or FMLDI = 100# Then
    FPTDI = 1E+15
Else
CLFL = FMLDI
FPTI = 550#
If TBP > 0# Or TBP = 0# Then FPTI = TBP

    ' again, this function will use the VP equation
Call DNEQNFP(0.00001, 1, 100, FPTI, DFPTDI, FNORM, CLFL)
FPTDI = DFPTDI + CFP
End If
If FMLGP <= 0# Then
    FPTGP = -1E+15
ElseIf FMLGP >= 100# Then
    FPTGP = 1E+15
Else
    CLFL = FMLGP
    FPTI = 550#
    If TBP > 0# Or TBP = 0# Then FPTI = TBP
    Call DNEQNFP(0.00001, 1, 100, FPTI, DFPTGP, FNORM, CLFL)
    FPTGP = DFPTGP + CFP
End If
    ' calculating the amount of air required to burn the compound
AIR = 0#
For J = 0 To ntel - 1
    If J = 35 Then
        If NET(J) > (NET(13) + NET(19) + NET(29) + NET(40)) Then
            AIR = AIR + CNU(J) * CDbl((NET(J) - NET(13) - NET(19) - NET(29) - NET(40)))
        End If
    ElseIf J <> 25 Then
       AIR = AIR + CNU(J) * CDbl(NET(J))
    End If
Next J
AIR = AIR / 0.21
If AIR <= 0# Then FMLCR = 100#
If AIR > 0# Then FMLCR = 100# / (1# + AIR / CAIR)
If NGPT(74) <> 0 Then FMLCR = 100#
If FMLCR >= 100# Then
    FPTCR = 1E+15
ElseIf FMLCR <= 0# Then
    FPTCR = -1E+15
Else
    CLFL = FMLCR
    FPTI = 550#
    If TBP > 0# Or TBP = 0 Then FPTI = TBP
        ' this function will use the VP equation
    Call DNEQNFP(0.00001, 1, 100, FPTI, DFPTCR, FNORM, CLFL)
    FPTCR = DFPTCR + CFP
End If
    ' this is quality code stuff, not currently being used by pearls
403   QC(0) = QB(JF)  ' fix JF to be correct index ????

    ' the following just give a little feedback for debugging purposes
    ' indicates the compound is non-combustible
If FML >= 100# Then
    QC(0) = "NC"    ' was flm ???
    'Block5frm!notelbl.Caption = "non-combustible"
End If
    ' indicates fpt is not applicable
'If FLP < 0# Then
 '   QC(0) = "NA"
 '   Block5frm!notelbl.Caption = "not applicable"
'End If
If FPTDA > dummyvalue Or FPTDA < -dummyvalue Or FPTDA = 0 Then
    InfoMethod(FP).Enabled(CNT + 1) = False
Else
    InfoMethod(FP).value(CNT + 1) = Format(FPTDA, "#.")
    InfoMethod(FP).MethodName(CNT + 1) = "LFL Data"
    InfoMethod(FP).Enabled(CNT + 1) = True
End If

If FPTGP > dummyvalue Or FPTGP < -dummyvalue Or FPTGP = 0 Then
    InfoMethod(FP).Enabled(CNT + 2) = False
Else
    InfoMethod(FP).value(CNT + 2) = Format(FPTGP, "#.")
    InfoMethod(FP).MethodName(CNT + 2) = "MTU LFL Group Contribution"
    InfoMethod(FP).Enabled(CNT + 2) = True
End If
If FPTDI > dummyvalue Or FPTDI < -dummyvalue Or FPTDI = 0 Then
    InfoMethod(FP).Enabled(CNT + 3) = False
Else
    InfoMethod(FP).value(CNT + 3) = Format(FPTDI, "#.")
    InfoMethod(FP).MethodName(CNT + 3) = "Penn State Group Contribution"
    InfoMethod(FP).Enabled(CNT + 3) = True
End If
If FPTCR > dummyvalue Or FPTCR < -dummyvalue Or FPTCR = 0 Then
    InfoMethod(FP).Enabled(CNT + 4) = False
Else
    InfoMethod(FP).value(CNT + 4) = Format(FPTCR, "#.")
    InfoMethod(FP).MethodName(CNT + 4) = "MTU LFL Combustion Reaction"
    InfoMethod(FP).Enabled(CNT + 4) = True
End If
If FLP > dummyvalue Or FLP < -dummyvalue Then
    InfoMethod(FP).Enabled(CNT) = False
Else
    InfoMethod(FP).value(CNT) = Format(FLP, "#.")
    InfoMethod(FP).MethodName(CNT) = "MTU Fire & Explosion Data"
    InfoMethod(FP).Enabled(CNT) = True
End If
Exit Sub
    ' now, the data we want in order of preference:
    '   1. FPTDA -> based on LFL Data
    '   2. FPTGP -> based on MTU LFL from Group Contributions
    '   3. FPTDI -> based on Penn State U FL
    '   4. FPTCR -> based on MTU LFL from Combustion Reaction
    '   FLP = flashpoint data (from file)
error_in_flpt:
    
    
1000 End Sub


Private Sub get_FPFN(X As Double, F As Double, N As Integer, CLFL As Double)
   
Dim VP As Double
On Error GoTo FPFN_error
Call eqnsub(X, VP)
F = VP - CLFL * 1013.25
Exit Sub
FPFN_error:
    
End Sub


Private Sub eqnsub(T As Double, value As Double)
    ' 3/23/97 : DMW  fixed array indexes so arrays start at 0
    
Dim C(5) As Double
Dim N As Integer
Dim i, J As Integer
    ' the vapor pressure coefficients (change this to use them directly)
For J = 0 To 4
    C(J) = vpc(J)
Next J
    ' N is the vapor pressure equation here
N = neqvp
value = 0#
If N = 100 Then
    For i = 0 To 4
        value = value + C(i) * T ^ (i - 1)
    Next i
ElseIf N = 101 Then
    value = Exp#(C(0) + C(1) / T + C(2) * Log#(T) + C(3) * T ^ C(4))
ElseIf N = 102 Then
    value = C(0) * T ^ C(1) / (1# + C(2) / T + C(3) / T ^ 2)
ElseIf N = 103 Then
    value = C(0) + C(1) * Exp#(-C(2) / T ^ C(3))
ElseIf N = 104 Then
    value = C(0) + C(1) / T + C(2) / T ^ 3 + C(3) / T ^ 8 + C(4) / T ^ 9
ElseIf N = 105 Then
    value = C(0) / C(1) ^ (1# + (1# - T / C(2)) ^ C(3))
ElseIf N = 106 Then
    value = C(0) * (1# - T) ^ (C(1) + C(2) * T + C(3) * T ^ 2 + C(4) * T ^ 3)
ElseIf N = 107 Then
    value = C(0) + C(1) * ((C(2) / T) / Sin#(C(2) / T)) ^ 2 + C(3) * ((C(4) / T) / Cos#(C(4) / T)) ^ 2
Else
    'debug.Print " INCORRECT EQUATION NUMBER IN EQNSUB"
End If

End Sub


Private Sub DNEQNFP(TOL As Double, NEQN As Integer, MITR As Double, FPTI As Double, DFPT As Double, FNORM As Double, CLFL1 As Double)
Dim fx As Double
Dim fx1 As Double
Dim fx2 As Double
Dim NITR As Integer
Dim DFDT As Double
Dim FPT1 As Double
On Error GoTo no_convergence
NITR = 0
FPT1 = FPTI
Call get_FPFN(FPT1, fx, NEQN, CLFL1)
1 Call get_FPFN(FPT1 * (1# - TOL), fx1, NEQN, CLFL1)
Call get_FPFN(FPT1 * (1# + TOL), fx2, NEQN, CLFL1)
If TOL = 0 Then
    GoTo no_convergence
End If
If FPT1 = 0 Then
    GoTo no_convergence
End If

DFDT = (fx2 - fx1) / (2# * TOL * FPT1)
If DFDT = 0 Then
    GoTo no_convergence
End If
DFPT = FPT1 - fx / DFDT
Call get_FPFN(DFPT, fx, NEQN, CLFL1)
FNORM = fx ^ 2
If DFPT = 0 Then
    'MsgBox ("2nd DFPT divide by 0")
End If
If Abs((DFPT - FPT1) / DFPT) < TOL Then Exit Sub
NITR = NITR + 1
FPT1 = DFPT
If NITR < MITR Then
    GoTo 1
End If
'Debug.Print "NO CONVERGENCE FOR FLASHPOINT"
Exit Sub
no_convergence:
   ' MsgBox ("no convergence")
    Exit Sub
End Sub



Private Sub EQNSUBL(T As Double, thevalue As Double)
    
        ' 3/23/97 : DMW  fixed array indexes so arrays start at 0
        
        Dim i As Integer
        Dim N As Integer
        
        N = neqvp
    If N = 100 Then
        thevalue = 0#
        For i = 0 To 4
          thevalue = thevalue + vpc(i) * T ^ (i - 1)
        Next i
    ElseIf N = 101 Then
       thevalue = Exp(vpc(0) + vpc(1) / T + vpc(2) * Log(T) + vpc(3) * T ^ vpc(4))
    ElseIf N = 102 Then
       thevalue = vpc(0) * T ^ vpc(1) / (1# + vpc(2) / T + vpc(3) / T ^ 2)
    ElseIf N = 103 Then
       thevalue = vpc(0) + vpc(1) * Exp(-vpc(2) / T ^ vpc(3))
    ElseIf N = 104 Then
       thevalue = vpc(0) + vpc(1) / T + vpc(2) / T ^ 3 + vpc(3) / T ^ 8 + vpc(4) / T ^ 9
    ElseIf N = 105 Then
       thevalue = vpc(0) / vpc(1) ^ (1# + (1# - T / vpc(2)) ^ vpc(3))
    ElseIf N = 106 Then
       thevalue = vpc(0) * (1# - T) ^ (vpc(1) + vpc(2) * T + vpc(3) * T ^ 2 + vpc(4) * T ^ 3)
    ElseIf N = 107 Then
       thevalue = vpc(0) + vpc(1) * ((vpc(2) / T) / Sin#(vpc(2) / T)) ^ 2 + vpc(3) * ((vpc(4) / T) / Cos#(vpc(4) / T)) ^ 2
    Else
       'debug.print " INVPCORREVPCT EQUATION NUMBER IN EQNSUB"
    End If
    
     

End Sub

Public Sub calc_AIT(CNT As Integer)
    
    ' calc_AIT:  a function that calculates the ait and puts the
    '           values in the text boxes on the main form
    '           arrays are indexed from 0 instead of 1 (as in the fortran code)
        ' 5/30/97 :  DMW  CNT is the position in the method array to put the first result
        ' 6/4/97  :  DMW  fixed sig figs (should have 0 # after decimal)
    Dim AITGP, AITG1 As Double                  ' auto-ig temp values
    Dim Block5Table As Recordset
    Dim SNUHGP As Double
    Dim SNUHG1 As Double
    Dim SNUHA0 As Double
    Dim ICODEP As Integer
    Dim A0G As Double
    Dim A0G1 As Double
    Dim EAIT(2)  As Double
    Dim dummyvalue As Double
    Dim NETC(120) As Integer
    
        ' values from tables
    Dim HAITGP(100) As Double
    'Dim NNINCL(10) As Double
    Dim HAITA0(100) As Double
    Dim HAITG1(100) As Double
        ' left out EMAX, EMIN, NERRT and ERR2T
    Dim NAIT(10) As Double
    
    Dim IGCOD(16) As Double ' need this?
    Dim JA As Integer
    Dim i, J As Integer
    Dim curindex As Integer
    Dim NERR As Integer
    
    'On Error GoTo error_in_ait
    dummyvalue = -10000000000000#
    AITGP = dummyvalue
    AITG1 = dummyvalue
    
    JA = 0
    NERR = 0
    A0G = 1500
    A0G1 = 0#
    ICODEP = 0
    
        ' first initialize these, then get original group cont values from table
    For i = 0 To ntgp - 1
        HAITA0(i) = 0
        HAITGP(i) = 0
        HAITG1(i) = 0
    Next i
        
        Set Block5Table = Block5DB.OpenRecordset("AITHORG", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            HAITGP(curindex - 1) = Block5Table("data")
            Block5Table.MoveNext
        Wend
        Block5Table.Close
        Set Block5Table = Block5DB.OpenRecordset("AITAORG", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            HAITA0(curindex - 1) = Block5Table("data")
            Block5Table.MoveNext
        Wend
        Block5Table.Close
        Set Block5Table = Block5DB.OpenRecordset("AITBORG", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            HAITG1(curindex - 1) = Block5Table("data")
            Block5Table.MoveNext
        Wend
        Block5Table.Close
   If inorganic = True Then
        Set Block5Table = Block5DB.OpenRecordset("AITHINO", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            If Block5Table("data") <> 0 Then
                HAITGP(curindex - 1) = Block5Table("data")
            End If
            Block5Table.MoveNext
        Wend
        Block5Table.Close
        Set Block5Table = Block5DB.OpenRecordset("AITAINO", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            If Block5Table("data") <> 0 Then
                HAITA0(curindex - 1) = Block5Table("data")
            End If
            Block5Table.MoveNext
        Wend
        Block5Table.Close
        Set Block5Table = Block5DB.OpenRecordset("AITBINO", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            If Block5Table("data") <> 0 Then
                HAITG1(curindex - 1) = Block5Table("data")
            End If
            Block5Table.MoveNext
        Wend
        Block5Table.Close
   End If
    
        ' making sure the dippr code is a valid one
    If ICODE < 0 Then
        GoTo end_of_calc
    End If
      
    For i = 0 To ntel - 1
        NETC(i) = 0
    Next i
    
    NETC(3) = NGPT(57)
    NETC(8) = NGPT(58)
    NETC(13) = NGPT(26) + 2 * NGPT(27) + 3 * NGPT(28)
    NETC(14) = NGPT(0) + NGPT(1) + NGPT(2) + NGPT(3) + NGPT(8) + NGPT(10) + NGPT(11) + NGPT(35) + 6 * NGPT(13) + NGPT(37) + NGPT(36) + 6 * (NGPT(14) + NGPT(15) + NGPT(16) + NGPT(17)) + NGPT(9) + NGPT(46) + 2 * NGPT(12)
    NETC(19) = NGPT(20) + 2 * NGPT(21) + 3 * NGPT(22)
    NETC(22) = NGPT(59)
    NETC(25) = NET(25)
    NETC(29) = NGPT(23) + 2 * NGPT(24) + 3 * NGPT(25)
    NETC(30) = NGPT(65)
    NETC(35) = 3 * NGPT(0) + 2 * NGPT(1) + NGPT(2) + NGPT(4) + NGPT(5) + NGPT(10) + 2 * NGPT(32) + 5 * NGPT(13) + NGPT(40) + 2 * NGPT(38) + 4 * (NGPT(14) + NGPT(15) + NGPT(16)) + NGPT(33) + NGPT(9)
    NETC(40) = NGPT(29) + 2 * NGPT(30) + 3 * NGPT(31)
    NETC(52) = NGPT(32) + NGPT(35) + NGPT(34) + NGPT(39) + NGPT(37) + NGPT(36) + NGPT(33) + 2 * NGPT(38)
    NETC(53) = NGPT(60)
    NETC(57) = NGPT(66)
    NETC(59) = NGPT(5) + NGPT(6) + NGPT(8) + 2 * (NGPT(10) + NGPT(11) + NGPT(39) + NGPT(43)) + 3 * (NGPT(44) + NGPT(49) + NGPT(46)) + 4 * (NGPT(45) + NGPT(50)) + NGPT(37) + 2 * NGPT(7) + NGPT(42) + NGPT(48) + NGPT(9) + 3 * NGPT(12)
    NETC(61) = NGPT(47) + NGPT(50) + NGPT(49) + NGPT(48)
    NETC(76) = NGPT(40) + NGPT(43) + NGPT(44) + NGPT(45) + NGPT(42) + NGPT(41)
    NETC(80) = (NGPT(51) + NGPT(52) + NGPT(53) + NGPT(54) + NGPT(55) + 2 * NGPT(56)) / 4
    NETC(98) = NGPT(67)

    For i = 0 To ntel - 1
        If NETC(i) = NET(i) Then
            GoTo next_i_netc
        End If
        ' Print Debug "Error in BALANCE near ???"
        NERR = NERR + 1
        
next_i_netc:
    Next i
    
          ' JA will hold the position in QB that corresponds to BL
    For i = 0 To 20
        JA = i
        If bl(2, 3) = QB(i) Then
            GoTo continue_605
        End If
    Next i
    If bl(2, 3) <> " " Then
        NERR = NERR + 1
    End If
continue_605:


        SNUHGP = 0#
        SNUHG1 = 0#
        SNUHA0 = 0#
        For J = 0 To ntgp - 1
            SNUHGP = SNUHGP + NGPT(J) * HAITGP(J)
            SNUHG1 = SNUHG1 + NGPT(J) * HAITG1(J)
            SNUHA0 = SNUHA0 + NGPT(J) * HAITA0(J)
        Next J
        
            ' 915 is the dippr code for AIR
        If ICODE = 915 Then
            SNUHGP = SNUHGP / 100#
            SNUHG1 = SNUHG1 / 100#
            SNUHA0 = SNUHA0 / 100#
        End If
5 If SNUHGP <= 0# Then
            AITGP = 0.000000000000001
        ElseIf SNUHGP > 0# Then
            AITGP = A0G * (1# + SNUHA0) / Log#(SNUHGP)
        End If
        EAIT(0) = AITGP - AITemp
        AITG1 = A0G1 + SNUHG1
        If AITG1 <= 0# Then
            AITG1 = -0.000000000000001
        End If
        EAIT(1) = AITG1 - AITemp
        
end_of_calc:
        
        If AITGP < dummyvalue Or AITGP > -dummyvalue Or AITGP = 0 Then
            InfoMethod(AIT).Enabled(CNT + 1) = False
        Else
            InfoMethod(AIT).value(CNT + 1) = Format(AITGP, "#.")
            InfoMethod(AIT).MethodName(CNT + 1) = "MTU Logarithmic Method"
            InfoMethod(AIT).Enabled(CNT + 1) = True
        End If
        If AITG1 < dummyvalue Or AITG1 > -dummyvalue Or AITG1 = 0 Then
            InfoMethod(AIT).Enabled(CNT + 2) = False
        Else
            If IsNumeric(Format(AITG1, "#.")) Then 'if valid number after format...
                InfoMethod(AIT).value(CNT + 2) = Format(AITG1, "#.")
                InfoMethod(AIT).MethodName(CNT + 2) = "MTU Linear Method"
                InfoMethod(AIT).Enabled(CNT + 2) = True
            End If
        End If
        If AITemp < dummyvalue Or AITemp = 0 Then
            InfoMethod(AIT).Enabled(CNT) = False
        Else
            InfoMethod(AIT).value(CNT) = Format(AITemp, "#.")
            InfoMethod(AIT).MethodName(CNT) = "MTU Fire & Explosion Data"
            InfoMethod(AIT).Enabled(CNT) = True
        End If
            ' the values of interest here in order of preference are:
            ' but we're putting the Fire & Explosion data file data first
            '   1.  AITGP -> Estimated AIT by MTU logarithmic group method
            '   2.  AITG1 -> Estimated AIT by MTU linear group method
            '   AITemp -> Actual data
            ' significant figures:  no # after
            Exit Sub
error_in_ait:
            
End Sub






Public Sub calc_lower(CNT As Integer)
    ' calc_lower:  a function that calculates the lfl and puts the
    '           values in the text boxes on the main form
    '           arrays are indexed from 0 instead of 1 (as in the fortran code)
    
    ' 3/23/97 :  DMW  changed array indexes to start from 0
    ' 5/30/97 :  DMW  CNT is the position in the method array to put the first result
    ' 6/4/97  :  DMW  fixed sig figs (should have 2 # after decimal)
Dim NERR, JF As Integer
Dim CFP, ICODEP, NCODR, JL, SNUHDI, SNUHGP, AIR, CAIR As Double ' ?? not sure of type
Dim VPFP As Double
Dim i, J As Integer
Dim HLFLDI(100) As Double
Dim HLFLGP(100) As Double
Dim dummyvalue As Double
    ' following ok
Dim NETC(120) As Integer
Dim string1, string2, string3 As String
Dim Block5Table As Recordset
Dim curindex As Integer
'On Error GoTo error_in_lower
    dummyvalue = 100000000000000#
    FMLDI = dummyvalue
    FMLGP = dummyvalue
    FMLCR = dummyvalue
    FMLFP = dummyvalue
    CAIR = 0.512
    CFP = 3#  ' was 0, set to 3 just for LFL
    ICODEP = 0
    NERR = 0
    NCODR = 0

For i = 0 To ntgp - 1
      HLFLGP(i) = 0#
      HLFLDI(i) = 0#
Next i
    ' Penn State values
      HLFLDI(14) = 9.1
      HLFLDI(19) = -4.38
      HLFLDI(25) = 2.17
      HLFLDI(35) = 2.17
      HLFLDI(40) = 17.5
      HLFLDI(52) = 1.38
      HLFLDI(59) = -2.68
      HLFLDI(61) = 9.6
      HLFLDI(76) = 10.9
      HLFLDI(80) = 1.3
      
    ' from database - MTU values
        Set Block5Table = Block5DB.OpenRecordset("LFLHORG", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            HLFLGP(curindex - 1) = Block5Table("data")
            Block5Table.MoveNext
        Wend
        Block5Table.Close
    If inorganic = True Then
        Set Block5Table = Block5DB.OpenRecordset("LFLHINO", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            ' subtract one to account for indexing arrays from 0
            If Block5Table("data") <> 0 Then
                HLFLGP(curindex - 1) = Block5Table("data")
            End If
            Block5Table.MoveNext
        Wend
        Block5Table.Close
    End If
       
      ' I think this is wrong because BL never has more than 2 chars in it ??
'For I = 0 To 9
'For J = 0 To 15
'        If I = 2 Then
'            GoTo next_j_i_BL
'        End If
'        If I > 2 And J = 0 Then
'            GoTo next_j_i_BL
'        End If
'        If bl(I, J) <> " " And bl(I, J) <> "  " Then
            'Debug.Print "ERROR IN SPACING ("; I; ") CODE NO.    "; ICODEP
'            NERR = NERR + 1
'        End If

'next_j_i_BL:
'    Next J: Next I


For J = 0 To ntel - 1
       NETC(J) = 0
Next J

      NETC(3) = NGPT(57)
      NETC(8) = NGPT(58)
      NETC(13) = NGPT(26) + 2 * NGPT(27) + 3 * NGPT(28)
      NETC(14) = NGPT(0) + NGPT(1) + NGPT(2) + NGPT(3) + NGPT(8) + NGPT(10) + NGPT(11) + NGPT(35) + 6 * NGPT(13) + NGPT(37) + NGPT(36) + 6 * (NGPT(14) + NGPT(15) + NGPT(16) + NGPT(17)) + NGPT(9) + NGPT(38) + NGPT(46) + 2 * NGPT(12)
      NETC(19) = NGPT(20) + 2 * NGPT(21) + 3 * NGPT(22)
      NETC(22) = NGPT(59)
      NETC(25) = NET(25)
      NETC(29) = NGPT(23) + 2 * NGPT(24) + 3 * NGPT(25)
      NETC(35) = 3 * NGPT(0) + 2 * NGPT(1) + NGPT(2) + NGPT(4) + NGPT(5) + NGPT(10) + 2 * NGPT(32) + 5 * NGPT(13) + NGPT(40) + 2 * NGPT(38) + 4 * (NGPT(14) + NGPT(15) + NGPT(16)) + NGPT(33) + NGPT(9)
      NETC(40) = NGPT(29) + 2 * NGPT(30) + 3 * NGPT(31)
      NETC(52) = NGPT(32) + NGPT(35) + NGPT(34) + NGPT(39) + NGPT(37) + NGPT(36) + NGPT(33) + NGPT(38)
      NETC(53) = NGPT(60)
      NETC(59) = NGPT(5) + NGPT(6) + NGPT(8) + 2 * (NGPT(10) + NGPT(11) + NGPT(39) + NGPT(43)) + 3 * (NGPT(44) + NGPT(49) + NGPT(46)) + 4 * (NGPT(45) + NGPT(50)) + NGPT(37) + 2 * NGPT(7) + NGPT(42) + NGPT(48) + NGPT(9) + NGPT(38) + 3 * NGPT(12)
      NETC(61) = NGPT(47) + NGPT(50) + NGPT(49) + NGPT(48)
      NETC(76) = NGPT(40) + NGPT(43) + NGPT(44) + NGPT(45) + NGPT(42) + NGPT(41)
      NETC(80) = (NGPT(51) + NGPT(52) + NGPT(53) + NGPT(54) + NGPT(55) + 2 * NGPT(56)) / 4
      
    For J = 0 To ntel - 1
        If NETC(J) <> NET(J) Then
            'Debug.Print " ERROR IN "; EL(J); " BALANCE NEAR CODE NO. "; ICODEP
            NERR = NERR + 1
        End If
    Next J
    
    If NET(35) > NET(29) Or NET(35) = NET(29) Then
       HLFLDI(29) = -4.18
    Else
       HLFLDI(29) = -2.55
    End If
      
        ' JL will hold the position in QB that corresponds to BL
    For i = 0 To 20
        JL = i
        If bl(2, 1) = QB(i) Then
            GoTo continue_605
        End If
    Next i
    If Trim(bl(2, 1)) <> "" Then
        NERR = NERR + 1
    End If
continue_605:
      
    SNUHDI = NGPT(19) * 14.07
    SNUHGP = 0#
    For J = 0 To ntgp - 1
        SNUHGP = SNUHGP + NGPT(J) * HLFLGP(J)
    Next J

    For J = 0 To ntel - 1
        SNUHDI = SNUHDI + NET(J) * HLFLDI(J)
    Next J
        
    If SNUHDI < 1# Or SNUHDI = 1# Then FMLDI = 1E+15
    If SNUHDI > 1# Then FMLDI = 100# / SNUHDI
    If SNUHGP < 1# Or SNUHGP = 1# Then FMLGP = 1E+15
    If SNUHGP > 1# Then FMLGP = 100# / SNUHGP
    If NGPT(74) <> 0 Then FMLDI = 1E+15
    If NGPT(74) <> 0 Then FMLGP = 1E+15
    If FML > 100# Or FML = 100 Then FMLDI = 1E+15
    If FML > 100# Or FML = 100 Then FMLGP = 1E+15
    AIR = 0#
    For J = 0 To ntel - 1
        If J = 35 Then
            If NET(J) > (NET(13) + NET(19) + NET(29) + NET(40)) Then
                AIR = AIR + CNU(J) * (NET(J) - NET(13) - NET(19) - NET(29) - NET(40))
            End If
        ElseIf J <> 25 Then
            AIR = AIR + CNU(J) * NET(J)
        End If
   Next J
    AIR = AIR / 0.21
    If AIR < 0# Or AIR = 0 Then FMLCR = 1E+15
    If AIR > 0# Then FMLCR = 100# / (1# + AIR / CAIR)
    If NGPT(74) <> 0 Then FMLCR = 1E+15
    If FML > 100# Or FML = 100 Then FMLCR = 1E+15
    FMLFP = -1E+15
    For i = 0 To 20
        JF = i
        If bl(2, 0) = QB(i) Then
            GoTo 606
        End If
606
    Next i
    If bl(2, 0) <> " " Then NERR = NERR + 1
    If FLP < 0# Then GoTo dont_calc_vp_method
    If FLP = 0# Then
        FMLFP = 0#
        GoTo dont_calc_vp_method
    End If
    If FLP > 10000# Then
        FMLFP = 1E+15
        GoTo dont_calc_vp_method
    End If
    If neqvp <> 0 Then
        Call EQNSUBL(FLP - CFP, VPFP)
        FMLFP = VPFP * 100# / 101325#
        If FMLFP > 100# Or FMLFP = 100 Then FMLFP = 1E+15
        If FML > 100# Or FML = 100 Then FMLFP = 1E+15
    End If
dont_calc_vp_method:
      QC(0) = QB(JL)
      If FML > 100# Or FML = 100 Then QC(0) = "NC"
      If FML < 0# Then QC(0) = "NA"
      

        ' the values in order of preference
        ' 1.  FMLGP -> MTU LFL Group Contribution data
        ' 2.  FMLDI -> Penn State U. Data
        ' 3.  FMLCR -> MTU for Combustion Reaction
        ' 4.  FMLFP -> Flashpoint Method
        ' fml - number from data file

    If FMLGP > dummyvalue Or FMLGP < -dummyvalue Then
        InfoMethod(LFL).Enabled(CNT + 1) = False
    Else
        InfoMethod(LFL).value(CNT + 1) = Format(FMLGP, "#.##")
        InfoMethod(LFL).MethodName(CNT + 1) = "MTU Group Contribution"
        InfoMethod(LFL).Enabled(CNT + 1) = True
    End If
    If FMLDI > dummyvalue Or FMLDI < -dummyvalue Then
        InfoMethod(LFL).Enabled(CNT + 2) = False
    Else
        InfoMethod(LFL).value(CNT + 2) = Format(FMLDI, "#.##")
        InfoMethod(LFL).MethodName(CNT + 2) = "Penn State Group Contribution"
        InfoMethod(LFL).Enabled(CNT + 2) = True
    End If
    If FMLCR > dummyvalue Or FMLCR < -dummyvalue Then
        InfoMethod(LFL).Enabled(CNT + 3) = False
    Else
        InfoMethod(LFL).value(CNT + 3) = Format(FMLCR, "#.##")
        InfoMethod(LFL).MethodName(CNT + 3) = "MTU Combustion Reaction"
        InfoMethod(LFL).Enabled(CNT + 3) = True
    End If
    If FMLFP > dummyvalue Or FMLFP < -dummyvalue Then
        InfoMethod(LFL).Enabled(CNT + 4) = False
    Else
        InfoMethod(LFL).value(CNT + 4) = Format(FMLFP, "#.##")
        InfoMethod(LFL).MethodName(CNT + 4) = "MTU FlashPoint Method"
        InfoMethod(LFL).Enabled(CNT + 4) = True
    End If
    If FML > dummyvalue Or FML < -dummyvalue Then
        InfoMethod(LFL).Enabled(CNT) = False
    Else
        InfoMethod(LFL).value(CNT) = Format(FML, "#.##")
        InfoMethod(LFL).MethodName(CNT) = "MTU Fire & Explosion Data"
        InfoMethod(LFL).Enabled(CNT) = True
    End If
        ' the values in order of preference
        ' 1.  FMLGP -> MTU LFL Group Contribution data
        ' 2.  FMLDI -> Penn State U. Data
        ' 3.  FMLCR -> MTU for Combustion Reaction
        ' 4.  FMLFP -> Flashpoint Method
        ' fml - number from data file

Exit Sub

error_in_lower:

End Sub

Public Sub calc_upper(CNT As Integer)
    ' calc_upper:  a function that calculates the ufl and puts the
    '           values in the text boxes on the main form
    '           arrays are indexed from 0 instead of 1 (as in the fortran code)
    
    ' 3/23/97  fixed array indexes so arrays start at 0
    ' 5/30/97 :  DMW  CNT is the position in the method array to put the first result
    ' 6/4/97  :  DMW  fixed sig figs (should have 2 # after decimal)
Dim FMUDI, FMUGP, FMUCR As Double           ' upper fl limit values
Dim Block5Table As Recordset
Dim NETC(120) As Integer
Dim HUFLDI(100) As Double
Dim HUFLGP(100) As Double
Dim NGPDI, SNUHDI, SNUHGP, AIR, CAIR As Double  ' ??? not sure of type
Dim dummyvalue As Double
Dim curindex As Integer
Dim i, J As Integer

Dim JU, ICODEP, NERR As Integer
'On Error GoTo error_in_upper
    dummyvalue = 100000000000000#
    FMUDI = dummyvalue
    FMUGP = dummyvalue
    FMUCR = dummyvalue
    CAIR = 3.8
    ICODEP = 0
    NERR = 0
 
        ' initialize all these to 0
For i = 0 To ntgp - 1
      HUFLGP(i) = 0#
      HUFLDI(i) = 0#
201 Next i

      HUFLDI(0) = -0.9307
      HUFLDI(1) = -0.5225
      HUFLDI(4) = -0.5625
      HUFLDI(6) = 1.458
      HUFLDI(7) = 2# * HUFLDI(6)
      HUFLDI(9) = HUFLDI(4) + HUFLDI(8)
      HUFLDI(12) = HUFLDI(8) + HUFLDI(10)
      HUFLDI(13) = HUFLDI(17) + 5# * HUFLDI(4)
      HUFLDI(14) = HUFLDI(17) + 4# * HUFLDI(4)
      HUFLDI(15) = HUFLDI(17) + 4# * HUFLDI(4)
      HUFLDI(16) = HUFLDI(17) + 4# * HUFLDI(4)
      HUFLDI(18) = 1.118
      HUFLDI(19) = 4.275
      HUFLDI(20) = -0.8153
      HUFLDI(21) = 1.311
      HUFLDI(22) = -2.011
      HUFLDI(27) = HUFLDI(26)
      HUFLDI(28) = HUFLDI(26)
      HUFLDI(30) = HUFLDI(29)
      HUFLDI(31) = HUFLDI(29)
      HUFLDI(33) = HUFLDI(34) + HUFLDI(4)
      HUFLDI(36) = HUFLDI(3) + HUFLDI(34)
      HUFLDI(37) = HUFLDI(8) + HUFLDI(34)
      HUFLDI(38) = HUFLDI(8) + HUFLDI(32)
      HUFLDI(40) = HUFLDI(41) + HUFLDI(4)
      HUFLDI(42) = HUFLDI(41) + HUFLDI(6)
      HUFLDI(43) = HUFLDI(41) + 2# * HUFLDI(6)
      HUFLDI(44) = HUFLDI(41) + 3# * HUFLDI(6)
      HUFLDI(45) = HUFLDI(41) + 4# * HUFLDI(6)
      HUFLDI(46) = HUFLDI(3) + 3# * HUFLDI(6)
      HUFLDI(48) = HUFLDI(47) + HUFLDI(6)
      HUFLDI(49) = HUFLDI(47) + 3# * HUFLDI(6)
      HUFLDI(50) = HUFLDI(47) + 4# * HUFLDI(6)
      HUFLDI(70) = HUFLDI(1) - 2# * HUFLDI(4)
      HUFLDI(71) = HUFLDI(2) - HUFLDI(4)
      HUFLDI(72) = HUFLDI(3)
      
        ' the mtu organic data - get from UFLHORG.DAT
        
        Set Block5Table = Block5DB.OpenRecordset("UFLHORG", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            HUFLGP(curindex - 1) = Block5Table("data")
            Block5Table.MoveNext
        Wend
        Block5Table.Close
    If inorganic = True Then
        
        Set Block5Table = Block5DB.OpenRecordset("UFLHINO", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            If Block5Table("data") <> 0 Then
                HUFLGP(curindex - 1) = Block5Table("data")
            End If
            Block5Table.MoveNext
        Wend
        Block5Table.Close
    End If
    For J = 0 To ntel - 1
        NETC(J) = 0
    Next J

      NETC(3) = NGPT(57)
      NETC(8) = NGPT(58)
      NETC(13) = NGPT(26) + 2 * NGPT(27) + 3 * NGPT(28)
      NETC(14) = NGPT(0) + NGPT(1) + NGPT(2) + NGPT(3) + NGPT(8) + NGPT(10) + NGPT(11) + NGPT(35) + 6# * NGPT(13) + NGPT(37) + NGPT(36) + 6 * (NGPT(14) + NGPT(15) + NGPT(16) + NGPT(17)) + NGPT(9) + NGPT(38) + NGPT(46) + 2 * NGPT(12) + 2 * (NGPT(18) + NGPT(19)) + (NGPT(70) + NGPT(71) + NGPT(72) + NGPT(73))
      NETC(19) = NGPT(20) + 2 * NGPT(21) + 3 * NGPT(22)
      NETC(22) = NGPT(59)
      NETC(25) = NET(25)
      NETC(29) = NGPT(23) + 2 * NGPT(24) + 3 * NGPT(25)
      NETC(35) = 3 * NGPT(0) + 2 * NGPT(1) + NGPT(2) + NGPT(4) + NGPT(5) + NGPT(10) + 2 * NGPT(32) + 5 * NGPT(13) + NGPT(40) + 2 * NGPT(38) + 4 * (NGPT(14) + NGPT(15) + NGPT(16)) + NGPT(33) + NGPT(9)
      NETC(40) = NGPT(29) + 2 * NGPT(30) + 3 * NGPT(31)
      NETC(52) = NGPT(32) + NGPT(35) + NGPT(34) + NGPT(39) + NGPT(37) + NGPT(36) + NGPT(33) + NGPT(38)
      NETC(53) = NGPT(60)
      NETC(59) = NGPT(5) + NGPT(6) + NGPT(8) + 2 * (NGPT(10) + NGPT(11) + NGPT(39) + NGPT(43)) + 3 * (NGPT(44) + NGPT(49) + NGPT(46)) + 4 * (NGPT(45) + NGPT(50)) + NGPT(37) + 2 * NGPT(7) + NGPT(42) + NGPT(48) + NGPT(9) + NGPT(38) + 3 * NGPT(12)
      NETC(61) = NGPT(47) + NGPT(50) + NGPT(49) + NGPT(48)
      NETC(76) = NGPT(40) + NGPT(43) + NGPT(44) + NGPT(45) + NGPT(42) + NGPT(41)
      NETC(80) = (NGPT(51) + NGPT(52) + NGPT(53) + NGPT(54) + NGPT(55) + 2 * NGPT(56)) / 4
      
    For J = 0 To ntel - 1
        If NETC(J) <> NET(J) Then
            'debug.print " ERROR IN ";EL(j);" BALANCE NEAR CODE NO. ";ICODEP
            NERR = NERR + 1
        End If
    Next J
    
        ' JU will hold the position in QB that corresponds to BL
    For i = 0 To 20
        JU = i
        If bl(2, 2) = QB(i) Then
            GoTo continue_605
        End If
    Next i
    If bl(2, 2) <> " " Then
        NERR = NERR + 1
    End If
continue_605:
    
    If FMU > 100# Then FMU = 0#
    If FMU > 100# Then FMU = 0#
    NGPDI = 0#
    SNUHDI = 0#
    SNUHGP = 0#
    ' check MTU values depend on SNUHGP
    For J = 0 To ntgp - 1
        SNUHDI = SNUHDI + NGPT(J) * HUFLDI(J)
        If J <= 50 Then NGPDI = NGPDI + NGPT(J)
        SNUHGP = SNUHGP + NGPT(J) * HUFLGP(J)
    Next J
    NGPDI = NGPDI + 4# * (NGPT(14) + NGPT(15) + NGPT(16)) + NGPT(9) + NGPT(7) + NGPT(37) + NGPT(36) + NGPT(33) + NGPT(38) + NGPT(12) - NGPT(41) + 5# * NGPT(13) - NGPT(70) + NGPT(72) - NGPT(47) + NGPT(43) + 2# * NGPT(44) + 3# * NGPT(45) + 3# * NGPT(50) + 2# * NGPT(49) + 2# * NGPT(46)
    If NGPDI > 0# Then SNUHDI = SNUHDI / NGPDI
    SNUHDI = SNUHDI + 3.817 + (-0.2627 + 0.0102 * CDbl(NET(14))) * CDbl(NET(14))
    FMUDI = Exp#(SNUHDI)
      
      If FMUDI > 100# Then FMUDI = 100#
      If SNUHGP <= 1# Then FMUGP = 100#
      If SNUHGP > 1# Then FMUGP = 100# / SNUHGP
      If NGPT(74) <> 0 Then FMU = 1E+15
      If NGPT(74) <> 0 Then FMUDI = 1E+15
      If NGPT(74) <> 0 Then FMUGP = 1E+15
      
      AIR = 0#
    For J = 0 To ntel - 1
        If J = 35 Then
        If NET(J) > (NET(13) + NET(19) + NET(29) + NET(40)) Then
            AIR = AIR + CNU(J) * CDbl((NET(J) - NET(13) - NET(19) - NET(29) - NET(40)))
        End If
        ElseIf J <> 25 Then
            AIR = AIR + CNU(J) * NET(J)
        End If
240   Next J
      AIR = AIR / 0.21
      If AIR <= 0# Then FMUCR = 0#
      If AIR > 0# Then FMUCR = 100# / (1# + AIR / CAIR)
      If NGPT(74) <> 0 Then FMUCR = 1E+15
      If FMU < 0# Then FMU = -1E+15
      If FMU = 0# Then FMU = 1E+15
      If FMUDI = 0# Then FMUDI = 1E+15
      If FMUGP = 0# Then FMUGP = 1E+15
      If FMUCR = 0# Then FMUCR = 1E+15
      QC(0) = QB(JU)
      If FMU < 0# Then QC(0) = "NA"
      If FMU > 100# Then QC(0) = "NC"


    ' DATA: (in order)
    '   1.  FMUGP = MTU Value
    '   2.  FMUCR = MTU Method using Combustion Reaction Method
    '   3.  FMUDI = Penn State method
    '   fmu - value from data file
       
    If FMUGP > dummyvalue Or FMUGP < -dummyvalue Then
        InfoMethod(UFL).Enabled(CNT + 1) = False
    Else
        InfoMethod(UFL).value(CNT + 1) = Format(FMUGP, "#.##")
        InfoMethod(UFL).MethodName(CNT + 1) = "MTU Group Contribution"
        InfoMethod(UFL).Enabled(CNT + 1) = True
    End If
    If FMUCR > dummyvalue Or FMUCR < -dummyvalue Then
        InfoMethod(UFL).Enabled(CNT + 2) = False
    Else
        InfoMethod(UFL).value(CNT + 2) = Format(FMUCR, "#.##")
        InfoMethod(UFL).MethodName(CNT + 2) = "MTU Combustion Reaction"
        InfoMethod(UFL).Enabled(CNT + 2) = True
    End If
    If FMUDI > dummyvalue Or FMUDI < -dummyvalue Then
        InfoMethod(UFL).Enabled(CNT + 3) = False
    Else
        InfoMethod(UFL).value(CNT + 3) = Format(FMUDI, "#.##")
        InfoMethod(UFL).MethodName(CNT + 3) = "Penn State Group Contribution"
        InfoMethod(UFL).Enabled(CNT + 3) = True
    End If
    If FMU > dummyvalue Or FMU < -dummyvalue Then
        InfoMethod(UFL).Enabled(CNT) = False
    Else
        InfoMethod(UFL).value(CNT) = Format(FMU, "#.##")
        InfoMethod(UFL).MethodName(CNT) = "MTU Fire & Explosion Data"
        InfoMethod(UFL).Enabled(CNT) = True
    End If
    ' DATA: (in order)
    '   1.  FMUGP = MTU Value
    '   2.  FMUCR = MTU Method using Combustion Reaction Method
    '   3.  FMUDI = Penn State method
    '   fmu - value from data file
       
        Exit Sub
error_in_upper:
End Sub


Public Function read_database(cas_as_string As String) As Boolean

Dim Block5DBName As String
Dim temp As String
Dim cur_cas As Long
Dim i, J As Integer
Dim success As Boolean
success = True
'On Error GoTo error_reading_database
    ' the cas number the user entered
    
cur_cas = CLng(cas_as_string)

 ' open the database table and find the chemical we want
 
Dim Block5Table As Recordset

Set Block5DB = OpenDatabase(PathBlock5, False, False)
Set Block5Table = Block5DB.OpenRecordset("fexp2", dbOpenTable)
Block5Table.Index = "PrimaryKey"
Block5Table.Seek "=", cur_cas
If Block5Table.NoMatch Then
    GoTo not_found_error
End If

        ' the input from the first line for this entry
    'ICODE = Block5Table![ICODE]
   ' On Error Resume Next
    On Error GoTo resume_error
    ICODE = Block5Table![Dippr Code]
    CHEM_NAME = Block5Table![name]
    For i = 0 To 6
        elr(i) = ""
        nelr(i) = 0
    Next i
    
    elr(0) = Block5Table![ELR1]
    nelr(0) = Block5Table![NELR1]
    elr(1) = Block5Table![ELR2]
    nelr(1) = Block5Table![NELR2]
    elr(2) = Block5Table![ELR3]
    nelr(2) = Block5Table![NELR3]
    elr(3) = Block5Table![ELR4]
    nelr(3) = Block5Table![NELR4]
    elr(4) = Block5Table![ELR5]
    nelr(4) = Block5Table![NELR5]
    elr(5) = Block5Table![ELR6]
    nelr(5) = Block5Table![NELR6]
    elr(6) = Block5Table![ELR7]
    nelr(6) = Block5Table![NELR7]
        ' kind of strange but this means the rest of the code doesn't have to be changed
    temp = " "
    i = 0
    temp = Block5Table![cat]
    While Len(temp) > 0
        cat(i) = Left(temp, 1)
        temp = Right(temp, Len(temp) - 1)
        i = i + 1
    Wend
   ' CAS = Block5Table![NCAS]
    
    '*********second line of input*********************
    ' remember these are now indexed starting at 0 so subtract
    ' one from the group number given in the database
    ' first initialize them just in case there's a null in the database
    For i = 0 To 15
        IGPCOD(i) = NULLCODE
        NGP(i) = 0
    Next i
    
    IGPCOD(0) = Block5Table![IGPCOD0] - 1
    NGP(0) = Block5Table![NGP0]
    IGPCOD(1) = Block5Table![IGPCOD1] - 1
    NGP(1) = Block5Table![NGP1]
    IGPCOD(2) = Block5Table![IGPCOD2] - 1
    NGP(2) = Block5Table![NGP2]
    IGPCOD(3) = Block5Table![IGPCOD3] - 1
    NGP(3) = Block5Table![NGP3]
    IGPCOD(4) = Block5Table![IGPCOD4] - 1
    NGP(4) = Block5Table![NGP4]
    IGPCOD(5) = Block5Table![IGPCOD5] - 1
    NGP(5) = Block5Table![NGP5]
    IGPCOD(6) = Block5Table![IGPCOD6] - 1
    NGP(6) = Block5Table![NGP6]
    IGPCOD(7) = Block5Table![IGPCOD7] - 1
    NGP(7) = Block5Table![NGP7]
    IGPCOD(8) = Block5Table![IGPCOD8] - 1
    NGP(8) = Block5Table![NGP8]
    IGPCOD(9) = Block5Table![IGPCOD9] - 1
    NGP(9) = Block5Table![NGP9]
    IGPCOD(10) = Block5Table![IGPCOD10] - 1
    NGP(10) = Block5Table![NGP10]
    IGPCOD(11) = Block5Table![IGPCOD11] - 1
    NGP(11) = Block5Table![NGP11]
    IGPCOD(12) = Block5Table![IGPCOD12] - 1
    NGP(12) = Block5Table![NGP12]
    IGPCOD(13) = Block5Table![IGPCOD13] - 1
    NGP(13) = Block5Table![NGP13]
    IGPCOD(14) = Block5Table![IGPCOD14] - 1
    NGP(14) = Block5Table![NGP14]
    IGPCOD(15) = Block5Table![IGPCOD15] - 1
    NGP(15) = Block5Table![NGP15]
   
'**************************Begin feeding in the third line of data

FLP = Block5Table![FLP]
bl(2, 0) = Block5Table![BL0]
'temp = Block5Table![BL1]
'For I = 0 To 15
 '   If Left(temp, 1) = " " Or Left(temp, 1) = "" Then Exit For
  '  BL(0, I) = Left(temp, 1)
   ' temp = Right(temp, Len(temp) - 1)
'Next I
FML = Block5Table![LFL]
bl(2, 1) = Block5Table![BL1]
'temp = Block5Table![BL2]
'For I = 0 To 15
 '   If Left(temp, 1) = " " Or Left(temp, 1) = "" Then Exit For
  '  BL(1, I) = Left(temp, 1)
   ' temp = Right(temp, Len(temp) - 1)
'Next I
FMU = Block5Table![UFL]
bl(2, 2) = Block5Table![BL2]
'temp = Block5Table![BL3]
'For I = 0 To 15
 '   If Left(temp, 1) = " " Or Left(temp, 1) = "" Then Exit For
  '  BL(2, I) = Left(temp, 1)
   ' temp = Right(temp, Len(temp) - 1)
'Next I
AITemp = Block5Table![AITemp]
bl(2, 3) = Block5Table![BL3]
'temp = Block5Table![BL4]
'For I = 0 To 15
 '   If Left(temp, 1) = " " Or Left(temp, 1) = "" Then Exit For
 '   BL(3, I) = Left(temp, 1)
 '   temp = Right(temp, Len(temp) - 1)
'Next I
hcom = Block5Table![hcom]
bl(2, 4) = Block5Table![BL4]
'temp = Block5Table![BL5]
'For I = 0 To 15
'    If Left(temp, 1) = " " Or Left(temp, 1) = "" Then Exit For
'    BL(4, I) = Left(temp, 1)
'    temp = Right(temp, Len(temp) - 1)
'Next I
hfor = Block5Table![hfor]
bl(2, 5) = Block5Table![BL5]
'temp = Block5Table![BL6]
'For I = 0 To 15
 '   If Left(temp, 1) = " " Or Left(temp, 1) = "" Then Exit For
 '   BL(5, I) = Left(temp, 1)
 '   temp = Right(temp, Len(temp) - 1)
'Next I
gfor = Block5Table![gfor]
bl(2, 6) = Block5Table![BL6]
'temp = Block5Table![BL7]
'For I = 0 To 15
'    If Left(temp, 1) = " " Or Left(temp, 1) = "" Then Exit For
'    BL(6, I) = Left(temp, 1)
'    temp = Right(temp, Len(temp) - 1)
'Next I
CMW = Block5Table![CMW]

    ' the following are what would have been in the 4-10 lines
    ' of data from the data file (now stored in the database 'input' table
' line 4
Pc = Block5Table![CP]           ' critical pressure of compound
PC_ECODE = Block5Table![CP-ECODE]
neqvp = Block5Table![VPEQ]      ' vapor pressure equation number
vpc(0) = Block5Table![VPC0]    ' regression coefficients for vapor pressure
vpc(1) = Block5Table![VPC1]
vpc(2) = Block5Table![VPC2]
vpc(3) = Block5Table![VPC3]
vpc(4) = Block5Table![VPC4]
' line 5
Tc = Block5Table![CT]           ' critical temp of compound
TC_ECODE = Block5Table![CT-ECODE]
neqhv = Block5Table![HVEQ]      ' heat of vaporization equation number
hvapc(0) = Block5Table![HVAPC0]  ' regression coefficients for heat of vaporization
hvapc(1) = Block5Table![HVAPC1]
hvapc(2) = Block5Table![HVAPC2]
hvapc(3) = Block5Table![HVAPC3]
hvapc(4) = Block5Table![HVAPC4]
' line 6
TBP = Block5Table![NBP]         ' normal boiling point of compound
TBP_ECODE = Block5Table![NBP-ECODE]
neqcpg = Block5Table![HCGEQ]    ' heat capacity of gas equation number
cpgc(0) = Block5Table![CPGC0]   ' regression coeff for gas phase heat capacity
cpgc(1) = Block5Table![CPGC1]
cpgc(2) = Block5Table![CPGC2]
cpgc(3) = Block5Table![CPGC3]
cpgc(4) = Block5Table![CPGC4]
' line 7
mp = Block5Table![mp]           ' was TMP??? melting point of compound
mp_ecode = Block5Table![MP-ECODE]
neqcpl = Block5Table![HCLEQ]    ' heat capacity of liquid equation number
cplc(0) = Block5Table![CPLC0]   ' regression coeff for liquid phase heat capacity
cplc(1) = Block5Table![CPLC1]
cplc(2) = Block5Table![CPLC2]
cplc(3) = Block5Table![CPLC3]
cplc(4) = Block5Table![CPLC4]
' line 8
HFUS = Block5Table![hf]         ' heat of fusion
HFUS_ECODE = Block5Table![HF-ECODE]
neqcps = Block5Table![HCSEQ]    ' heat capacity of solid equation number
cpsc(0) = Block5Table![CPSC0]   ' regression coeff for solid phase heat capacity
cpsc(1) = Block5Table![CPSC1]
cpsc(2) = Block5Table![CPSC2]
cpsc(3) = Block5Table![CPSC3]
cpsc(4) = Block5Table![CPSC4]
' line 9
VLIQ = Block5Table![MV]         ' molar volume of liquid
VLIQ_ECODE = Block5Table![MV-ECODE]
neqdnl = Block5Table![LDEQ]     ' liquid density equation
denlc(0) = Block5Table![DENLC0]  ' regression coeff for liquid density
denlc(1) = Block5Table![DENLC1]
denlc(2) = Block5Table![DENLC2]
denlc(3) = Block5Table![DENLC3]
denlc(4) = Block5Table![DENLC4]
' line 10
neqdns = Block5Table![SDEQ]     ' solid density equation number
densc(0) = Block5Table![DENSC0]  ' regression coefficients for solid density
densc(1) = Block5Table![DENSC1]
densc(2) = Block5Table![DENSC2]
densc(3) = Block5Table![DENSC3]
densc(4) = Block5Table![DENSC4]

Block5Table.Close
Call set_quality_codes
'Block5frm!nametbx.Text = chem_name
read_database = True
Exit Function
error_reading_database:
    If Error = 94 Then Resume Next
    read_database = False
    
    Exit Function
resume_error:
    Debug.Print ("error")
    Resume Next
not_found_error:
    'If Error = 94 Then REsume Next
    read_database = False
    
    Exit Function
End Function

Public Function is_inorganic() As Boolean

    'If cat(0) = "R" And cat(1) = "S" Then
    '    is_inorganic = True
    '    Exit Function
    'End If
    
    If cat(0) = "I" Or (cat(0) = "L" And cat(1) = "X") Or (cat(0) = "X" And cat(1) = "S") Or (cat(0) = "O" And cat(1) = "S") Or (cat(0) = "O" And cat(1) = "I") Then
        is_inorganic = True
    Else
        is_inorganic = False
    End If
End Function


Public Sub do_block5_calcs(flag As Integer, Position As Integer)

    ' a function which reads the database for the chemical
    ' selected and then calls the calculation routines to
    ' calculate flashpoint, upper and lower flammability limits
    
Dim success As Boolean
Dim cur_cas As String
Dim CNT As Integer

    ' an error handler, just aborts the calculation
'On Error GoTo error_reading_db
CNT = Position
cur_cas = Cur_Info.CAS

    ' initialize elements and their values
Call Initialize
    ' read the fire and explosion data
success = read_database(cur_cas)
If success = False Then
    'MsgBox (cur_cas & " not found in Block 5 database")
    Exit Sub
End If
    ' set the elements for this compound
Call set_elements
    ' set the groups in this compound
Call set_groups
    ' set the flag for organic/inorganic
inorganic = is_inorganic
    ' do the calculations (also fills the text boxes on the form)
    Call calc_upper(CNT)
    Call calc_lower(CNT)
    Call calc_flpt(CNT)
    Call calc_AIT(CNT)

    Exit Sub
error_reading_db:
    MsgBox ("Unable to open " & databasename)
    
End Sub


Public Sub set_quality_codes()

    ' QC(0) -> fp
    ' QC(1) -> lfl
    ' QC(2) -> ufl
    ' QC(3) -> ait
    
    QC(0) = bl(2, 0)
    QC(1) = bl(2, 1)
    QC(2) = bl(2, 2)
    QC(3) = bl(2, 3)
End Sub

Public Sub set_elements()
Dim i As Integer
Dim J As Integer
Dim NERR As Integer

ntel = 100
ktel = 7    ' the number of elements in this compound

    For J = 0 To ntel - 1
       NET(J) = 0
    Next J
    
    For i = 0 To 6
        If Trim(elr(i)) = "" Or nelr(i) = 0 Then
            ktel = ktel - 1
            GoTo next_element
        End If
        For J = 0 To ntel - 1
            If Trim(elr(i)) = Trim(EL(J)) Or UCase(Trim(elr(i))) = UCase(Trim(EL(J))) Then
                If nelr(i) <> 0 Then
                    GoTo next_element
                Else
                    NERR = NERR + 1
                    GoTo next_element
                End If
            End If
        Next J
             NERR = NERR + 1
       
next_element:
    Next i
    
   For J = 0 To ntel - 1
        NET(J) = 0
    Next J
      
         ' assign the number of each element to the NET variable
        ' NET seems to hold the number of each element in this
        ' chemical with indexes corresponding to the whole
        ' array of elements
    For i = 0 To ktel - 1
        If Trim(elr(i)) = "" Then
            GoTo next_i_60
        End If
        For J = 0 To ntel - 1
            If Trim(elr(i)) <> Trim(EL(J)) And UCase(Trim(elr(i))) <> UCase(Trim(EL(J))) Then
                GoTo next_j_50
            End If
            NET(J) = nelr(i)
            GoTo next_i_60
next_j_50:
        Next J
next_i_60:
    Next i
    NET(35) = NET(35) + NET(25)
End Sub

Public Sub set_groups()
    Dim i As Integer
    Dim J As Integer
    Dim NERR As Integer
    ntgp = 75
        ' check for errors in the category field
    If Trim(cat(0)) = "" Or Trim(cat(1)) = "" Then
        NERR = NERR + 1
    Else
        ktgp = 16
    End If
    
     ' now get the number of groups
     ' took the error checking out for now
    For i = 0 To 15
        If IGPCOD(i) = NULLCODE Or NGP(i) = 0 Then
            ktgp = ktgp - 1
        End If
    Next i
    
          ' initialize these
    For J = 0 To ntgp - 1
       NGPT(J) = 0
    Next J
    
    ' set the number of each group in the compound
   
    For i = 0 To 15
        If IGPCOD(i) <> NULLCODE Then
            NGPT(IGPCOD(i)) = NGP(i)
        End If
    Next i
End Sub
