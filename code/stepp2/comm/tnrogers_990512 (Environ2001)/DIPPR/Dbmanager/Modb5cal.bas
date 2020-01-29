Attribute VB_Name = "modb5calc"
Option Explicit

Dim Block5Table As Recordset
Dim Block5DB As Database
Dim PathBlock5 As String
Dim ntgp As Integer ' the number of groups
Dim ntel As Integer ' the number of elements
Dim ktel As Integer ' the number of elements in compound
Dim ktgp As Integer ' the number of groups in compound
Dim NGPT(100) As Integer
Dim NET(120) As Integer
Dim EL(120) As String
Dim CNU(120) As Double
Dim elr(7) As String
Dim nelr(7) As Integer
Dim IGPCOD(16) As Integer    ' group number for each chemical group in compound
Dim NGP(16) As Integer       ' number of each group in compound
 
    ' values from fire and explosion data file
    ' for dbman not sure yet where to get these (DENISE fix)
Dim FML As Double  ' lower flammability limit
Dim FLP As Double   ' flashpoint
Dim FMU As Double  ' upper flammability limit
Dim AITemp As Double  ' autoignition temp
Dim hcom, hfor, gfor, CMW As Double


Public Sub calc_upper(FMUDI As Double, FMUCR As Double, FMUGP As Double, organic As Boolean)
    
Dim dummyvalue As Double
Dim AIR As Double
Dim CAIR As Double
Dim NETC(120) As Integer
Dim HUFLDI(100) As Double
Dim HUFLGP(100) As Double
Dim NGPDI, SNUHDI, SNUHGP As Double  ' ??? not sure of type
Dim curindex As Integer
Dim I, J As Integer

    
Call init_constants
Call set_elements
Call set_groups

' comment this stuff out, for now we'll do it all in this function
' note:  this has to assume we have all inputs
If check_inputs(UFL, 0) = False Then
    Exit Sub
End If
'If check_inputs(UFL, 1) = True Then
'    FMUCR = do_ufl_fmucr
'End If
'If check_inputs(UFL, 2) = True Then
'    FMUGP = do_ufl_fmugp
'End If

dummyvalue = 1E+15
CAIR = 3.8
 
        ' initialize all these to 0
For I = 0 To ntgp - 1
      HUFLGP(I) = 0#
      HUFLDI(I) = 0#
201 Next I

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
        
        Set Block5DB = OpenDatabase(PathBlock5 & "\block5.mdb", False, False)
        Set Block5Table = Block5DB.OpenRecordset("UFLHORG", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            HUFLGP(curindex - 1) = Block5Table("data")
            Block5Table.MoveNext
        Wend
        Block5Table.Close
        
    If organic = False Then
        Set Block5Table = Block5DB.OpenRecordset("UFLHINO", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            If Block5Table("data") <> 0# Then
                HUFLGP(curindex - 1) = Block5Table("data")
            End If
            Block5Table.MoveNext
        Wend
        Block5Table.Close
    End If
    Block5DB.Close
    
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
      
    
    
   
continue_605:
    
    'If FMU > 100# Then FMU = 0#
    'If FMU > 100# Then FMU = 0#
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
      'If FMU < 0# Then FMU = -1E+15
      'If FMU = 0# Then FMU = 1E+15
      If FMUDI = 0# Then FMUDI = 1E+15
      If FMUGP = 0# Then FMUGP = 1E+15
      If FMUCR = 0# Then FMUCR = 1E+15
      

    ' DATA: (in order)
    '   1.  FMUGP = MTU Value
    '   2.  FMUCR = MTU Method using Combustion Reaction Method
    '   3.  FMUDI = Penn State method
    '   fmu - value from data file
       
   

error_in_upper:


End Sub

Public Sub calc_lower(FMLDI As Double, FMLCR As Double, FMLGP As Double, organic As Boolean)
   
Dim NERR, JF As Integer
Dim CFP, ICODEP, NCODR, JL, SNUHDI, SNUHGP, AIR, CAIR As Double ' ?? not sure of type
Dim VPFP As Double
Dim I, J As Integer
Dim HLFLDI(100) As Double
Dim HLFLGP(100) As Double
Dim dummyvalue As Double
    ' following ok
Dim NETC(120) As Integer
Dim string1, string2, string3 As String
Dim curindex As Integer
   
Call init_constants
Call set_elements
Call set_groups

' comment this stuff out, for now we'll do it all in this function
' note:  this has to assume we have all inputs
If check_inputs(LFL, 0) = False Then
    Exit Sub
End If
'If check_inputs(LFL, 1) = True Then
'    FMLCR = do_lfl_fmlcr
'End If
'If check_inputs(LFL, 2) = True Then
'    FMLGP = do_lfl_fmlgp
'End If

On Error GoTo error_in_lower
    dummyvalue = 1E+15
    FMLDI = dummyvalue
    FMLGP = dummyvalue
    FMLCR = dummyvalue
    CAIR = 0.512
    CFP = 3#  ' was 0, set to 3 just for LFL
    ICODEP = 0
    NERR = 0
    NCODR = 0

For I = 0 To ntgp - 1
      HLFLGP(I) = 0#
      HLFLDI(I) = 0#
Next I
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
        Set Block5DB = OpenDatabase(PathBlock5 & "\block5.mdb", False, False)
        Set Block5Table = Block5DB.OpenRecordset("LFLHORG", dbOpenTable)
        While Not Block5Table.EOF
            curindex = Block5Table("groupindex")
            HLFLGP(curindex - 1) = Block5Table("data")
            Block5Table.MoveNext
        Wend
        Block5Table.Close
    
    If organic = False Then
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
    Block5DB.Close


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
    'If FML > 100# Or FML = 100 Then FMLDI = 1E+15
    'If FML > 100# Or FML = 100 Then FMLGP = 1E+15
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
    'If FML > 100# Or FML = 100 Then FMLCR = 1E+15
    'FMLFP = -1E+15
    
    'If FLP < 0# Then GoTo dont_calc_vp_method
    'If FLP = 0# Then
     '   FMLFP = 0#
     '   GoTo dont_calc_vp_method
    'End If
    'If FLP > 10000# Then
    '    FMLFP = 1E+15
    '    GoTo dont_calc_vp_method
    'End If
    'If neqvp <> 0 Then
    '    Call EQNSUBL(FLP - CFP, VPFP)
    '    FMLFP = VPFP * 100# / 101325#
    '    If FMLFP > 100# Or FMLFP = 100 Then FMLFP = dummyvalue
    '    If FML > 100# Or FML = 100 Then FMLFP = dummyvalue
    'End If
    'If vpc(1) = 0# And vpc(2) = 0# And vpc(3) = 0# And vpc(4) = 0# Then
    '    FMLFP = dummyvalue
    'End If
dont_calc_vp_method:
      'If FML > 100# Or FML = 100 Then QC(0) = "NC"
      'If FML < 0# Then QC(0) = "NA"
      

        ' the values in order of preference
        ' 1.  FMLGP -> MTU LFL Group Contribution data
        ' 2.  FMLDI -> Penn State U. Data
        ' 3.  FMLCR -> MTU for Combustion Reaction
        ' 4.  FMLFP -> Flashpoint Method
        ' fml - number from data file

   

error_in_lower:

End Sub

Public Sub init_constants()

    Dim I, J As Integer
    ' get the block 5 path for these calcs
    PathBlock5 = find_file("block5.mdb")
    ' This is the global that is used as a delimiter in the data file????

For I = 0 To 6
    nelr(I) = 0
    elr(I) = " "
Next I
For I = 0 To 15
    NGP(I) = 0
    IGPCOD(I) = 0
Next I

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

Public Sub set_elements()

Dim I As Integer
Dim J As Integer
Dim element_count As Integer
' get the elements from the edit wizard screen

ntel = 100
ktel = 7    ' the number of elements in this compound
' first make sure there are some elements there
frmeditwizard!grdelements.Row = 1
frmeditwizard!grdelements.Col = 1
If Trim(frmeditwizard!grdelements.Text) = "" Then
    MsgBox "no elements listed for this compound"
End If
' get the elements from the edit wizard screen
    element_count = 0
    For I = 0 To ktel - 1
        frmeditwizard!grdelements.Row = I + 1
        frmeditwizard!grdelements.Col = 1
        If Trim(frmeditwizard!grdelements.Text) <> "" Then
            elr(I) = Trim(frmeditwizard!grdelements.Text)
            frmeditwizard!grdelements.Col = 2
            nelr(I) = CInt(Trim(frmeditwizard!grdelements.Text))
            element_count = element_count + 1
        Else
            elr(I) = 0
            nelr(I) = 0
        End If
    Next I
    ' get the number of elements
    ktel = element_count
    
    For J = 0 To ntel - 1
       NET(J) = 0
    Next J
      
         ' assign the number of each element to the NET variable
        ' NET seems to hold the number of each element in this
        ' chemical with indexes corresponding to the whole
        ' array of elements
    For I = 0 To ktel - 1
        If Trim(elr(I)) = "" Then
            GoTo next_i_60
        End If
        For J = 0 To ntel - 1
            If Trim(elr(I)) <> Trim(EL(J)) And UCase(Trim(elr(I))) <> UCase(Trim(EL(J))) Then
                GoTo next_j_50
            End If
            NET(J) = nelr(I)
            GoTo next_i_60
next_j_50:
        Next J
next_i_60:
    Next I
    NET(35) = NET(35) + NET(25)
    NET(25) = 0#
End Sub

Public Sub set_groups()

    Dim I As Integer
    Dim J As Integer
    ntgp = 75
    ktgp = 16
           ' initialize these
    For J = 0 To ntgp - 1
       NGPT(J) = 0
    Next J
    
    I = 0
    While I < ktgp - 1
        If cur_chem_groups(I) <> -1 Then
            'IGPCOD(I) = CInt(Trim(frmeditwizard!lblgroup(I).Text))
            NGPT(cur_chem_groups(I)) = num_cur_chem_groups(I)
            I = I + 1
        Else
            GoTo done_assign
        End If
    Wend
done_assign:
   ' now get the number of groups
   ktgp = I
    
   
    
End Sub

Public Sub calc_AIT(AITG1 As Double, AITGP As Double, organic As Boolean)

    Dim SNUHGP As Double
    Dim SNUHG1 As Double
    Dim SNUHA0 As Double
    Dim ICODEP As Integer
    Dim A0G As Double
    Dim A0G1 As Double
    Dim EAIT(2)  As Double
    Dim dummyvalue As Double
    Dim NETC(120) As Integer
    Dim HAITGP(100) As Double
    Dim HAITA0(100) As Double
    Dim HAITG1(100) As Double
    Dim NAIT(10) As Double
    Dim IGCOD(16) As Double ' need this?
    Dim JA As Integer
    Dim I, J As Integer
    Dim curindex As Integer
    Dim NERR As Integer
    
Call init_constants
Call set_elements
Call set_groups

' comment this stuff out, for now we'll do it all in this function
' note:  this has to assume we have all inputs
If check_inputs(AIT, 0) = False Then
    Exit Sub
End If
'If check_inputs(AIT, 1) = True Then
'    AITGP = do_ait_aitgp
'End If

    
        ' values from tables
    
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
    For I = 0 To ntgp - 1
        HAITA0(I) = 0
        HAITGP(I) = 0
        HAITG1(I) = 0
    Next I
        
        Set Block5DB = OpenDatabase(PathBlock5 & "\block5.mdb", False, False)
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
   
   If organic = False Then
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
   Block5DB.Close
    For I = 0 To ntel - 1
        NETC(I) = 0
    Next I
    
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

    For I = 0 To ntel - 1
        If NETC(I) = NET(I) Then
            GoTo next_i_netc
        End If
        ' Print Debug "Error in BALANCE near ???"
        NERR = NERR + 1
        
next_i_netc:
    Next I
    
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
        If UCase(Trim(selected_name)) = "AIR" Then
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
        
        ' The values of interest (in order of preference):
        '  AITemp = MTU Fire and Explosion Data File
        '  AITGP  = MTU Logarithmic Method
        '  AITG1  = MTU Linear Method
        
        
   
error_in_ait:
End Sub

Public Sub calc_FP(FPTGP As Double, organic As Boolean)
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
Dim NETC(120) As Integer
Dim NERR, NCODR As Integer
Dim JF, JL, ICODEP As Integer
Dim I, J, curindex As Integer
Dim dummyvalue As Double
Dim inorganic As Boolean
Call init_constants
Call set_elements
Call set_groups

' comment this stuff out, for now we'll do it all in this function
' note:  this has to assume we have all inputs
If check_inputs(FP, 0) = False Then
    Exit Sub
End If


dummyvalue = 1E+15
On Error GoTo error_in_flpt
CAIR = 0.512
CFP = 0#
ICODEP = 0
NERR = 0
NCODR = 0

For I = 0 To ntgp - 1
    HLFLGP(I) = 0#
    HLFLDI(I) = 0
Next I
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
    Set Block5DB = OpenDatabase(PathBlock5 & "\block5.mdb", False, False)
    Set Block5Table = Block5DB.OpenRecordset("LFLHORG", dbOpenTable)
    While Not Block5Table.EOF
        curindex = Block5Table("groupindex")
        HLFLGP(curindex - 1) = Block5Table("data")
        Block5Table.MoveNext
    Wend
    Block5Table.Close

If organic = False Then
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
Block5DB.Close
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

continue_JF_calc:

  
continue_JL_calc:

'If FLP > 10000# Then FLP = 1E+15
If NGPT(75) <> 0 Then FLP = 1E+15
'If FML >= 100# Then FLP = 1E+15
DA = -1E+15
DI = -1E+15
GP = -1E+15
CR = -1E+15
'If FML >= 100# Then DA = 1E+15
'If FML >= 100# Then DI = 1E+15
'If FML >= 100# Then GP = 1E+15
'If FML >= 100# Then CR = 1E+15
'If neqvp = 0 Then GoTo 403

'For J = 1 To 4
'      If vpc(J) <> 0# Then GoTo 402
'Next J
'GoTo 403
'402 If NGPT(74) <> 0 Then FML = 100#
'If FML >= 100# Then
'    FPTDA = 1E+15
'ElseIf FML <= 0# Then
'    FPTDA = -1E+15
'ElseIf JL > 10 Then
'    FPTDA = -1E+15
'Else
'    CLFL = FML
'    FPTI = 550#
'    If TBP > 0# Or TBP = 0# Then FPTI = TBP
    ' for these purposes DNEQNFP will always use the NEQVP equation
'    Call DNEQNFP(0.00001, 1, 100, FPTI, DFPTDA, FNORM, CLFL)
'    FPTDA = DFPTDA + CFP
'End If


SNUHDI = NGPT(19) * 14.07
SNUHGP = 0#
For J = 0 To ntgp - 1
    SNUHGP = SNUHGP + NGPT(J) * HLFLGP(J)
Next J
    
For J = 0 To ntel - 1
    SNUHDI = SNUHDI + CDbl(NET(J)) * HLFLDI(J)
Next J

'If SNUHDI <= 1# Then FMLDI = 100#
'If SNUHDI > 1# Then FMLDI = 100# / SNUHDI
'If SNUHGP <= 1# Then FMLGP = 100#
'If SNUHGP > 1# Then FMLGP = 100# / SNUHGP
'If NGPT(74) <> 0 Then FMLDI = 100#
'If NGPT(74) <> 0 Then FMLGP = 100#
'If FMLDI <= 0# Then
'    FPTDI = -1E+15
'ElseIf FMLDI > 100# Or FMLDI = 100# Then
'    FPTDI = 1E+15
'Else
'CLFL = FMLDI
'FPTI = 550#
'If TBP > 0# Or TBP = 0# Then FPTI = TBP

    ' again, this function will use the VP equation
'Call DNEQNFP(0.00001, 1, 100, FPTI, DFPTDI, FNORM, CLFL)
'FPTDI = DFPTDI + CFP
'End If
'If FMLGP <= 0# Then
    FPTGP = -1E+15
'ElseIf FMLGP >= 100# Then
 '   FPTGP = 1E+15
'Else
'    CLFL = FMLGP
'    FPTI = 550#
'    If TBP > 0# Or TBP = 0# Then FPTI = TBP
'    Call DNEQNFP(0.00001, 1, 100, FPTI, DFPTGP, FNORM, CLFL)
'    FPTGP = DFPTGP + CFP
'End If
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
'If AIR <= 0# Then FMLCR = 100#
'If AIR > 0# Then FMLCR = 100# / (1# + AIR / CAIR)
'If NGPT(74) <> 0 Then FMLCR = 100#
'If FMLCR >= 100# Then
'    FPTCR = 1E+15
'ElseIf FMLCR <= 0# Then
'    FPTCR = -1E+15
'Else
'    CLFL = FMLCR
 '   FPTI = 550#
 '   If TBP > 0# Or TBP = 0 Then FPTI = TBP
        ' this function will use the VP equation
 '   Call DNEQNFP(0.00001, 1, 100, FPTI, DFPTCR, FNORM, CLFL)
 '   FPTCR = DFPTCR + CFP
'End If
    ' this is quality code stuff, not currently being used by pearls
'403   QC(0) = QB(JF)  ' fix JF to be correct index ????

    ' the following just give a little feedback for debugging purposes
    ' indicates the compound is non-combustible
'If FML >= 100# Then
'    QC(0) = "NC"    ' was flm ???
    'Block5frm!notelbl.Caption = "non-combustible"
'End If
    ' indicates fpt is not applicable
'If FLP < 0# Then
 '   QC(0) = "NA"
 '   Block5frm!notelbl.Caption = "not applicable"
'End If
On Error GoTo error_in_flpt


    ' now, the data we want in order of preference:
    '   1. FPTDA -> based on LFL Data
    '   2. FPTGP -> based on MTU LFL from Group Contributions
    '   3. FPTDI -> based on Penn State U FL
    '   4. FPTCR -> based on MTU LFL from Combustion Reaction
    '   FLP = flashpoint data (from file)

   

error_in_flpt:
End Sub
