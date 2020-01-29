Attribute VB_Name = "moddoconvert"
Option Explicit



Public Function do_Swater_convert(ConvertFrom As String, ConvertTo As String, value As Double) As Double
    ' Author:  DMW
    ' Date:  6/19/97
    ' Purpose:  This function is called when a conversion for Solubility water in chem is
    '           required.
    ' Bugs:     Assumes dilute solutions
    '           needs more units added
    Dim result As Double
    Dim tempMW As Double
    Dim match(5) As Boolean
    
   
        ' get MW in default units in case we need it
    tempMW = 18.015    ' molecular wt of water in kg/kmol
    result = value
    If Trim(ConvertFrom) = Trim(ConvertTo) Or UCase(Trim(ConvertFrom)) = UCase(Trim(ConvertTo)) Then
        do_Swater_convert = value
        Exit Function
    End If
    
    match(0) = Trim(ConvertFrom) Like "ppm*" Or UCase(ConvertFrom) Like UCase("ppm(wt)*")
    match(1) = Trim(ConvertFrom) Like "kmol*m3*" Or UCase(ConvertFrom) Like UCase("kmol*m3*")
    match(2) = Trim(ConvertFrom) Like "mol*L" Or UCase(ConvertFrom) Like UCase("mol*L")
    match(3) = Trim(ConvertFrom) Like "mol*dm3" Or UCase(ConvertFrom) Like UCase("mol*dm3")
    match(4) = Trim(ConvertFrom) Like "mg*kg" Or UCase(ConvertFrom) Like UCase("mg*kg")
    
        ' first convert to default units ppm(wt)
    If match(0) = True Then     ' ppm(wt)
        result = value
    ElseIf match(1) = True Then
        result = result / 1000#
    ElseIf match(2) = True Then
        result = result * tempMW * 1000#
    ElseIf match(3) = True Then
        result = result * tempMW * 1000#
    ElseIf match(4) = True Then
        result = value
    End If
    
    match(0) = Trim(ConvertTo) Like "ppm*" Or UCase(ConvertTo) Like UCase("ppm(wt)*")
    match(1) = Trim(ConvertTo) Like "kmol*m3" Or UCase(ConvertTo) Like UCase("kmol*m3")
    match(2) = Trim(ConvertTo) Like "mol*L" Or UCase(ConvertTo) Like UCase("mol*L")
    match(3) = Trim(ConvertTo) Like "mol*dm3" Or UCase(ConvertTo) Like UCase("mol*dm3")
    match(4) = Trim(ConvertTo) Like "mg*kg" Or UCase(ConvertTo) Like UCase("mg*kg")
    
    If match(0) = True Then
        result = result
    ElseIf match(1) = True Then
        result = result * 1000#
    ElseIf match(2) = True Then
        result = result / tempMW
        result = result / 1000#
    ElseIf match(3) = True Then
        result = result / tempMW
        result = result / 1000#
    ElseIf match(4) = True Then
        result = value
    Else
        result = value  ' fix this
    End If
    do_Swater_convert = result

End Function

Public Sub fill_units_form(PROPERTY_CODE As Long)

    frmunits!CMBUnits.Clear
    Select Case PROPERTY_CODE
    
        Case MW
            frmunits!CMBUnits.AddItem "kg/kmol"
            frmunits!CMBUnits.AddItem "lb/lbmol"
            frmunits!CMBUnits.AddItem "g/gmol"
        Case LD25, LD
            frmunits!CMBUnits.AddItem "kmol/m3"
            frmunits!CMBUnits.AddItem "g/cm3"
            frmunits!CMBUnits.AddItem "g/mL"
            frmunits!CMBUnits.AddItem "g/L"
            frmunits!CMBUnits.AddItem "gmol/L"
            frmunits!CMBUnits.AddItem "mg/mL"
            frmunits!CMBUnits.AddItem "g/m3"
            frmunits!CMBUnits.AddItem "mol/L"
            frmunits!CMBUnits.AddItem "lb/gal"
            frmunits!CMBUnits.AddItem "mmol/L"
            frmunits!CMBUnits.AddItem "ng/L"
            frmunits!CMBUnits.AddItem "mg/L"
            frmunits!CMBUnits.AddItem "mol/cm3"
            frmunits!CMBUnits.AddItem "kg/L"
            frmunits!CMBUnits.AddItem "kg/m3"
            frmunits!CMBUnits.AddItem "lb/ft3"
            frmunits!CMBUnits.AddItem "mol/m3"
            frmunits!CMBUnits.AddItem "mol/dm3"
        Case VP25, VP, CP
            frmunits!CMBUnits.AddItem "Pa"
            frmunits!CMBUnits.AddItem "mPa"
            frmunits!CMBUnits.AddItem "kPa"
            frmunits!CMBUnits.AddItem "MPa"
            frmunits!CMBUnits.AddItem "psig"
            frmunits!CMBUnits.AddItem "lb/ft2"
            frmunits!CMBUnits.AddItem "psia"
            frmunits!CMBUnits.AddItem "atm"
            frmunits!CMBUnits.AddItem "cm Hg"
            frmunits!CMBUnits.AddItem "kN/m2"
            frmunits!CMBUnits.AddItem "bar"
            frmunits!CMBUnits.AddItem "lb/in2"
            frmunits!CMBUnits.AddItem "mm Hg"
            frmunits!CMBUnits.AddItem "mbar"
            frmunits!CMBUnits.AddItem "torr"
        Case hfor, Hvap, Hvap25, HvapNBP, Hcomb
            frmunits!CMBUnits.AddItem "J/kmol"
            frmunits!CMBUnits.AddItem "kJ/mol"
            frmunits!CMBUnits.AddItem "kJ/kmol"
            frmunits!CMBUnits.AddItem "J/mol"
            frmunits!CMBUnits.AddItem "cal/mol"
            frmunits!CMBUnits.AddItem "kcal/mol"
            frmunits!CMBUnits.AddItem "cal/lbmol"
            frmunits!CMBUnits.AddItem "cal/g"
            frmunits!CMBUnits.AddItem "kcal/g"
            frmunits!CMBUnits.AddItem "J/g"
            frmunits!CMBUnits.AddItem "kJ/kg"
            frmunits!CMBUnits.AddItem "Btu/lb"
        Case Dwater, Dair
            frmunits!CMBUnits.AddItem "cm2/s"
            frmunits!CMBUnits.AddItem "m2/s"
            frmunits!CMBUnits.AddItem "in2/s"
            frmunits!CMBUnits.AddItem "ft2/s"
        Case VV, LV
            frmunits!CMBUnits.AddItem "Pa*s"
            frmunits!CMBUnits.AddItem "cp"
            frmunits!CMBUnits.AddItem "kg/m*h"
            frmunits!CMBUnits.AddItem "kg/m*s"
            frmunits!CMBUnits.AddItem "lb/ft*hr"
            frmunits!CMBUnits.AddItem "lb/ft*s"
        Case HC
'msh 4/4/99
            frmunits!CMBUnits.AddItem "Pa*mol/mol"
            frmunits!CMBUnits.AddItem "kPa*mol/mol"
            frmunits!CMBUnits.AddItem "unit-less"
'msh            FRMUnits!CMBUnits.AddItem "atm"
            frmunits!CMBUnits.AddItem "atm*m3/mol"
            frmunits!CMBUnits.AddItem "kPa*m3/kmol"
'msh            FRMUnits!CMBUnits.AddItem "kPa*m3/mol"
            frmunits!CMBUnits.AddItem "MPa*mol/mol"
            frmunits!CMBUnits.AddItem "atm*mol/mol"
'msh            FRMUnits!CMBUnits.AddItem "bar*mol/mol"
'msh            FRMUnits!CMBUnits.AddItem "cm Hg*mol/mol"
'msh            FRMUnits!CMBUnits.AddItem "kN/m2*mol/mol"
'msh            FRMUnits!CMBUnits.AddItem "lbf/in2*mol/mol"
'msh            FRMUnits!CMBUnits.AddItem "lbf/ft2*mol/mol"
'msh            FRMUnits!CMBUnits.AddItem "mPa*mol/mol"
'**msh      frmunits!cmbunits.additem "atm/M"
'**msh      frmunits!cmbunits.additem "mm Hg*mol/mol"
'**msh      frmunits!cmbunits.additem "kPa*m3/mol"
'**msh      frmunits!cmbunits.additem "bar*m3/mol"
'**msh      frmunits!cmbunits.additem "atm*L/mol"
'**msh      frmunits!cmbunits.additem "torr*mol/mol"

            Exit Sub
        Case Schem, Swater
            frmunits!CMBUnits.AddItem "ppm(wt)"
'**msh      frmunits!cmbunits.additem "mg/L"
'**msh      frmunits!cmbunits.additem "mmol/L"
            frmunits!CMBUnits.AddItem "kmol/m3"
            frmunits!CMBUnits.AddItem "mol/L"
            frmunits!CMBUnits.AddItem "mol/dm3"
            frmunits!CMBUnits.AddItem "mg/kg"
            frmunits!CMBUnits.AddItem "g/dm3"
            frmunits!CMBUnits.AddItem "ppb"
            frmunits!CMBUnits.AddItem "g/100 cm3"
            frmunits!CMBUnits.AddItem "g/100 mL"
            frmunits!CMBUnits.AddItem "Molar"
            frmunits!CMBUnits.AddItem "cm3/cm3"
            frmunits!CMBUnits.AddItem "vol%"
            frmunits!CMBUnits.AddItem "cm3/L"
            frmunits!CMBUnits.AddItem "cm3/mL"
            frmunits!CMBUnits.AddItem "mL/L"
            frmunits!CMBUnits.AddItem "cm3/100 cm3"
            frmunits!CMBUnits.AddItem "cm3/100 mL"
            frmunits!CMBUnits.AddItem "mL/100 mL"
            frmunits!CMBUnits.AddItem "g/kg"
            frmunits!CMBUnits.AddItem "mass%"
            frmunits!CMBUnits.AddItem "wt%"
            Exit Sub
        
        Case ST, ST25
            frmunits!CMBUnits.AddItem "N/m"
            frmunits!CMBUnits.AddItem "erg/cm2"
            frmunits!CMBUnits.AddItem "lbf/ft"
            frmunits!CMBUnits.AddItem "lbf/in"
            frmunits!CMBUnits.AddItem "dynes/cm"
            Exit Sub
        Case UFL, LFL
            frmunits!CMBUnits.AddItem "vol% in air"
            Exit Sub
        Case LHC, VHC
            frmunits!CMBUnits.AddItem "J/kmol*K"
            frmunits!CMBUnits.AddItem "Btu/lb*F"
            frmunits!CMBUnits.AddItem "kJ/kmol*K"
            frmunits!CMBUnits.AddItem "cal/g*C"
            frmunits!CMBUnits.AddItem "kJ/kg*K"
            Exit Sub
        Case mp, NBP, CT, FP, AIT
            frmunits!CMBUnits.AddItem "K"
            frmunits!CMBUnits.AddItem "F"
            frmunits!CMBUnits.AddItem "C"
            frmunits!CMBUnits.AddItem "R"
            Exit Sub
        Case ACwater, ACchem, logKow, BCF
            frmunits!CMBUnits.AddItem "unit-less"
            Exit Sub
        Case logKoc
            frmunits!CMBUnits.AddItem "cm3/g OC"
            Exit Sub
        Case ThODcarb, ThODcomb, COD, BOD
            frmunits!CMBUnits.AddItem "g O2/g chem"
            frmunits!CMBUnits.AddItem "mg O2/g chem"
            frmunits!CMBUnits.AddItem "mol O2/g chem"
            Exit Sub
        Case LTC, VTC
            frmunits!CMBUnits.AddItem "W/m*K"
            frmunits!CMBUnits.AddItem "kcal/m*hr*C"
            frmunits!CMBUnits.AddItem "cal/cm*s*C"
            frmunits!CMBUnits.AddItem "Btu/ft*hr*F"
            Exit Sub
        Case CV
            frmunits!CMBUnits.AddItem "m3/kmol"
            frmunits!CMBUnits.AddItem "m3/kg"
            frmunits!CMBUnits.AddItem "m3/g"
            frmunits!CMBUnits.AddItem "cm3/g"
            frmunits!CMBUnits.AddItem "L/g"
            frmunits!CMBUnits.AddItem "L/kg"
            frmunits!CMBUnits.AddItem "L/mg"
            frmunits!CMBUnits.AddItem "mL/mg"
            frmunits!CMBUnits.AddItem "mL/g"
            frmunits!CMBUnits.AddItem "L/ng"
            frmunits!CMBUnits.AddItem "L/mol"
            frmunits!CMBUnits.AddItem "L/mmol"
            frmunits!CMBUnits.AddItem "gal/lb"
            frmunits!CMBUnits.AddItem "ft3/lb"
            Exit Sub
        Case AltSpecies, Mysid96L, Daph24E, Daph48E, Daph24L, Daph48L, _
             Fat48E, Fat96E, Fat24L, Fat48L, Fat96L, Sal24L, _
             Sal48L, Sal96L
             
            frmunits!CMBUnits.AddItem "mg/L"
            Exit Sub
        End Select
End Sub


 
 
Public Function get_mw_element(element As String) As Double

    Dim weight As Double
    weight = 0#
    Select Case Trim(element)
        Case "Ac"
            weight = 227
        Case "Al"
            weight = 26.98154
        Case "Am"
            weight = 243
        Case "Sb"
            weight = 121.75
        Case "Ar"
            weight = 39.948
        Case "As"
            weight = 74.9216
        Case "At"
            weight = 210
        Case "Ba"
            weight = 137.33
        Case "Bk"
            weight = 247
        Case "Be"
            weight = 9.01218
        Case "Bi"
            weight = 208.9804
        Case "B"
            weight = 10.81
        Case "Br"
            weight = 79.904
        Case "Cd"
            weight = 112.41
        Case "Ca"
            weight = 40.08
        Case "Cf"
            weight = 251
        Case "C"
            weight = 12.011
        Case "Ce"
            weight = 140.12
        Case "Cs"
            weight = 132.9054
        Case "Cl"
            weight = 35.453
        Case "Cr"
            weight = 51.996
        Case "Co"
            weight = 58.9332
        Case "Cu"
            weight = 63.546
        Case "Cm"
            weight = 247
        Case "Dy"
            weight = 162.5
        Case "Es"
            weight = 254
        Case "Er"
            weight = 167.26
        Case "Eu"
            weight = 151.96
        Case "Fm"
            weight = 257
        Case "F"
            weight = 18.998403
        Case "Fr"
            weight = 223
        Case "Gd"
            weight = 157.25
        Case "Ga"
            weight = 69.72
        Case "Ge"
            weight = 72.59
        Case "Au"
            weight = 196.9665
        Case "Hf"
            weight = 178.49
        Case "He"
            weight = 4.0026
        Case "Ho"
            weight = 164.9304
        Case "H"
            weight = 1.0079
        Case "In"
            weight = 114.82
        Case "I"
            weight = 126.9045
        Case "Ir"
            weight = 192.22
        Case "Fe"
            weight = 55.847
        Case "Kr"
            weight = 83.8
        Case "La"
            weight = 138.9055
        Case "Lr"
            weight = 260
        Case "Pb"
            weight = 207.2
        Case "Li"
            weight = 6.941
        Case "Lu"
            weight = 174.97
        Case "Mg"
            weight = 24.305
        Case "Mn"
            weight = 54.938
        Case "Md"
            weight = 258
        Case "Hg"
            weight = 200.59
        Case "Mo"
            weight = 95.94
        Case "Nd"
            weight = 114.24
        Case "Ne"
            weight = 20.179
        Case "Np"
            weight = 237.0482
        Case "Ni"
            weight = 58.7
        Case "Nb"
            weight = 92.9064
        Case "N"
            weight = 14.0067
        Case "No"
            weight = 255
        Case "Os"
            weight = 190.2
        Case "O"
            weight = 15.9994
        Case "Pd"
            weight = 106.4
        Case "P"
            weight = 30.97376
        Case "Pt"
            weight = 195.09
        Case "Pu"
            weight = 244
        Case "Po"
            weight = 209
        Case "K"
            weight = 39.0983
        Case "Pr"
            weight = 140.9077
        Case "Pm"
            weight = 145
        Case "Pa"
            weight = 231.0359
        Case "Ra"
            weight = 226.0254
        Case "Rn"
            weight = 222
        Case "Re"
            weight = 186.207
        Case "Rh"
            weight = 102.9055
        Case "Rb"
            weight = 85.4678
        Case "Ru"
            weight = 101.07
        Case "Sm"
            weight = 150.4
        Case "Sc"
            weight = 44.9559
        Case "Se"
            weight = 78.96
        Case "Si"
            weight = 28.0855
        Case "Ag"
            weight = 107.868
        Case "Na"
            weight = 22.98977
        Case "Sr"
            weight = 87.62
        Case "S"
            weight = 32.06
        Case "Ta"
            weight = 180.9479
        Case "Tc"
            weight = 97
        Case "Te"
            weight = 127.6
        Case "Tb"
            weight = 158.9254
        Case "Tl"
            weight = 204.37
        Case "Th"
            weight = 232.0381
        Case "Tm"
            weight = 168.9342
        Case "Sn"
            weight = 118.69
        Case "Ti"
            weight = 47.9
        Case "W"
            weight = 183.85
        Case "U"
            weight = 238.029
        Case "V"
            weight = 50.9414
        Case "Xe"
            weight = 131.3
        Case "Yb"
            weight = 173.04
        Case "Y"
            weight = 88.9059
        Case "Zn"
            weight = 65.38
        Case "Zr"
            weight = 91.22
    End Select
    get_mw_element = weight
End Function




Public Function find_elements(selected_structure As String, elements() As String, num_elements() As Integer) As Boolean

    ' this is taken from the element finder routine in dbman, to be used only when we don't have
    ' a mw for an element (usually during unit conversions)
    
    ' this function finds the elements in a compound based
    ' on it's formula
    Dim i As Integer
    Dim J As Integer
   
    Dim count As Integer
    Dim local_EL(100) As String
    Dim structure_string As String
    Dim cur_string As String
    Dim found As Boolean
    Dim tempchar As String
    If Trim(selected_structure) = "" Then
        MsgBox ("need chemical formula to find local_ELements")
        find_elements = False
        Exit Function
    End If
    On Error GoTo failed_finder
    ' initialize the possible elements
local_EL(0) = "A"
local_EL(1) = "Ac"
local_EL(2) = "Ag"
local_EL(3) = "Al"
local_EL(4) = "Am"
local_EL(5) = "As"
local_EL(6) = "At"
local_EL(7) = "Au"
local_EL(8) = "B"
local_EL(9) = "Ba"
local_EL(10) = "Be"
local_EL(11) = "Bi"
local_EL(12) = "Bk"
local_EL(13) = "Br"
local_EL(14) = "C"
local_EL(15) = "Ca"
local_EL(16) = "Cd"
local_EL(17) = "Ce"
local_EL(18) = "Cf"
local_EL(19) = "Cl"
local_EL(20) = "Cm"
local_EL(21) = "Co"
local_EL(22) = "Cr"
local_EL(23) = "Cs"
local_EL(24) = "Cu"
local_EL(25) = "D"
local_EL(26) = "Dy"
local_EL(27) = "Er"
local_EL(28) = "Eu"
local_EL(29) = "F"
local_EL(30) = "Fe"
local_EL(31) = "Fr"
local_EL(32) = "Ga"
local_EL(33) = "Gd"
local_EL(34) = "Ge"
local_EL(35) = "H"
local_EL(36) = "He"
local_EL(37) = "Hf"
local_EL(38) = "Hg"
local_EL(39) = "Ho"
local_EL(40) = "I"
local_EL(41) = "In"
local_EL(42) = "Ir"
local_EL(43) = "K"
local_EL(44) = "Kr"
local_EL(45) = "La"
local_EL(46) = "Li"
local_EL(47) = "Lu"
local_EL(48) = "Mg"
local_EL(49) = "Mn"
local_EL(50) = "Mo"
local_EL(51) = "Mv"
local_EL(52) = "N"
local_EL(53) = "Na"
local_EL(54) = "Nb"
local_EL(55) = "Nd"
local_EL(56) = "Ne"
local_EL(57) = "Ni"
local_EL(58) = "Np"
local_EL(59) = "O"
local_EL(60) = "Os"
local_EL(61) = "P"
local_EL(62) = "Pa"
local_EL(63) = "Pb"
local_EL(64) = "Pd"
local_EL(65) = "Pm"
local_EL(66) = "Po"
local_EL(67) = "Pr"
local_EL(68) = "Pt"
local_EL(69) = "Pu"
local_EL(70) = "Ra"
local_EL(71) = "Rb"
local_EL(72) = "Re"
local_EL(73) = "Rh"
local_EL(74) = "Rn"
local_EL(75) = "Ru"
local_EL(76) = "S"
local_EL(77) = "Sb"
local_EL(78) = "Sc"
local_EL(79) = "Se"
local_EL(80) = "Si"
local_EL(81) = "Sm"
local_EL(82) = "Sn"
local_EL(83) = "Sr"
local_EL(84) = "Ta"
local_EL(85) = "Tb"
local_EL(86) = "Tc"
local_EL(87) = "Te"
local_EL(88) = "Th"
local_EL(89) = "Ti"
local_EL(90) = "Tl"
local_EL(91) = "Tm"
local_EL(92) = "U"
local_EL(93) = "V"
local_EL(94) = "W"
local_EL(95) = "Xe"
local_EL(96) = "Y"
local_EL(97) = "Yb"
local_EL(98) = "Zn"
local_EL(99) = "Zr"
' now find what's in this compound
structure_string = Trim(selected_structure)

cur_string = Left(structure_string, 1)
structure_string = Right(structure_string, Len(structure_string) - 1)
' take care of elements with two letters
If Len(structure_string) > 1 Then
    tempchar = Left(structure_string, 1)
    If Asc(tempchar) < 123 And Asc(tempchar) > 96 Then
        cur_string = cur_string & tempchar
        structure_string = Right(structure_string, Len(structure_string) - 1)
    End If
End If
        
count = 0
While Len(structure_string) > -1
    found = False
    
    For J = 0 To 99
        
        If local_EL(J) = cur_string Then
            found = True
            'structure_string = Right(structure_string, Len(structure_string) - 1)
            elements(count) = cur_string
            If Len(structure_string) > 1 And IsNumeric(Left(structure_string, 2)) Then
                num_elements(count) = CInt(Left(structure_string, 2))
                structure_string = Right(structure_string, Len(structure_string) - 2)
                count = count + 1
                GoTo next_element
            End If
            If Len(structure_string) > 0 And IsNumeric(Left(structure_string, 1)) Then
                num_elements(count) = CInt(Left(structure_string, 1))
                structure_string = Right(structure_string, Len(structure_string) - 1)
            Else
                num_elements(count) = 1
            End If
            count = count + 1
            GoTo next_element
        End If
    Next J
next_element:
    If found = False Then
        'cur_string = cur_string & Left(structure_string, 1)
    Else
        If Len(structure_string) > 0 Then
            cur_string = Left(structure_string, 1)
            structure_string = Right(structure_string, Len(structure_string) - 1)
            ' take care of elements with two letters
            If Len(structure_string) > 1 Then
                tempchar = Left(structure_string, 1)
                    If Asc(tempchar) < 123 And Asc(tempchar) > 96 Then
                        cur_string = cur_string & tempchar
                        structure_string = Right(structure_string, Len(structure_string) - 1)
                    End If
            End If
        Else
            GoTo done_find
        End If
    End If
Wend
done_find:
    find_elements = True
    Exit Function

failed_finder:
    find_elements = False
End Function



Public Function get_MW(unit_desired As String, casno As Long) As Double

    Dim structure_string As String
    Dim answer As Double
    Dim elements(16) As String
    Dim num_elements(16) As Integer
    Dim i As Integer
    Dim success As Boolean
    Dim subtotal As Double
    subtotal = 0#
    structure_string = Trim(Cur_Info.Formula)
    If structure_string <> "" Then
        success = find_elements(structure_string, elements, num_elements)
        If success = True Then
            For i = 0 To 15
                If elements(i) <> "" Then
                    subtotal = get_mw_element(elements(i))
                    subtotal = subtotal * num_elements(i)
                Else
                    Exit For
                End If
            Next i
        Else
            get_MW = ERROR_FLAG
            Exit Function
        End If
    Else
        get_MW = ERROR_FLAG
        Exit Function
    End If
    get_MW = answer
End Function

 
