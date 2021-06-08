Attribute VB_Name = "modtools"
Option Explicit

Public Function do_dbman_shredder() As Integer

    ' arguments:    groupfile -> the file containing the groups for input
    '               smilesarg -> the smiles string we're disassembling
    '               search_type -> the type of search (1 or 2???)
    
    Dim testfile As String
    Dim i As Integer
    Dim j As Integer
    On Error GoTo error_in_disassembly
    ' first check that the group file is there because a nasty
    ' error is generated in the DLL if it's not
    testfile = Dir(AppPath & "" & global_groupfile)
    If Trim(testfile) <> "" Then
        If FileLen(AppPath & "" & global_groupfile) < 10 Then
            MsgBox (global_groupfile & " not found")
            do_dbman_shredder = -1  ' indicates failure??
            Exit Function
        End If
    Else
        MsgBox (global_groupfile & " not found")
        do_dbman_shredder = -1
        Exit Function
    End If
        
    ' must be there, ok to proceed
    Call load_form_struct
    frmstruct.Show 1
   
    ' finally, fill in the wizard form groups grid
    Call update_groups
    
    do_dbman_shredder = 1
    
    Exit Function
    
    
        
       
    
error_in_disassembly:
        MsgBox ("an error occurred in disassembly")
        do_dbman_shredder = -1
End Function


Public Function do_element_finder() As Boolean
    ' this function finds the elements in a compound based
    ' on it's formula
    Dim i As Integer
    Dim j As Integer
    Dim elements(16) As String
    Dim num_elements(16) As Integer
    Dim count As Integer
    Dim local_EL(100) As String
    Dim structure_string As String
    Dim cur_string As String
    Dim found As Boolean
    Dim tempchar As String
    If Trim(selected_structure) = "" Then
        MsgBox ("need chemical formula to find local_ELements")
        do_element_finder = False
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
    
    For j = 0 To 99
        
        If local_EL(j) = cur_string Then
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
    Next j
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
    frmeditwizard!grdelements.ColWidth(0) = 100
    frmeditwizard!grdelements.ColWidth(1) = 900
    frmeditwizard!grdelements.ColWidth(2) = 600
    
For i = 1 To count
    frmeditwizard!grdelements.Row = i
    frmeditwizard!grdelements.Col = 1
    frmeditwizard!grdelements.Text = elements(i - 1)
    frmeditwizard!grdelements.Col = 2
    frmeditwizard!grdelements.Text = num_elements(i - 1)
Next i
For i = count + 1 To 8
    frmeditwizard!grdelements.Row = i
    frmeditwizard!grdelements.Col = 1
    frmeditwizard!grdelements.Text = ""
    frmeditwizard!grdelements.Col = 2
    frmeditwizard!grdelements.Text = ""
Next i
frmeditwizard!grdelements.Refresh
do_element_finder = True
Exit Function

failed_finder:
do_element_finder = False
End Function

Public Function get_prop_code(propnum As Integer) As String

    ' this function returns the dippr property code based
    ' on the global cur property (PEARLS Code)
    Dim answer As String
     Select Case propnum
            Case BOD
                answer = "1a"
            Case COD
                answer = "1b"
            Case ThODcarb
                answer = "1cc"
            Case ThODcomb
                answer = "1cn"
            Case logKow
                answer = "2a"
            Case logKoc
                answer = "2c"
            Case BCF
                answer = "2d"
            Case MW
                answer = "3a"
            Case LD25
                answer = "3b"
            Case LD
                answer = "3bt"
            Case Schem
                answer = "3c"
            Case mp
                answer = "3d"
            Case NBP
                answer = "3e"
            Case VP25
                answer = "3f"
            Case Vp
                answer = "3g"
            Case Dair
                answer = "3h"
            Case Dwater
                answer = "3i"
            Case VV
                answer = "3j"
            Case LV
                answer = "3k"
            Case ST25
                answer = "3l"
            Case ST
                answer = "3lt"
            Case LTC
                answer = "3ml"
            Case VTC
                answer = "3mv"
            Case hfor
                answer = "3n"
            Case LHC
                answer = "3o"
            Case VHC
                answer = "3p"
            Case CT
                answer = "3q"
            Case CP
                answer = "3r"
            Case CV
                answer = "3s"
            Case Hvap25
                answer = "3t"
            Case Hvap
                answer = "3tt"
            Case HvapNBP
                answer = "3tz"
            Case ACchem
                answer = "4a"
            Case ACwater
                answer = "4b"
            Case HC
                answer = "4c"
            Case LFL
                answer = "5al"
            Case UFL
                answer = "5au"
            Case FP
                answer = "5b"
            Case AIT
                answer = "5c"
            Case Hcomb
                answer = "5d"
            Case Swater
                answer = "2b"
            Case Else
                answer = " "
        End Select
    get_prop_code = answer
End Function

Public Function get_reference_chemical(family_group As String) As Long

Dim temptable As Recordset
Dim ref_cas As Long
Dim found As Boolean

Set temptable = chembrowsedb.OpenRecordset("reference chemicals", dbOpenSnapshot)
temptable.MoveFirst
found = False
ref_cas = 0
While Not temptable.EOF
    If Trim(temptable("family group")) = Trim(family_group) Then
        found = True
        ref_cas = temptable("CAS")
        GoTo finish
    End If
    temptable.MoveNext
Wend
finish:
temptable.Close
If found = True Then
    get_reference_chemical = ref_cas
Else
    get_reference_chemical = 0
End If
End Function



Public Function is_valid_loll_range(cas_arg As Long, T_arg As Double, T_Units As String) As Boolean

    ' this functions needs the temperature in C to compare to the valid
    ' range in the reference chemical table
    
    Dim temptable As Recordset
    Dim local_temp As Double    ' will hold the temperature for comparison in case we need a conversion
    Dim answer As Boolean
    
    'If T_units <> "C" Then
    '    local_temp = Convert()
    'Else
        local_temp = T_arg
    'End If
    
    answer = True
    On Error Resume Next
    Set temptable = chembrowsedb.OpenRecordset("reference chemicals", dbOpenSnapshot)
    temptable.FindFirst "CAS = " & Val(cas_arg)
    If Not temptable.NoMatch Then
        If local_temp >= temptable("tmin") And local_temp <= temptable("tmax") Then
            answer = True
        Else
            answer = False
        End If
    Else
        ' this shouldn't ever happen, already checked for cas here
        answer = True
    End If
    temptable.Close
    is_valid_loll_range = answer
    
End Function




Public Function load_BIP_data(code As Integer, biparray() As Double) As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim MainGroup As Integer
    Dim DBTbl As Recordset
        
    Select Case code
        Case 1
            'Load AGLB BIP database
            On Error GoTo not_there
            Set DBTbl = chembrowsedb.OpenRecordset("AGLB", dbOpenTable)
            On Error Resume Next
            For i = 1 To 58
                For j = 1 To 58
                    MainGroup = DBTbl("Main Group")
                    biparray(MainGroup, j) = DBTbl(j)
                Next j
                DBTbl.MoveNext
            Next i
            DBTbl.Close
            load_BIP_data = True
        Case 2
            'Load AVLE BIP database
            On Error GoTo not_there
            Set DBTbl = chembrowsedb.OpenRecordset("AVLE", dbOpenTable)
            On Error Resume Next
            For i = 1 To 58
                For j = 1 To 58
                    MainGroup = DBTbl("Main Group")
                    biparray(MainGroup, j) = DBTbl(j)
                Next j
                DBTbl.MoveNext
            Next i
            DBTbl.Close
            load_BIP_data = True
        Case 3
            'Load AENV BIP database
            On Error GoTo not_there
            Set DBTbl = chembrowsedb.OpenRecordset("AENV", dbOpenTable)
            On Error Resume Next
            For i = 1 To 58
                For j = 1 To 58
                    MainGroup = DBTbl("Main Group")
                    biparray(MainGroup, j) = DBTbl(j)
                Next j
                DBTbl.MoveNext
            Next i
            DBTbl.Close
            load_BIP_data = True
        Case Else
            'Load ALLE BIP database
            On Error GoTo not_there
            Set DBTbl = chembrowsedb.OpenRecordset("ALLE", dbOpenTable)
            On Error Resume Next
            For i = 1 To 32
                For j = 1 To 32
                    MainGroup = DBTbl("Main Group")
                    biparray(MainGroup, j) = DBTbl(j)
                Next j
                DBTbl.MoveNext
            Next i
            DBTbl.Close
            load_BIP_data = True
    End Select
    Exit Function
not_there:
    load_BIP_data = False
    
End Function

Public Function load_UNIFAC_data(MGSGarg() As Long, RIarg() As Double, QIarg() As Double, MWSarg() As Double, MVSarg() As Double) As Boolean

    Dim i As Integer
    Dim DBTbl As Recordset
    
     'Load area and group parameters and MW and MV group contributions
    On Error GoTo not_there
    Set DBTbl = chembrowsedb.OpenRecordset("UNIFAC", dbOpenTable)
    On Error Resume Next
    For i = 1 To 116
        MGSGarg(i) = DBTbl("Main Group")
        RIarg(i) = DBTbl("Rk")
        QIarg(i) = DBTbl("Qk")
        MWSarg(i) = DBTbl("MW Group")
        MVSarg(i) = DBTbl("MV Group")
        DBTbl.MoveNext
    Next i
    DBTbl.Close
    load_UNIFAC_data = True
    Exit Function
not_there:
    load_UNIFAC_data = False
End Function

Public Function get_ref_smiles_string(cas_arg As Long) As String

Dim temptable As Recordset
Dim local_smiles As String
Dim found As Boolean

Set temptable = chembrowsedb.OpenRecordset("reference chemicals", dbOpenSnapshot)
temptable.MoveFirst
found = False
local_smiles = "error"
While Not temptable.EOF
    If Trim(temptable("CAS")) = cas_arg Then
        found = True
        local_smiles = temptable("SMILES")
        GoTo finish
    End If
    temptable.MoveNext
Wend
finish:
temptable.Close
If found = True Then
    get_ref_smiles_string = local_smiles
Else
    get_ref_smiles_string = "error"
End If
End Function

'-----------------------------------------------------------
' SUB: Sets Curent DB name and path
'-----------------------------------------------------------
'
Sub Set_CurName(ByVal strCurName As String)
Dim i As Integer
Dim mypos As String
    If InStr(strCurName, "\") <> 0 Then
        mypos = InStr(1, strCurName, "\", 1)
        For i = 0 To 20
            If (InStr(mypos + 1, strCurName, "\", 1) > 0) Then
                mypos = InStr(mypos + 1, strCurName, "\", 1)
            End If
        Next i
        curpath = Left(strCurName, mypos)
        curname = Mid(strCurName, Len(curpath) + 1, Len(strCurName) - Len(curpath))
        Exit Sub
    Else
        For i = 0 To dbman_apps - 1
            If strCurName = dbman_(i, 0) Then
                curname = dbman_(i, 0)
                curpath = dbman_(i, 1)
                Exit Sub
            End If
        Next i
    End If
    MsgBox "Current DB name and path was not set."
End Sub

