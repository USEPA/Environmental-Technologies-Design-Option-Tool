Attribute VB_Name = "modpropcalc"
Option Explicit


Public Sub start_do_ufl()

    Dim FMUDI As Double
    Dim FMUCR As Double
    Dim FMUGP As Double
    Dim init_value As Double
    Dim organic As Boolean
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    If frmeditwizard!optinorganic.Value = True Then
        organic = False
    Else
        organic = True
    End If
    init_value = 0#
    FMUDI = init_value
    FMUCR = init_value
    FMUGP = init_value
    
    ' this fills the form with the proper info for this property
    ' and then calls the function to calculate
    frmviewcalc!frresults.Caption = input_name(UFL) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(UFL, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    Call calc_upper(FMUDI, FMUCR, FMUGP, organic)
    ' fill the labels with whatever results we have
    If FMUDI <> init_value Then
        frmviewcalc!lblvalue(0) = Format(FMUDI, "#.##")
        frmviewcalc!lblUnits(0) = "vol % in air"
        frmviewcalc!ckmethod(0).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
    End If
    If FMUCR <> init_value Then
        frmviewcalc!lblvalue(1) = Format(FMUCR, "#.##")
        frmviewcalc!lblUnits(1) = "vol % in air"
        frmviewcalc!ckmethod(1).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(1) = "na"
    End If
    If FMUGP <> init_value Then
        frmviewcalc!lblvalue(2) = Format(FMUGP, "#.##")
        frmviewcalc!lblUnits(2) = "vol % in air"
        frmviewcalc!ckmethod(2).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(2) = "na"
    End If
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate upper flammability limit values for " & selected_name)
    End If
End Sub

Public Sub start_do_lfl()

    Dim FMLDI As Double
    Dim FMLCR As Double
    Dim FMLGP As Double
    Dim init_value As Double
    Dim organic As Boolean
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    If frmeditwizard!optinorganic.Value = True Then
        organic = False
    Else
        organic = True
    End If
    init_value = 0#
    FMLDI = init_value
    FMLCR = init_value
    FMLGP = init_value
    position = 0
    frmviewcalc!frresults.Caption = input_name(LFL) & " Results"
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(LFL, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    
    Call calc_lower(FMLDI, FMLCR, FMLGP, organic)
    ' fill the labels with whatever results we have
    If FMLDI <> init_value Then
        frmviewcalc!lblvalue(0) = Format(FMLDI, "#.##")
        frmviewcalc!lblUnits(0) = "vol % in air"
        frmviewcalc!ckmethod(0).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
    End If
    If FMLCR <> init_value Then
        frmviewcalc!lblvalue(1) = Format(FMLCR, "#.##")
        frmviewcalc!lblUnits(1) = "vol % in air"
        frmviewcalc!ckmethod(1).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(1) = "na"
    End If
    If FMLGP <> init_value Then
        frmviewcalc!lblvalue(2) = Format(FMLGP, "#.##")
        frmviewcalc!lblUnits(2) = "vol % in air"
        frmviewcalc!ckmethod(2).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(2) = "na"
    End If
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate lower flammability limit for " & selected_name)
    End If
End Sub

Public Sub start_do_fp()

    Dim FPTGP As Double
    Dim init_value As Double
    Dim organic As Boolean
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    If frmeditwizard!optinorganic.Value = True Then
        organic = False
    Else
        organic = True
    End If
    init_value = 0#
    FPTGP = init_value
    
    frmviewcalc!frresults.Caption = input_name(FP) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(FP, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    Call calc_FP(FPTGP, organic)
    ' fill the labels with whatever results we have
    If FPTGP <> init_value And FPTGP > -1E+15 And FPTGP < 1E+15 Then
        frmviewcalc!lblvalue(0) = Format(FPTGP, "#")
        frmviewcalc!lblUnits(0) = "K"
        frmviewcalc!ckmethod(0).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
    End If
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate flashpoint for " & selected_name)
    End If
End Sub

Public Sub start_do_ait()

    Dim AITG1 As Double
    Dim AITGP As Double
    Dim init_value As Double
    Dim organic As Boolean
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    If frmeditwizard!optinorganic.Value = 1 Then
        organic = False
    Else
        organic = True
    End If
    init_value = 0#
    AITG1 = init_value
    AITGP = init_value
    frmviewcalc!frresults.Caption = input_name(AIT) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(AIT, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    Call calc_AIT(AITG1, AITGP, organic)
    ' fill the labels with whatever results we have
    If AITG1 <> init_value And Not AITG1 < 0.000000001 And AITG1 < 10000000000# Then
        frmviewcalc!lblvalue(0) = Format(AITG1, "#")
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblUnits(0) = "K"
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!ckmethod(0).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
    End If
    If AITGP <> init_value And Not AITGP < 0.0000000001 And AITGP < 10000000000# Then
        frmviewcalc!lblvalue(1) = Format(AITGP, "#")
        frmviewcalc!lblvalue(1).Visible = True
        frmviewcalc!lblUnits(1) = "K"
        frmviewcalc!lblUnits(1).Visible = True
        frmviewcalc!ckmethod(1).Value = 1
        success = True
    Else
        frmviewcalc!lblvalue(1) = "na"
    End If
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate autoignition temperature for " & selected_name)
    End If
End Sub

Public Sub start_do_LD()

    Dim LDRogers As Double
    Dim LDTemperature As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    LDRogers = init_value
    frmviewcalc!frresults.Caption = input_name(LD) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(LD, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    LDTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            LDTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    ' this should never be true (unless the user elicitly set it this way) so I won't worry about it too much
    If LDTemperature = -999.9 Then
        LDTemperature = STANDARD_TEMPERATURE
        selected_temperature = STANDARD_TEMPERATURE
    End If
    
    Call calc_LDRogers(LDRogers, LDTemperature)
    ' fill the labels with whatever results we have
    If LDRogers <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(LDRogers, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "g/cm3"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate Liquid Density for " & selected_name)
    End If
    
End Sub

Public Sub calc_LDRogers(return_value As Double, working_temp As Double)

    ' this function needs the molecular weight and
    ' molar volume, so first calc those
    ' this method needs:
    '   1.  groups (broken down from smiles)
    '   2.  table of lookup values for groups
    '   3.  number of rings
    
    ' molecular weight variables
    
    Dim water As Boolean
    Dim i As Integer
    Dim xts(MAX_GROUPS) As Double   ' get from db table
    Dim value_xts As Double         ' the value for current group id
    Dim group_id As Integer
    Dim group_count As Integer
    Dim vtm As Double   ' unadjusted MV
    Dim vbm As Double   ' adjusted MV
    Dim temptable As Recordset
    Dim num_rings As Integer
    ' molecular weight
    Dim fwt As Double
    ' liquid density variables
    Dim PW As Double
    Dim A1 As Double
    Dim A2 As Double
    Dim A3 As Double
    Dim A4 As Double
    Dim A5 As Double
    Dim XAVG As Double
    Dim FAVG As Double
    Dim XN As Double
    Dim FN As Double
    Dim FX As Double
    Dim org_density As Double
    
    If check_inputs(LD, 0) = False Then
        Exit Sub
    End If
    i = 0
    
    num_rings = selected_rings
rings_found:
    ' first calculate the molar volume
    ' initialize xts values
    For i = 0 To MAX_GROUPS - 1
        xts(i) = -1
    Next i
    ' get the xts values from the database table
    Set temptable = chembrowsedb.OpenRecordset("Schroeder Values", dbOpenSnapshot)
    On Error Resume Next
    temptable.MoveFirst
    While Not temptable.EOF
        group_id = temptable("Sub Group")
        xts(group_id) = temptable("Value")
        temptable.MoveNext
    Wend
    temptable.Close
    On Error GoTo error_handler
    ' now calc the molar volume
    For i = 0 To MAX_GROUPS_PER_CHEM - 1
        group_id = cur_chem_groups(i)
        group_count = num_cur_chem_groups(i)
        If group_id <> -1 Then
            value_xts = xts(group_id)
            If value_xts < 0 Then
                ' can't calc this one, return value passed to indicate failure
                Exit Sub
            Else
                vtm = vtm + value_xts * group_count
            End If
        End If
    Next i
    vbm = vtm
    If i = 1 And num_cur_chem_groups(0) = 1 Then
        GoTo after_ring_adjustment
    Else
        vbm = vbm - (num_rings * 7#)
    End If
after_ring_adjustment:
    
    ' now get the molecular weight
    Call calc_MW(fwt)
    ' get the temperature entered or default value
    
    ' now calculate Liquid Density using Molar Volume = vbm
    
    PW = 0.95
    A1 = -1.4176800403
    A2 = 8.976651524
    A3 = -12.275501969
    A4 = 7.4584410413
    A5 = -1.738491605
    XAVG = 324.65
    FAVG = 0.98396
    XN = working_temp / XAVG
    FN = A1 + A2 * XN + A3 * XN ^ 2 + A4 * XN ^ 3 + A5 * XN ^ 4
    FX = FN * FAVG
    org_density = PW * FX * (fwt / vbm) / (18.015 / 21#)
    return_value = org_density
    Exit Sub
error_handler:
    
End Sub

Public Sub start_do_CV()

    Dim CVLydersen As Double
    Dim CVTemperature As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    
    init_value = 0#
    CVLydersen = init_value
    frmviewcalc!frresults.Caption = input_name(CV) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(CV, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) <> "" And Len(Trim(frmeditwizard!tbxinputprop(i))) > 0 And Trim(frmeditwizard!lblinputprop(i).Caption) = "Temperature (C)" Then
            CVTemperature = CDbl(Trim(frmeditwizard!tbxinputprop(i).Text))
        End If
    Next i
    If CVTemperature = -99.9 Then
        CVTemperature = STANDARD_TEMPERATURE
    End If
    ' DENISE WRITE THIS
    Call calc_CVLydersen(CVLydersen, CVTemperature)
    ' fill the labels with whatever results we have
    If CVLydersen <> init_value Then
        frmviewcalc!lblvalue(0) = Format(CVLydersen, "###0.00")
        frmviewcalc!lblUnits(0) = "cm3/mol"
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate Critical Volume for " & selected_name)
    End If
    
End Sub

Public Sub calc_MW(return_value As Double)

    ' this function will calculate the molecular weight of a
    ' compound based on
    '   1. a list of elements and numbers of elements
    '   2. the wt of each element
    
    Dim element_value As Double
    Dim element_weight As Double
    Dim weight As Double
    Dim element As String
    
    Dim i As Integer
    If check_inputs(MW, 0) = False Then
        Exit Sub
    End If
    For i = 0 To MAX_ELEMENTS - 1
        frmeditwizard!grdelements.Row = i + 1
        frmeditwizard!grdelements.Col = 1
        If Trim(frmeditwizard!grdelements.Text) <> "" Then
            element = Trim(frmeditwizard!grdelements.Text)
            element_weight = get_mw_element(element)
            frmeditwizard!grdelements.Col = 2
            weight = CDbl(Trim(frmeditwizard!grdelements.Text)) * element_weight
       End If
    Next i
    return_value = weight
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

Public Sub start_do_MW()

    Dim MWatomic As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    MWatomic = init_value
    frmviewcalc!frresults.Caption = input_name(MW) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(MW, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    
    Call calc_MW(MWatomic)
    ' fill the labels with whatever results we have
    If MWatomic <> init_value Then
        frmviewcalc!lblvalue(0) = Format(MWatomic, "###0.00")
        frmviewcalc!lblUnits(0) = "g"
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate Molecular Weight for " & selected_name)
    End If
    
End Sub

Public Sub calc_CVLydersen(return_value As Double, working_temp As Double)

    Dim temptable As Recordset
    Dim group_mult As Double
    Dim i As Integer
    
    If check_inputs(CV, 0) = False Then
        Exit Sub
    End If
    Set temptable = chembrowsedb.OpenRecordset("Lydersen", dbOpenSnapshot)
    
    On Error Resume Next
    For i = 0 To MAX_GROUPS_PER_CHEM - 1
        If cur_chem_groups(i) = -1 Or num_cur_chem_groups(i) = 0 Then
            GoTo done_loop
        End If
        temptable.MoveFirst
        While Not temptable.EOF
            If temptable("Group ID") = cur_chem_groups(i) Then
                group_mult = temptable("Value")
                GoTo next_loop
            End If
            temptable.MoveNext
        Wend
next_loop:
        return_value = return_value + (group_mult * num_cur_chem_groups(i))
    Next i
done_loop:
temptable.Close
return_value = return_value + 40#

End Sub

Public Sub start_do_HC()

    Dim HCHineMoo As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    HCHineMoo = init_value
    frmviewcalc!frresults.Caption = input_name(HC) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(HC, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    
    
    Call calc_HCHineMookerjee(HCHineMoo)
    ' fill the labels with whatever results we have
    If HCHineMoo <> init_value Then
        frmviewcalc!lblvalue(0) = Format(HCHineMoo, "###0.00")
        frmviewcalc!lblUnits(0) = "unit-less"
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate Henry's Constant for " & selected_name)
    End If
    
End Sub

Public Sub calc_HCHineMookerjee(return_value As Double)

    Dim temptable As Recordset
    Dim group_mult As Double
    Dim inter_value As Double
    Dim i As Integer
    inter_value = 0#
    return_value = 0#
    If check_inputs(HC, 0) = False Then
        Exit Sub
    End If
    Set temptable = chembrowsedb.OpenRecordset("Hine & Mookerjee", dbOpenSnapshot)
    
    On Error Resume Next
    For i = 0 To MAX_GROUPS_PER_CHEM - 1
        If cur_chem_groups(i) = -1 Or num_cur_chem_groups(i) = 0 Then
            GoTo done_loop
        End If
        temptable.MoveFirst
        While Not temptable.EOF
            If temptable("Group ID") = cur_chem_groups(i) Then
                group_mult = temptable("HC Value")
                GoTo next_loop
            End If
            temptable.MoveNext
        Wend
next_loop:
        inter_value = inter_value + (group_mult * num_cur_chem_groups(i))
    Next i
done_loop:
temptable.Close
return_value = 10 ^ inter_value

End Sub

Public Sub start_do_vp()

    Dim VPloll As Double
    Dim VPTemperature As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    VPloll = init_value
    frmviewcalc!frresults.Caption = input_name(LD) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(LD, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(Vp, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    VPTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            VPTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    If VPTemperature = -999.9 Then
        VPTemperature = STANDARD_TEMPERATURE
    End If
    
    Call do_loll_calc(VPloll, VPTemperature)
    ' fill the labels with whatever results we have
    If VPloll <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(VPloll, "###0.00")
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "Pa"
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate Vapor Pressure for " & selected_name)
    End If
    
End Sub


Public Sub do_loll_calc(return_value As Double, temparg As Double)


' The Loll Vapor Pressure Calculation as it's written here needs
' the master database for look up tables for:
'   chemical smiles
'   reference chemical for each chemical family
'   ak and bk values
'   A B and C values
'
' tables added to master.mdb are: "reference chemicals" and "othmer fragments"

    Dim i As Integer
    Dim temp_value As Double    ' the temperature (we made need to convert) needs it in C
    Dim dummyvalue As Double    ' indicates n/a in ak and bk table
    Dim message_answer As Integer
    Dim logchemvp As Double
    Dim logrefvp As Double
    Dim ref_groups_quant() As Long
    Dim ref_groups_id() As Long
    Dim chem_groups_quant() As Long
    Dim chem_groups_id() As Long
    Dim temp(0 To 20) As Long
    Dim result As Byte
    Dim ref_cas As Long
    Dim chem_family_group As String
    Dim ref_smiles As String
    Dim chem_smiles As String
    Dim ref_A As Double
    Dim ref_B As Double
    Dim ref_C As Double
    Dim ref_ak(16) As Double
    Dim ref_bk(16) As Double
    Dim chem_ak(16) As Double
    Dim chem_bk(16) As Double
    Dim sum_ref_ak As Double
    Dim sum_ref_bk As Double
    Dim sum_chem_ak As Double
    Dim sum_chem_bk As Double
    Dim temptable As Recordset
    dummyvalue = -999
    
    ' need this in C, hardcode this in for right now DENISE fix
    temp_value = temparg
    
    
    ' get the family group the chemical belongs to
    If Trim(selected_family) <> "" Then
        chem_family_group = selected_family
    Else
        MsgBox ("can't find family group for " & selected_name)
        Exit Sub
    End If
    
    
    ' get the smiles for the chemical
    chem_smiles = selected_smiles
    chem_family_group = selected_family
    
    ' get the reference chemical for this chemical group
    ref_cas = get_reference_chemical(chem_family_group)
    If ref_cas = 0 Then
        MsgBox ("can't find reference chemical")
        Exit Sub
    End If
    
    ' get the smiles for the reference chemical
    ' DENISE write this code
    ref_smiles = get_ref_smiles_string(ref_cas)
    If Trim(ref_smiles) = "error" Then
        MsgBox ("can't find reference chemical")
        Exit Sub
    End If
    
    ' check that the temperature is within the range for the reference chemical
        If is_valid_loll_range(ref_cas, temp_value, "C") = False Then
            message_answer = MsgBox("temperature is outside of valid range, continue calculation?", vbYesNo)
            If message_answer = vbNo Then
                'do_loll_calc = 0
                Exit Sub
            End If
        End If
    
    ' get the group breakdown for chemical (using shredder)
    'Call do_temp_structure_disassembly("unifac2.dat", chem_smiles, 1, chem_groups_quant, chem_groups_id)
    ReDim chem_groups_id(0 To 99) As Long, chem_groups_quant(0 To 99) As Long
    Call MOSDAP(chem_smiles, 0, AppPath & "unifac.dat", "", 2, result, chem_groups_id(0), chem_groups_quant(0), temp(0), temp(0))

    ' get the group breakdown for reference chemical
    'Call do_temp_structure_disassembly("unifac2.dat", ref_smiles, 1, ref_groups_quant, ref_groups_id)
    ReDim ref_groups_id(0 To 99) As Long, ref_groups_quant(0 To 99) As Long
    Call MOSDAP(ref_smiles, 0, AppPath & "unifac.dat", "", 2, result, ref_groups_id(0), ref_groups_quant(0), temp(0), temp(0))
 
        
    
     ' get the A B and C parameters for the reference chemical
        Set temptable = chembrowsedb.OpenRecordset("reference chemicals", dbOpenSnapshot)
        On Error Resume Next
        temptable.FindFirst "CAS = " & Val(ref_cas)
        If Not temptable.NoMatch Then
            ref_A = temptable("A")
            ref_B = temptable("B")
            ref_C = temptable("C")
        Else
            MsgBox ("unable to find reference chemical")
            GoTo no_calc
        End If
        temptable.Close
       
    ' get the ak and bk fragment values for chemical, -999 is the n/a value
        Set temptable = chembrowsedb.OpenRecordset("othmer fragments", dbOpenSnapshot)
        
        For i = 0 To 15
            If chem_groups_id(i) <> -1 Then
                temptable.FindFirst "family group =" & chem_groups_id(i)
                If Not temptable.NoMatch Then
                    chem_ak(i) = temptable("ak")
                    chem_bk(i) = temptable("bk")
                Else
                    chem_ak(i) = dummyvalue
                    chem_bk(i) = dummyvalue
                End If
            End If
        Next i
        temptable.Close
        
    ' sum up ak and bk for chemical
        sum_chem_ak = 0
        sum_chem_bk = 0
        For i = 0 To 15
            If chem_groups_id(i) <> -1 Then
                If chem_ak(i) <> dummyvalue Then
                    sum_chem_ak = sum_chem_ak + chem_groups_quant(i) * chem_ak(chem_groups_id(i))
                End If
                If chem_bk(i) <> dummyvalue Then
                    sum_chem_bk = sum_chem_bk + chem_groups_quant(i) * chem_bk(chem_groups_id(i))
                End If
            End If
        Next i
        
    ' get the ak and bk fragment values for reference chemical -999 is the n/a value
        Set temptable = chembrowsedb.OpenRecordset("othmer fragments", dbOpenSnapshot)
        
        For i = 0 To 15
            If ref_groups_id(i) <> -1 Then
                temptable.FindFirst "family group =" & ref_groups_id(i)
                If Not temptable.NoMatch Then
                    ref_ak(i) = temptable("ak")
                    ref_bk(i) = temptable("bk")
                Else
                    ref_ak(i) = 0
                    ref_bk(i) = 0
                End If
            End If
        Next i
        temptable.Close
        
    ' sum up ak and bk for reference chemical
        sum_ref_ak = 0
        sum_ref_bk = 0
        For i = 0 To 15
            If ref_groups_id(i) <> -1 Then
                If ref_ak(i) <> dummyvalue Then
                    sum_ref_ak = sum_ref_ak + ref_groups_quant(i) * ref_ak(ref_groups_id(i))
                End If
                If ref_bk(i) <> dummyvalue Then
                    sum_ref_bk = sum_ref_bk + ref_groups_quant(i) * ref_bk(ref_groups_id(i))
                End If
            End If
        Next i
        
    ' calculate the reference vapor pressure logrefvp
    
        logrefvp = sum_ref_ak - (sum_ref_bk / (temp_value + ref_C))
     
    ' do loll equation (model 4 from thesis)
    
        ' check for divide by zero before doing the calculation
        If sum_ref_bk = 0 Then
            
            Exit Sub
        End If
        logchemvp = ((sum_chem_bk / sum_ref_bk) * logrefvp) - (sum_ref_ak * (sum_chem_bk / sum_ref_bk)) + sum_chem_ak
    
    ' return the value
    return_value = logchemvp
    Exit Sub
no_calc:
    
End Sub

Public Sub start_do_schem()
    Dim schemUNIFAC As Double
    Dim schemTemperature As Double
    Dim init_value As Double
    Dim MWatomic As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    schemUNIFAC = init_value
    frmviewcalc!frresults.Caption = input_name(Schem) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(Schem, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(Schem, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    schemTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            schemTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    If schemTemperature = -999.9 Then
        schemTemperature = STANDARD_TEMPERATURE
    End If
    ' get the molecular weight
    MWatomic = 0#
    Call calc_MW(MWatomic)
    
    'Activity coefficient of chemical in water" - "Aqueous Solubility" - "Log10 Kow"
    If Calc_Unifac(schemUNIFAC, "", schemTemperature, "Aqueous Solubility", MWatomic) = False Then
        GoTo Sub_End
    End If
    
    ' fill the labels with whatever results we have
    If schemUNIFAC <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(schemUNIFAC, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "kmol/m3"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
Sub_End:
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate Solubility Chemical in Water for " & selected_name)
    End If
    
    
End Sub

Public Sub start_do_swater()

    Dim swaterUNIFAC As Double
    Dim swaterTemperature As Double
    Dim init_value As Double
    Dim MWatomic As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    swaterUNIFAC = init_value
    frmviewcalc!frresults.Caption = input_name(Swater) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(Swater, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(Swater, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    swaterTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            swaterTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    If swaterTemperature = -999.9 Then
        swaterTemperature = STANDARD_TEMPERATURE
    End If
    ' get the molecular weight
    MWatomic = 0#
    Call calc_MW(MWatomic)
    
    Call CalcSwaterUNIFAC(swaterTemperature, MWatomic, swaterUNIFAC)
    ' fill the labels with whatever results we have
    If swaterUNIFAC <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(swaterUNIFAC, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "kmol/m3"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate Solubility Water in Chemical for " & selected_name)
    End If
    
End Sub

Public Sub start_do_acchem()

    Dim acchemUNIFAC As Double
    Dim acchemTemperature As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    acchemUNIFAC = init_value
    frmviewcalc!frresults.Caption = input_name(ACchem) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(ACchem, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(ACchem, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    acchemTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            acchemTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    If acchemTemperature = -999.9 Then
        acchemTemperature = STANDARD_TEMPERATURE
    End If
    
    'Activity coefficient of chemical in water" - "Aqueous Solubility" - "Log10 Kow"
    If Calc_Unifac(acchemUNIFAC, "", acchemTemperature, "Activity coefficient of chemical in water") = False Then
        GoTo Sub_End
    End If
    
    ' fill the labels with whatever results we have
    If acchemUNIFAC <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(acchemUNIFAC, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "unit-less"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
Sub_End:
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate Activity Coefficient for " & selected_name)
    End If
    
End Sub

Public Sub start_do_acwater()

    Dim acwaterUNIFAC As Double
    Dim acwaterTemperature As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    acwaterUNIFAC = init_value
    frmviewcalc!frresults.Caption = input_name(ACwater) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(ACwater, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(ACwater, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    acwaterTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            acwaterTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    If acwaterTemperature = -999.9 Then
        acwaterTemperature = STANDARD_TEMPERATURE
    End If
    
    
    Call CalcACwaterUNIFAC(acwaterTemperature, acwaterUNIFAC)
    ' fill the labels with whatever results we have
    If acwaterUNIFAC <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(acwaterUNIFAC, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "unit-less"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate " & input_name(ACwater) & " for " & selected_name)
    End If
    
End Sub

Public Sub start_do_logKow()

    Dim logkowUNIFAC As Double
    Dim logkowTemperature As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    logkowUNIFAC = init_value
    frmviewcalc!frresults.Caption = input_name(logKow) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(logKow, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(logKow, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    logkowTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            logkowTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    If logkowTemperature = -999.9 Then
        logkowTemperature = STANDARD_TEMPERATURE
    End If
    
    'Activity coefficient of chemical in water" - "Aqueous Solubility" - "Log10 Kow"
    If Calc_Unifac(logkowUNIFAC, "", logkowTemperature, "Log10 Kow") = False Then
        GoTo Sub_End
    End If
    
    ' fill the labels with whatever results we have
    If logkowUNIFAC <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(logkowUNIFAC, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "unit-less"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If

Sub_End:
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate " & input_name(logKow) & " for " & selected_name)
    End If
End Sub

Public Sub start_do_logkoc()

    Dim logkocbaker As Double
    Dim logkowvalue As Double
    Dim logkocTemperature As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    logkocbaker = init_value
    logkowvalue = init_value
    frmviewcalc!frresults.Caption = input_name(logKoc) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(logKoc, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(logKoc, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    logkocTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            logkocTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    If logkocTemperature = -999.9 Then
        logkocTemperature = STANDARD_TEMPERATURE
    End If
    
    'Activity coefficient of chemical in water" - "Aqueous Solubility" - "Log10 Kow"
    If Calc_Unifac(logkowvalue, "", logkocTemperature, "Log10 Kow") = False Then
        GoTo Sub_End
    End If
    
    If logkowvalue <> init_value Then
        Call CalclogKocBaker(logkowvalue, logkocbaker)
    End If
    ' fill the labels with whatever results we have
    If logkocbaker <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(logkocbaker, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "cm3/g"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If

Sub_End:
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate " & input_name(logKoc) & " for " & selected_name)
    End If
End Sub

Public Sub calc_ThODcarbBaker(Value As Double)

'Calculation for Carbonaceous ThOD (unit-less)
'
'Method: Baker (1994)
'
'Equation Inputs: Formula
'
    
    Dim i As Integer
    Dim CNT As Integer
    Dim NumAtoms(10) As Integer
    Dim Atom(10) As String
    
    Dim cnum As Integer
    Dim HNum As Integer
    Dim BrNum As Integer
    Dim ClNum As Integer
    Dim FNum As Integer
    Dim INum As Integer
    Dim SiNum As Integer
    Dim NNum As Integer
    Dim SNum As Integer
    Dim ONum As Integer
    Dim PNum As Integer
    Dim Tmp1 As String
    Dim Tmp2 As String
    Dim method_units As String
    Dim Formula As String
    
    Formula = selected_structure
    method_units = "unit-less"
    On Error Resume Next
    
    If Formula = "" Then Exit Sub

    Formula = Trim(Formula) + " "
      
    'Reset number of atoms
    cnum = 0
    HNum = 0
    BrNum = 0
    ClNum = 0
    FNum = 0
    INum = 0
    SiNum = 0
    NNum = 0
    SNum = 0
    ONum = 0
    PNum = 0

    'Break up formula
    i = 1
    CNT = 1
    Do While i < Len(Formula)
        Tmp1 = Mid(Formula, i, 1)
        Tmp2 = Mid(Formula, i + 1, 1)
        If Asc(Tmp1) < 58 And Asc(Tmp1) > 47 Then
            If Asc(Tmp2) < 58 And Asc(Tmp2) > 47 Then
                NumAtoms(CNT) = Val(Tmp1 + Tmp2)
                CNT = CNT + 1
                i = i + 2
            Else
                NumAtoms(CNT) = Val(Tmp1)
                CNT = CNT + 1
                i = i + 1
            End If
        Else
            If Asc(Tmp2) < 123 And Asc(Tmp2) > 96 Then
                Atom(CNT) = Tmp1 + Tmp2
                i = i + 2
            Else
                NumAtoms(CNT) = 1
                Atom(CNT) = Tmp1
                CNT = CNT + 1
                i = i + 1
            End If
        End If
    Loop
            
    CNT = CNT - 1
    
    'Find number of each atom
    For i = 1 To CNT
        Select Case Atom(i)
            Case "C"
                cnum = cnum + NumAtoms(i)
            Case "H"
                HNum = HNum + NumAtoms(i)
            Case "Br"
                BrNum = BrNum + NumAtoms(i)
            Case "Cl"
                ClNum = ClNum + NumAtoms(i)
            Case "F"
                FNum = FNum + NumAtoms(i)
            Case "I"
                INum = INum + NumAtoms(i)
            Case "N"
                NNum = NNum + NumAtoms(i)
            Case "Si"
                SiNum = SiNum + NumAtoms(i)
            Case "S"
                SNum = SNum + NumAtoms(i)
            Case "O"
                ONum = ONum + NumAtoms(i)
            Case "P"
                PNum = PNum + NumAtoms(i)
        End Select
    Next i
    
    'Calculate carbonaceous ThOD
    Value = cnum + ((HNum - (BrNum + FNum + ClNum + INum) - (3 * NNum) - (2 * SNum) - (3 * PNum)) / 4) - (ONum / 2) + (2 * SNum) + (2 * PNum)
    
    'make sure answer is in correct units
    'If Trim(method_units) = Trim(default_units) Then
    '    CalcThODcarbBaker = value
    'Else
    '    CalcThODcarbBaker = Convert(ThODcarb, method_units, default_units, value)
    'End If


End Sub

Public Sub calc_ThODcombBaker(Value As Double)

'Calculation for Combined ThOD (unit-less)
'
'Method: Baker (1994)
'
'Equation Inputs: Formula
'
' Modified 6/19/97 BGH: Added unit parameters
    Dim i As Integer
    Dim CNT As Integer
    Dim NumAtoms(10) As Integer
    Dim Atom(10) As String
    
    Dim cnum As Integer
    Dim HNum As Integer
    Dim BrNum As Integer
    Dim ClNum As Integer
    Dim FNum As Integer
    Dim INum As Integer
    Dim SiNum As Integer
    Dim NNum As Integer
    Dim SNum As Integer
    Dim ONum As Integer
    Dim PNum As Integer
    Dim Tmp1 As String
    Dim Tmp2 As String
    Dim method_units As String
    Dim Formula As String
    
    method_units = "unit-less"
    On Error Resume Next
    
    Formula = selected_structure
    
    If Formula = "" Then Exit Sub

    Formula = Trim(Formula) + " "
      
    'Reset number of atoms
    cnum = 0
    HNum = 0
    BrNum = 0
    ClNum = 0
    FNum = 0
    INum = 0
    SiNum = 0
    NNum = 0
    SNum = 0
    ONum = 0
    PNum = 0

    'Break up formula
    i = 1
    CNT = 1
    Do While i < Len(Formula)
        Tmp1 = Mid(Formula, i, 1)
        Tmp2 = Mid(Formula, i + 1, 1)
        If Asc(Tmp1) < 58 And Asc(Tmp1) > 47 Then
            If Asc(Tmp2) < 58 And Asc(Tmp2) > 47 Then
                NumAtoms(CNT) = Val(Tmp1 + Tmp2)
                CNT = CNT + 1
                i = i + 2
            Else
                NumAtoms(CNT) = Val(Tmp1)
                CNT = CNT + 1
                i = i + 1
            End If
        Else
            If Asc(Tmp2) < 123 And Asc(Tmp2) > 96 Then
                Atom(CNT) = Tmp1 + Tmp2
                i = i + 2
            Else
                NumAtoms(CNT) = 1
                Atom(CNT) = Tmp1
                CNT = CNT + 1
                i = i + 1
            End If
        End If
    Loop
            
    CNT = CNT - 1
    
    'Find number of each atom
    For i = 1 To CNT
        Select Case Atom(i)
            Case "C"
                cnum = cnum + NumAtoms(i)
            Case "H"
                HNum = HNum + NumAtoms(i)
            Case "Br"
                BrNum = BrNum + NumAtoms(i)
            Case "Cl"
                ClNum = ClNum + NumAtoms(i)
            Case "F"
                FNum = FNum + NumAtoms(i)
            Case "I"
                INum = INum + NumAtoms(i)
            Case "N"
                NNum = NNum + NumAtoms(i)
            Case "Si"
                SiNum = SiNum + NumAtoms(i)
            Case "S"
                SNum = SNum + NumAtoms(i)
            Case "O"
                ONum = ONum + NumAtoms(i)
            Case "P"
                PNum = PNum + NumAtoms(i)
        End Select
    Next i
    
    'Calculate carbonaceous ThOD
    Value = cnum + ((HNum - (BrNum + FNum + ClNum + INum) - NNum - (2 * SNum) - (3 * PNum)) / 4) - (ONum / 2) + ((3 * NNum) / 2) + (2 * SNum) + (2 * PNum)
    
    'check to make sure answer is in correct units
    'If Trim(method_units) = Trim(default_units) Then
    '    CalcThODcombBaker = VALUE
    'Else
    '    CalcThODcombBaker = Convert(ThODcomb, method_units, default_units, VALUE)
    'End If
End Sub

Public Sub start_do_ThODcomb()

    Dim thodcombBaker As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    thodcombBaker = init_value
    frmviewcalc!frresults.Caption = input_name(ThODcomb) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(ThODcomb, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(ThODcomb, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    
    Call calc_ThODcombBaker(thodcombBaker)
    ' fill the labels with whatever results we have
    If thodcombBaker <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(thodcombBaker, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "unit-less"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate " & input_name(ThODcomb) & " for " & selected_name)
    End If
End Sub

Public Sub start_do_ThODcarb()

    Dim thodcarbBaker As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    thodcarbBaker = init_value
    frmviewcalc!frresults.Caption = input_name(ThODcarb) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(ThODcarb, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(ThODcarb, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    
    Call calc_ThODcarbBaker(thodcarbBaker)
    ' fill the labels with whatever results we have
    If thodcarbBaker <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(thodcarbBaker, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "unit-less"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate " & input_name(ThODcarb) & " for " & selected_name)
    End If
    
End Sub

Public Sub start_do_BCF()

    Dim logkowUNIFAC As Double
    Dim logkowTemperature As Double
    Dim bcfKobayshi As Double
    Dim init_value As Double
    Dim success As Boolean
    Dim position As Integer
    Dim i As Integer
    success = False
    
    init_value = 0#
    logkowUNIFAC = init_value
    bcfKobayshi = init_value
    frmviewcalc!frresults.Caption = input_name(BCF) & " Results"
    position = 0
    For i = 0 To MAX_METHODS_EACH - 1
        If Trim(wiz_methods(logKow, i)) <> "" Then
            frmviewcalc!lblMethod(position).Visible = True
            frmviewcalc!lblMethod(position).Caption = wiz_methods(BCF, i)
            frmviewcalc!lblvalue(position).Visible = True
            frmviewcalc!lblvalue(position).Caption = ""
            frmviewcalc!lblUnits(position).Visible = True
            frmviewcalc!lblUnits(position).Caption = ""
            frmviewcalc!ckmethod(position).Visible = True
            frmviewcalc!ckmethod(position).Value = False
            position = position + 1
        End If
    Next i
    logkowTemperature = -999.9
    For i = 0 To MAX_INPUTS_EACH - 1
        If Trim(frmeditwizard!lblinputprop(i).Caption) = input_name(CONST_TEMP) And Len(Trim(frmeditwizard!tbxinputprop(i).Text)) > 0 Then
            logkowTemperature = CDbl(frmeditwizard!tbxinputprop(i).Text)
            Exit For
        End If
    Next i
    If logkowTemperature = -999.9 Then
        logkowTemperature = STANDARD_TEMPERATURE
    End If
    
    'Activity coefficient of chemical in water" - "Aqueous Solubility" - "Log10 Kow"
    If Calc_Unifac(logkowUNIFAC, "", logkowTemperature, "Log10 Kow") = False Then
        GoTo Sub_End
    End If
    
    If logkowUNIFAC <> init_value Then
        Call CalcBCFKobayshi(logkowUNIFAC, bcfKobayshi)
    End If
    ' fill the labels with whatever results we have
    If bcfKobayshi <> init_value Then
        frmviewcalc!lblvalue(0).Visible = True
        frmviewcalc!lblvalue(0) = Format(bcfKobayshi, "###0.00")
        frmviewcalc!lblUnits(0).Visible = True
        frmviewcalc!lblUnits(0) = "unit-less"
        frmviewcalc!ckmethod(0).Visible = True
        frmviewcalc!ckmethod(0).Value = CHECKED
        success = True
    Else
        frmviewcalc!lblvalue(0) = "na"
        frmviewcalc!ckmethod(0).Value = UNCHECKED
    End If
    
Sub_End:
    ' if at least one method was calculated then show the form
    If success = True Then
        frmviewcalc.Show 1
    Else
        MsgBox ("unable to calculate " & input_name(BCF) & " for " & selected_name)
    End If
End Sub
