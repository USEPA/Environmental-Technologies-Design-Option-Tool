Attribute VB_Name = "modsol"
Attribute VB_Description = "String*"
Option Explicit
Global Const NO_FORM = 20
' forms for Swater
Global Const G_G_FORM = 0
Global Const G_L_FORM = 1
Global Const M_L_FORM = 2
Global Const L_L_FORM = 3
Global Const PPM_WATER_FORM = 4
' forms for Schem
Global Const PPM_CHEM_FORM = 5
' for Henry's Constant conversions
Global Const HC_PRESSURE_FORM = 1
Global Const HC_ULESS_FORM = 2
Global Const HC_MIXED_FORM = 3
Global Const HC_GAS_CONSTANT = 82.06    ' gas constant in m3 atm/mol-K

'Global Const M_L_CHEM_FORM = 6
'Global Const L_L_CHEM_FORM = 7
'Global Const G_G_CHEM_FORM = 8


Global Const WATER_DENSITY = 996.5  ' needs to be in g/L









Public Function do_sol_convert(value_from As Double, value_to As Double, unit_from As String, unit_to As String, whichprop As Integer) As Boolean

    ' FORMS:    1 = g/g
    '           2 = g/L
    '           3 = mol/L
    '           4 = ppm water
    '           5 = ppm chem
    
   
    Dim density1 As Double
    Dim density2 As Double
    Dim mw1 As Double
    Dim mw2 As Double
    Dim temp As Double
    Dim xmass1 As Double
    Dim xmole As Double
    Dim interm_value As Double
    Dim form_from As Integer
    Dim form_to As Integer
    Dim i As Integer
    Dim J As Integer
    Dim success As Boolean
    Dim temp_value As Double
    
    If value_from = 0 Or value_from = ERROR_FLAG Then
        success = False
        do_sol_convert = False
        Exit Function
    End If
    success = True
    ' first invert it and make sure it's in chem/solution form if it's an Swater conversion
    ' do we have to do this if it's in a ppm or ppb form??????
   ' If whichprop = Swater Then
   '     value_from = 1# / (value_from + 1#) ' make it in solution
        
    'End If
    ' get the form this is in and standardize the value to that form
    form_from = find_form_from(value_from, unit_from, whichprop)
    form_to = get_form(unit_to, whichprop)
    If form_from = NO_FORM Or form_to = NO_FORM Then
        MsgBox ("Can't convert " & unit_from)
        value_to = ERROR_FLAG
        do_sol_convert = False
        Exit Function
    End If
    
    ' now we have everything in one of 5 forms, get all the data we might need:
    '       density of chem and water
    '       MW of chem and water
    '       xmass and xmole
    ' density (g/L) and MW (g/mol) of water will be constants for now
    ' first we assume Schem and water is subscript 2, chem is subscript1
    density1 = 0#
    
    If InfoMethod(LD).Enabled(InfoMethod(LD).CurMethod) = False Then
        
        For i = 0 To NumMethods - 1
            density1 = CalcFofT(LD, i)
            If density1 <> 0 Then
                Exit For
            End If
        Next i
        If density1 = 0# Then
            do_sol_convert = False
            Exit Function
        Else
            density1 = Convert(density1, LD, Get_DefaultUnit(LD), "g/L", False)
        End If
    Else
        temp_value = InfoMethod(LD).value(InfoMethod(LD).CurMethod)
        ' this is a temporary hack to get this to work the first time it calcs
        If Trim(InfoMethod(LD).Unit) = "" Then
            density1 = Convert(temp_value, LD, "kmol/m3", "g/L", False)
        Else
            density1 = Convert(temp_value, LD, InfoMethod(LD).Unit, "g/L", False)
        End If
    End If
    
    
    If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = False Then
        mw1 = get_MW("g/mol", Cur_Info.CAS)
    Else
        temp_value = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
        mw1 = Convert(temp_value, MW, InfoMethod(MW).Unit, "g/mol", False)
        
    End If
    mw2 = WATER_MW
    density2 = WATER_DENSITY
    ' now if it's swater need to switch the above values
    If whichprop = Swater Then
        temp = density1
        density1 = density2
        density2 = temp
        temp = mw1
        mw1 = mw2
        mw2 = temp
    End If
    
    ' get the conversion factors xmass and xmole
    If form_from = G_G_FORM Then
        xmass1 = value_from ' its already an xmass
    ElseIf form_from = G_L_FORM Then
        xmass1 = 1# / ((((1# / value_from) - (1# / density1)) * density2) + 1#)
    ElseIf form_from = PPM_CHEM_FORM Then
        xmass1 = value_from / 1000000#
    ElseIf form_from = PPM_WATER_FORM Then
        xmass1 = value_from / 1000000#
    ElseIf form_from = L_L_FORM Then
        xmass1 = value_from * (density1 / WATER_DENSITY)
    ElseIf form_from = M_L_FORM Then
        value_from = value_from * mw1   ' put in g/L form
        form_from = G_L_FORM            ' in case it matters?
        xmass1 = 1# / ((((1# / value_from) - (1# / density1)) * density2) + 1#)
    End If
    xmole = 1# / (1# + ((1# / xmass1 - 1#) * (mw1 / mw2)))
    
    ' now convert to standard form of desired units
    
    If form_to = PPM_CHEM_FORM Then
        value_to = xmass1 * 1000000#
    ElseIf form_to = PPM_WATER_FORM Then
        value_to = xmass1 * 1000000#
    ElseIf form_to = G_G_FORM Then
        value_to = xmass1
    ElseIf form_to = G_L_FORM Then
        value_to = 1# / ((1# / density1) + (((1# / xmass1) - 1#) * (1# / density2)))
    ElseIf form_to = M_L_FORM Then
        value_to = (1# / mw1) * (1# / ((1# / density1) + (((1# / xmass1) - 1#) * (1# / density2))))
    ElseIf form_to = L_L_FORM Then
        value_to = xmass1 * (density2 / density1)
    Else
        MsgBox ("Can't convert " & unit_from & " to " & unit_to)
        value_to = ERROR_FLAG
        success = False
        GoTo end_calc
    End If
    ' convert to the specific unit desired
    success = find_form_to(value_to, unit_to, form_to, whichprop)
    ' undo the inversion if it's an Schem conversion
    'If whichprop = Swater Then
     '   value_to = 1# / value_to
     '   value_to = value_to - 1
    'End If
    
end_calc:
    do_sol_convert = success
End Function



Public Function get_density(unit_desired As String, casno As Long) As Double
Dim dbtable As Recordset
Dim density1 As Double
Dim temp_value As Double
Dim converted_answer As Double
Dim Criteria As String
Dim unit_from As String
Dim i As Integer

If InfoMethod(LD).Enabled(InfoMethod(LD).CurMethod) = False Then
        
        For i = 0 To NumMethods - 1
            density1 = CalcFofT(LD, i)  ' will return it in default units
            If density1 <> 0 Then
                Exit For
            End If
        Next i
        If density1 = 0# Then
            get_density = False
            Exit Function
        Else
'msh    I'm assuming that density1 is not jsut for "viewing" purposes,
'       there was an error in the function call because density1 is being passed
'       through this function (as a double) when it's looking for a boolean.
'
'msh    density1 = Convert(LD, ConvertToDefault(LD), unit_desired, density1)
'msh    note: not sure if I should pass the boolean value as T or F?
'
        density1 = Convert(density1, LD, Get_DefaultUnit(LD), unit_desired, True)
        End If
    Else
        temp_value = InfoMethod(LD).value(InfoMethod(LD).CurMethod)
'msh    density1 = Convert(LD, InfoMethod(LD).Unit, unit_desired, temp_value)
        density1 = Convert(temp_value, LD, InfoMethod(LD).Unit, unit_desired, True)
    End If
    get_density = density1
'get_density = Convert(LD, unit_from, unit_desired, answer)
End Function




Public Function find_form_from(value As Double, unit_string As String, whichprop As Integer) As Integer

' NOTE:
' since this was written to ultimately be integrated in other software,
' in which 'form' will not be a global, we're setting form (global in ucon.vbp)
' and returning it
' The objective here is to return a standard form and do simple
' manipulations on the value to make the given unit conform to the form
Dim Form As Integer
unit_string = Trim(unit_string)
Select Case whichprop
    Case HC:
        If unit_string Like "Pa*mol/mol" Or unit_string Like "Pa" Then
            Form = HC_PRESSURE_FORM
            value = value / 101325#
        ElseIf unit_string Like "kPa*mol/mol" Or unit_string Like "kPa" Then
            Form = HC_PRESSURE_FORM
            value = value / 101.325
        ElseIf unit_string Like "*unit-less*" Then
            Form = HC_ULESS_FORM        ' no change to value
        ElseIf unit_string Like "atm" Or unit_string Like "atm*mol/mol" Then
            Form = HC_PRESSURE_FORM     ' no change to value
        ElseIf unit_string Like "atm*m3/mol" Then
            Form = HC_MIXED_FORM        ' no change to value
        ElseIf unit_string Like "kPa*m3/kmol" Then
            Form = HC_MIXED_FORM
            value = value / 101325#
'msh        ElseIf unit_string Like "kPa*m3/mol" Then
'            Form = HC_MIXED_FORM
'            value = value / 101.325
        ElseIf unit_string Like "MPa*mol/mol" Or unit_string Like "MPa" Then
            Form = HC_PRESSURE_FORM
            value = value / 101325#
            value = value * 1000000#
'msh        ElseIf unit_string Like "bar*mol/mol" Or unit_string Like "bar" Then
'            Form = HC_PRESSURE_FORM
'            value = value / 1.01325
'msh        ElseIf unit_string Like "cm*Hg*mol/mol" Or unit_string Like "cm*Hg" Then
'            Form = HC_PRESSURE_FORM
'            value = value / 76#
'msh        ElseIf unit_string Like "kN/m2*mol/mol" Or unit_string Like "kN/m2" Then
'            Form = HC_PRESSURE_FORM
'            value = value / 101.325
'msh        ElseIf unit_string Like "lb*/in2*mol/mol" Or unit_string Like "lb*in2" Then
'            Form = HC_PRESSURE_FORM
'            value = value / 14.696
'msh        ElseIf unit_string Like "lb*/ft2*mol/mol" Or unit_string Like "lb*ft2" Then
'            Form = HC_PRESSURE_FORM
'            value = value / 2116.224
'msh        ElseIf unit_string Like "mPa*mol/mol" Or unit_string Like "mPa" Then
'            Form = HC_PRESSURE_FORM
'            value = value / 101325000#
'msh 4/12/99
'**********************************************************
'   To be added in next release...
'        ElseIf unit_string Like "atm/M" Then
'
'        ElseIf unit_string Like "mm Hg*mol/mol" Then
'
'        ElseIf unit_string Like "kPa*m3/mol" Then
'
'        ElseIf unit_string Like "bar*m3/mol" Then
'
'        ElseIf unit_string Like "atm*L/mol" Then
'
'        ElseIf unit_string Like "torr*mol/mol" Then
'**********************************************************
        
        Else
            Form = NO_FORM
            value = ERROR_FLAG
        End If
        
    Case Schem, Swater
        If unit_string Like "ppm(wt)*chem*" Then
            Form = PPM_CHEM_FORM
        ElseIf unit_string Like "ppm(wt)*wat*" Then
            Form = PPM_WATER_FORM
        ElseIf unit_string Like "ppm*" And whichprop = Swater Then
            Form = PPM_WATER_FORM
        ElseIf unit_string Like "ppm*" And whichprop = Schem Then
            Form = PPM_CHEM_FORM
        ElseIf unit_string Like "g/L" Then
            Form = G_L_FORM
        ElseIf unit_string Like "mg/L" Then
            Form = G_L_FORM
            value = value / 1000#
        ElseIf unit_string Like "kmol/m3*" Then
            value = value
            'unit_string = "mol/L"
            Form = M_L_FORM
        ElseIf unit_string Like "mol/L*" Then
            Form = M_L_FORM
        ElseIf unit_string Like "mol/dm3*" Then
            ' value stays the same
            'unit_string = "mol/L"
            Form = M_L_FORM
        ElseIf unit_string Like "*mg/kg*" Then
            value = value / 1000000#
            'unit_string = "g/g"
            Form = G_G_FORM
        ElseIf unit_string Like "g/dm3*" Then
            ' value stays the same
            'unit_string = "g/L"
            Form = G_L_FORM
        ElseIf unit_string Like "ppb*wat*" Then
            value = value / 1000#
            'unit_string = "ppm(wt) water"
            Form = PPM_WATER_FORM
        ElseIf unit_string Like "ppb*chem*" Then
            value = value / 1000#
            'unit_string = "ppm(wt) chem"
            Form = PPM_CHEM_FORM
        ElseIf unit_string Like "g/100 cm3*" Then
            value = value * 10000#
            'unit_string = "g/L"
            Form = G_L_FORM
        ElseIf unit_string Like "g/100 mL*" Or unit_string Like "g/100*ml*" Then
            value = value * 10#
            'unit_string = "g/L"
            Form = G_L_FORM
        ElseIf unit_string Like "Molar" Then
            ' value stays the same
            Form = M_L_FORM
        ElseIf unit_string Like "cm3/cm3*" Then
            ' value stays the same
            'unit_string = "L/L"
            Form = L_L_FORM
        ElseIf unit_string Like "*vol%*" Or unit_string Like "*Vol%*" Then
            value = value / 100#
            Form = L_L_FORM
        ElseIf unit_string Like "cm3/L*" Then
            value = value / 1000#
            'unit_string = "L/L"
            Form = L_L_FORM
        ElseIf unit_string Like "cm3/mL*" Then
            ' value stays the same ??
            'unit_string = "L/L"
            Form = L_L_FORM
        ElseIf unit_string Like "mL/L*" Then
            value = value / 1000#
            'unit_string = "L/L"
            Form = L_L_FORM
        ElseIf unit_string Like "cm3/100 cm3*" Then
            value = value / 100#
            'unit_string = "L/L"
            Form = L_L_FORM
        ElseIf unit_string Like "cm3/100 mL*" Then
            value = value / 100#
            'unit_string = "L/L"
            Form = L_L_FORM
        ElseIf unit_string Like "mL/100 mL*" Then
            value = value / 100#
            'unit_string = "L/L"
            Form = L_L_FORM
        ElseIf Trim(unit_string) Like "g/kg*" Then
            value = value / 1000#
            'unit_string = "g/g"
            Form = G_G_FORM
        ElseIf unit_string Like "mass%" Then
            value = value / 100#
            'unit_string = "g/g"
            Form = G_G_FORM
        ElseIf unit_string Like "wt%" Then
            value = value / 100#
            'unit_string = "g/g"
            Form = G_G_FORM
        Else
            Form = NO_FORM
        End If
    Case Else
        Form = NO_FORM
    End Select
    
    find_form_from = Form
End Function



Public Function find_form_to(value As Double, unit_string As String, form_to As Integer, whichprop As Integer) As Boolean
' NOTE:
' the objective here is to get a value in a standard form and convert to a more
' specific unit within the group that that form encompasses
Dim success As Integer
success = True
Select Case whichprop
    Case HC:
        If form_to = HC_PRESSURE_FORM Then
            If unit_string Like "Pa*mol/mol" Or unit_string Like "Pa" Then
                value = value * 101325#
            ElseIf unit_string Like "kPa*mol/mol" Or unit_string Like "kPa" Then
                value = value * 101.325
            ElseIf unit_string Like "atm" Or unit_string Like "atm*mol/mol" Then
                ' no conversion
            ElseIf unit_string Like "MPa*mol/mol" Or unit_string Like "MPa" Then
                value = value * 101325#
                value = value / 1000000#
            ElseIf unit_string Like "bar*mol/mol" Or unit_string Like "bar" Then
                value = value * 1.01325
            ElseIf unit_string Like "cm*Hg*mol/mol" Or unit_string Like "cm*Hg" Then
                value = value * 76#
            ElseIf unit_string Like "kN/m2*mol/mol" Or unit_string Like "kN/m2" Then
                value = value * 101.325
            ElseIf unit_string Like "lb*/in2*mol/mol" Or unit_string Like "lb*in2" Then
                value = value * 14.696
            ElseIf unit_string Like "lb*/ft2*mol/mol" Or unit_string Like "lb*ft2" Then
                value = value * 2116.224
            ElseIf unit_string Like "mPa*mol/mol" Or unit_string Like "mPa" Then
                value = value * 101325000#
            Else
                value = ERROR_FLAG
            End If
        ElseIf form_to = HC_ULESS_FORM Then
            If unit_string Like "*unit-less*" Then
                ' no conversion
            Else
                value = ERROR_FLAG
            End If
        ElseIf form_to = HC_MIXED_FORM Then
            If unit_string Like "atm*m3/mol" Then
                ' no change to value
            ElseIf unit_string Like "kPa*m3/kmol" Then
                value = value * 101325#
            ElseIf unit_string Like "kPa*m3/mol" Then
                value = value * 101.325
            Else
                value = ERROR_FLAG
            End If
        Else
            value = ERROR_FLAG
        End If
            
    Case Schem, Swater
        If form_to = PPM_CHEM_FORM Then
            If unit_string Like "ppm*" Then
                value = value
            ElseIf unit_string Like "ppb*" Then
                value = value * 1000#
            End If
        ElseIf form_to = PPM_WATER_FORM Then
            If unit_string Like "ppm*" Then
                value = value
            ElseIf unit_string Like "ppb*" Then
                value = value * 1000#
            End If
        ElseIf form_to = G_L_FORM Then
            If unit_string Like "g/L" Then
                value = value
            ElseIf unit_string Like "mg/L" Then
                value = value * 1000#
            ElseIf unit_string Like "g/dm3" Then
                value = value
            ElseIf unit_string Like "g/100 cm3" Then
                value = value / 10#
            ElseIf unit_string Like "g/100 mL" Then
                value = value / 10#
            End If
       ElseIf form_to = M_L_FORM Then
            If unit_string Like "kmol/m3" Then
                value = value
            ElseIf unit_string Like "mol/L" Then
                value = value
            ElseIf unit_string Like "mol/dm3" Then
                value = value
            ElseIf unit_string Like "molar" Then
                value = value
            End If
        ElseIf form_to = G_G_FORM Then
            If unit_string Like "mg/kg" Then
                value = value * 1000000#
            ElseIf unit_string Like "*g/kg*" Then
                value = value * 1000#
            ElseIf unit_string Like "mass*" Then
                value = value * 100#
            ElseIf unit_string Like "wt*" Then
                value = value * 100#
            End If
        ElseIf form_to = L_L_FORM Then
            If unit_string Like "cm3/cm3" Then
                value = value
            ElseIf unit_string Like "*vol*" Or unit_string Like "*Vol*" Then
                value = value * 100#
            ElseIf unit_string Like "cm3/L*" Then
                value = value * 1000#
            ElseIf unit_string Like "cm3/mL*" Then
                value = value
            ElseIf unit_string Like "mL/L*" Then
                value = value * 1000#
            ElseIf unit_string Like "cm3/100 cm3*" Then
                value = value * 100#
            ElseIf unit_string Like "cm3/100 mL*" Then
                value = value * 100#
            ElseIf unit_string Like "mL/100 mL*" Then
                value = value * 100#
            End If
        Else
            success = False

        End If
    Case Else
        success = False
    End Select
    find_form_to = success
End Function

Public Function get_form(unit_string As String, whichprop As Integer) As Integer
' NOTE:
' since this was written to ultimately be integrated in other software,
' in which 'form' will not be a global, we're setting form (global in ucon.vbp)
' and returning it

' this simply returns a form,  no value manipulation at all
' called by the conversion routine to find out how to begin phase I of the conversion
' (to get into a standard form)
Dim Form As Integer
Select Case whichprop
    Case HC:
        If unit_string Like "Pa*mol/mol" Or unit_string Like "Pa" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "kPa*mol/mol" Or unit_string Like "kPa" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "*unit-less*" Then
            Form = HC_ULESS_FORM
        ElseIf unit_string Like "atm" Or unit_string Like "atm*mol/mol" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "atm*m3/mol" Then
            Form = HC_MIXED_FORM
        ElseIf unit_string Like "kPa*m3/kmol" Then
            Form = HC_MIXED_FORM
        ElseIf unit_string Like "kPa*m3/mol" Then
            Form = HC_MIXED_FORM
        ElseIf unit_string Like "MPa*mol/mol" Or unit_string Like "MPa" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "bar*mol/mol" Or unit_string Like "bar" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "cm*Hg*mol/mol" Or unit_string Like "cm*Hg" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "kN/m2*mol/mol" Or unit_string Like "kN/m2" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "lb*/in2*mol/mol" Or unit_string Like "lb*in2" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "lb*/ft2*mol/mol" Or unit_string Like "lb*ft2" Then
            Form = HC_PRESSURE_FORM
        ElseIf unit_string Like "mPa*mol/mol" Or unit_string Like "mPa" Then
            Form = HC_PRESSURE_FORM
        Else
            Form = NO_FORM
        End If
    
    Case Schem, Swater
        If unit_string Like "ppm(wt)*chem*" Then
            Form = PPM_CHEM_FORM
        ElseIf unit_string Like "ppm(wt)*wat*" Then
            Form = PPM_WATER_FORM
        ElseIf unit_string Like "ppm*" And whichprop = Swater Then
            Form = PPM_WATER_FORM
        ElseIf unit_string Like "ppm*" And whichprop = Schem Then
            Form = PPM_CHEM_FORM
        ElseIf unit_string Like "g/L" Then
            Form = G_L_FORM
        ElseIf unit_string Like "mg/L" Then
            Form = G_L_FORM
        ElseIf unit_string Like "kmol/m3" Then
            Form = M_L_FORM
        ElseIf unit_string Like "mol/L" Then
            Form = M_L_FORM
        ElseIf unit_string Like "mol/dm3" Then
            Form = M_L_FORM
        ElseIf unit_string Like "*mg/kg*" Then
            Form = G_G_FORM
        ElseIf unit_string Like "g/dm3" Then
            Form = G_L_FORM
        ElseIf unit_string Like "ppb*wat*" Then
            Form = PPM_WATER_FORM
        ElseIf unit_string Like "ppb*chem*" Then
            Form = PPM_CHEM_FORM
        ElseIf unit_string Like "g/100 cm3" Then
            Form = G_L_FORM
        ElseIf unit_string Like "g/100 mL" Then
            Form = G_L_FORM
        ElseIf unit_string Like "Molar" Then
            Form = M_L_FORM
        ElseIf unit_string Like "cm3/cm3" Then
            Form = L_L_FORM
        ElseIf unit_string Like "cm3/L*" Then
            Form = L_L_FORM
        ElseIf unit_string Like "cm3/mL*" Then
            Form = L_L_FORM
        ElseIf unit_string Like "mL/L*" Then
            Form = L_L_FORM
        ElseIf unit_string Like "cm3/100 cm3*" Then
            Form = L_L_FORM
        ElseIf unit_string Like "cm3/100 mL*" Then
            Form = L_L_FORM
        ElseIf unit_string Like "mL/100 mL*" Then
            Form = L_L_FORM
        ElseIf Trim(unit_string) Like "g/kg*" Then
            Form = G_G_FORM
        ElseIf unit_string Like "mass%" Then
            Form = G_G_FORM
        ElseIf unit_string Like "wt%" Then
            Form = G_G_FORM
        ElseIf unit_string Like "*vol*" Or unit_string Like "*Vol*" Then
            Form = L_L_FORM
        Else
            Form = NO_FORM
        End If
    Case Else
        Form = NO_FORM
    End Select
    
    get_form = Form
End Function


Public Function do_HC_convert(value_from As Double, value_to As Double, unit_from As String, unit_to As String) As Boolean

    Dim success As Boolean
    Dim form_from As Integer
    Dim form_to As Integer
    'Dim density1 As Double
    Dim density2 As Double
    Dim temperature As Double
    Dim i As Integer
    success = True
    form_from = find_form_from(value_from, unit_from, HC)
    form_to = get_form(unit_to, HC)
    If form_from = NO_FORM Or form_to = NO_FORM Then
        MsgBox ("Can't convert " & unit_from)
        value_to = ERROR_FLAG
        do_HC_convert = False
        Exit Function
    End If
    
    density2 = Convert(WATER_DENSITY, LD, "g/L", "gmol/L", False)
    ' now get temperature in K
    temperature = Convert(Cur_Info.OpT, OptTemp, Cur_Info.OpTUnit, "K", False)
    
    Select Case form_from
    
        Case HC_PRESSURE_FORM
            Select Case form_to
                Case HC_PRESSURE_FORM
                    value_to = value_from
                Case HC_ULESS_FORM
                    value_to = value_from / (density2 * HC_GAS_CONSTANT * temperature)
                Case HC_MIXED_FORM
                    value_to = value_from / density2
            End Select
        Case HC_ULESS_FORM
            Select Case form_to
                Case HC_PRESSURE_FORM
                    value_to = value_from * density2 * HC_GAS_CONSTANT * temperature
                Case HC_ULESS_FORM
                    value_to = value_from
                Case HC_MIXED_FORM
                    value_to = (value_from * density2 * HC_GAS_CONSTANT) / density2
            End Select
        Case HC_MIXED_FORM
            Select Case form_to
                Case HC_PRESSURE_FORM
                    value_to = value_from * density2
                Case HC_ULESS_FORM
                    value_to = (value_from * density2) / (density2 * HC_GAS_CONSTANT * temperature)
                Case HC_MIXED_FORM
                    value_to = value_from
            End Select
    End Select
    success = find_form_to(value_to, unit_to, form_to, HC)
    do_HC_convert = success
End Function
