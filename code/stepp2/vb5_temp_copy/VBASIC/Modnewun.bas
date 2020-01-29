Attribute VB_Name = "modnewunits"
'************************************************************************
'mrt 11/8/98 - fixed unit conversions (I have copy of old module)
'msh 2/9/99  - in FofT unit conversions below, database reads in paradox
'              multplying sign as a dot("·") rather than a ("*")...fixed
'              case statements accordingly. NOTE: search for msh to find
'              accurances.
'************************************************************************

'holds what is used for display in other units
Global Cur_Disp As CurInfoType

'Define storage for each method
Global DispMethod(NumProperties) As MethodInfoType
Public Function Toxicity_To_Standard(value As Double, Units_To As String) As Double
' This function converts Toxicity values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    Toxicity_To_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Toxic_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "mg/L"
        result = value
    Case Else
        result = ERROR_FLAG
End Select

Toxicity_To_Standard = result
Exit Function

err_Toxic_To_Standard:
    result = ERROR_FLAG

End Function
Public Function Toxicity_From_Standard(value As Double, Units_To As String) As Double
' This function converts toxicity values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String
 
'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
'msh    Toxicity_To_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Toxic_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "mg/L"
        result = value
    Case Else
        result = ERROR_FLAG
End Select

Toxicity_From_Standard = result
Exit Function

err_Toxic_From_Standard:
    result = ERROR_FLAG
End Function

Public Sub update_DisplayData()
    Dim i As Integer
    Cur_Disp = Cur_Info
    
    For i = 0 To NumProperties
        DispMethod(i) = InfoMethod(i)
    Next
    
    Cur_Info.OpT = Convert(Cur_Info.OpT, OptTemp, Cur_Info.OpTUnit, "K", False)
    Cur_Info.OpTUnit = "K"
    Cur_Info.OpP = Convert(Cur_Info.OpP, OptPress, Cur_Info.OpPUnit, "Pa", False)
    Cur_Info.OpPUnit = "Pa"

    Cur_Disp.OpT = Convert(Cur_Info.OpT, OptTemp, Cur_Info.OpTUnit, "C", False)
    Cur_Disp.OpTUnit = "C"
    
End Sub

Function Get_DefaultUnit(Code As Integer) As String
    
    Select Case Code
            Case OptTemp
                Get_DefaultUnit = "K"
            Case OptPress
                Get_DefaultUnit = "Pa"
            Case MW
                Get_DefaultUnit = "kg/kmol"
            Case LD25
                Get_DefaultUnit = "kg/m3"
            Case LD
                Get_DefaultUnit = "kmol/m3"
            Case mp
                Get_DefaultUnit = "K"
            Case NBP
                Get_DefaultUnit = "K"
            Case VP25
                Get_DefaultUnit = "Pa"
            Case VP
                Get_DefaultUnit = "Pa"
            Case hfor
                Get_DefaultUnit = "J/kmol"
            Case LHC
                Get_DefaultUnit = "J/kmol*K"
            Case VHC
                Get_DefaultUnit = "J/kmol*K"
            Case Hvap25
                Get_DefaultUnit = "J/kmol"
            Case HvapNBP
                Get_DefaultUnit = "J/kmol"
            Case Hvap
                Get_DefaultUnit = "J/kmol"
            Case CT
                Get_DefaultUnit = "K"
            Case CP
                Get_DefaultUnit = "Pa"
            Case Dwater
                Get_DefaultUnit = "cm2/s"
            Case Dair
                Get_DefaultUnit = "cm2/s"
            Case ST25
                Get_DefaultUnit = "N/m"
            Case ST
                Get_DefaultUnit = "N/m"
            Case VV
                Get_DefaultUnit = "Pa*s"
            Case LV
                Get_DefaultUnit = "Pa*s"
            Case LTC
                Get_DefaultUnit = "W/m*K"
            Case VTC
                Get_DefaultUnit = "W/m*K"
            Case UFL
                Get_DefaultUnit = "vol% in air"
            Case LFL
                Get_DefaultUnit = "vol% in air"
            Case FP
                Get_DefaultUnit = "K"
            Case AIT
                Get_DefaultUnit = "K"
            Case Hcomb
                Get_DefaultUnit = "J/kmol"
            Case ThODcarb
                Get_DefaultUnit = "g O2/g chem"
            Case ThODcomb
                Get_DefaultUnit = "g O2/g chem"
            Case COD
                Get_DefaultUnit = "g O2/g chem"
            Case BOD
                Get_DefaultUnit = "g O2/g chem"
            Case ACwater
                Get_DefaultUnit = "unit-less"
            Case HC
                Get_DefaultUnit = "kPa*mol/mol"
            Case ACchem
                Get_DefaultUnit = "unit-less"
            Case logKow
                Get_DefaultUnit = "unit-less"
            Case logKoc
                Get_DefaultUnit = "cm3/g OC"
            Case BCF
                Get_DefaultUnit = "unit-less"
            Case CV
                Get_DefaultUnit = "m3/kmol"
            Case Schem
                Get_DefaultUnit = "ppm(wt)"
            Case Swater
                Get_DefaultUnit = "ppm(wt)"
            Case Fat48E
                Get_DefaultUnit = "mg/L"
            Case Fat96E
                Get_DefaultUnit = "mg/L"
            Case Fat24L
                Get_DefaultUnit = "mg/L"
            Case Fat48L
                Get_DefaultUnit = "mg/L"
            Case Fat96L
                Get_DefaultUnit = "mg/L"
            Case Sal24L
                Get_DefaultUnit = "mg/L"
            Case Sal48L
                Get_DefaultUnit = "mg/L"
            Case Sal96L
                Get_DefaultUnit = "mg/L"
            Case Daph24E
                Get_DefaultUnit = "mg/L"
            Case Daph48E
                Get_DefaultUnit = "mg/L"
            Case Daph24L
                Get_DefaultUnit = "mg/L"
            Case Daph48L
                Get_DefaultUnit = "mg/L"
            Case Mysid96L
                Get_DefaultUnit = "mg/L"
            Case AltSpecies
                Get_DefaultUnit = "mg/L"
        End Select
        
End Function

Function Convert(value As Double, Code As Integer, ConvertFrom As String, ConvertTo As String, External As Boolean) As Double
'handles ALL unit conversion from one unit to another for ALL properties

'NOTE: External is used to determine if the conversion is just for external viewing
' purposes. if so then we want to convert the cur_info data to desired value. IF
' the conversion is for internal purposes ie doing calculations then External should
' should be false so the function does not use the cur_info struct

Dim FromStan As Boolean
Dim ToStan As Boolean
Dim answer As Double

'determine if were coming or going to internal standard
FromStan = IsDefault(ConvertFrom, Code)
ToStan = IsDefault(ConvertTo, Code)

'store in temp variable
answer = value

On Error GoTo err_convert

'chose which Property to convert
Select Case (Code)
    Case OptTemp    'Operating Tempreture
        If (External) Then
            answer = Temp_From_Standard(Cur_Info.OpT, ConvertTo)
        Else
            If Not (FromStan) Then answer = Temp_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Temp_From_Standard(answer, ConvertTo)
        End If
        
    Case OptPress   'Operating Pressure
        If (External) Then
            answer = Press_To_Standard(Cur_Info.OpP, ConvertTo)
        Else
            If Not (FromStan) Then answer = Press_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Press_From_Standard(answer, ConvertTo)
        End If
    
    Case MW         'Molecular Weight
        If (External) Then
            answer = MW_To_Standard(InfoMethod(MW).value(InfoMethod(MW).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = MW_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = MW_From_Standard(answer, ConvertTo)
        End If
    
    Case LD25       'Liquid Density @ 25C
        If (External) Then
            answer = LD25_To_Standard(InfoMethod(LD25).value(InfoMethod(LD25).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = LD25_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = LD25_From_Standard(answer, ConvertTo)
        End If
    
    Case LD         'Liquid Density as f(T)
        If (External) Then
            answer = LD_To_Standard(InfoMethod(LD).value(InfoMethod(LD).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = LD_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = LD_From_Standard(answer, ConvertTo)
        End If
        
    Case mp         'Melting Point
        If (External) Then
            answer = Temp_From_Standard(InfoMethod(mp).value(InfoMethod(mp).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Temp_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Temp_From_Standard(answer, ConvertTo)
        End If
        
    Case NBP        'Normal Boiling Point
        If (External) Then
            answer = Temp_From_Standard(InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Temp_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Temp_From_Standard(answer, ConvertTo)
        End If
        
    Case VP25       'Vapor Pressure @ 25C
        If (External) Then
            answer = Press_To_Standard(InfoMethod(VP25).value(InfoMethod(VP25).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Press_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Press_From_Standard(answer, ConvertTo)
        End If
    
    Case VP         'Vapor Pressure as f(T)
        If (External) Then
            answer = Press_To_Standard(InfoMethod(VP).value(InfoMethod(VP).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Press_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Press_From_Standard(answer, ConvertTo)
        End If
    
    Case hfor       'Heat of Formation
        If (External) Then
            answer = Heat_To_Standard(InfoMethod(hfor).value(InfoMethod(hfor).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Heat_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Heat_From_Standard(answer, ConvertTo)
        End If
    
    Case LHC        'Liquid Heat Capacity as f(T)
        If (External) Then
            answer = HeatCapacity_To_Standard(InfoMethod(LHC).value(InfoMethod(LHC).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = HeatCapacity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = HeatCapacity_From_Standard(answer, ConvertTo)
        End If
    
    Case VHC        'Vapor Heat Capacity as f(T)
        If (External) Then
            answer = HeatCapacity_To_Standard(InfoMethod(VHC).value(InfoMethod(VHC).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = HeatCapacity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = HeatCapacity_From_Standard(answer, ConvertTo)
        End If
    
    Case Hvap25     'Heat of Vaporization @ 25C
        If (External) Then
            answer = Heat_To_Standard(InfoMethod(Hvap25).value(InfoMethod(Hvap25).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Heat_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Heat_From_Standard(answer, ConvertTo)
        End If
    
    Case HvapNBP    'Heat of Vaporization @ NBP
        If (External) Then
            answer = Heat_To_Standard(InfoMethod(HvapNBP).value(InfoMethod(HvapNBP).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Heat_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Heat_From_Standard(answer, ConvertTo)
        End If
    
    Case Hvap       'Heat of Vaporization as f(T)
        If (External) Then
            answer = Heat_To_Standard(InfoMethod(Hvap).value(InfoMethod(Hvap).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Heat_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Heat_From_Standard(answer, ConvertTo)
        End If
    
    Case CT         'Critical Temperature
        If (External) Then
            answer = Temp_From_Standard(InfoMethod(CT).value(InfoMethod(CT).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Temp_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Temp_From_Standard(answer, ConvertTo)
        End If
        
    Case CP         'Critical Pressure
        If (External) Then
            answer = Press_To_Standard(InfoMethod(CP).value(InfoMethod(CP).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Press_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Press_From_Standard(answer, ConvertTo)
        End If
    
    Case CV         'Critical Volume
        If (External) Then
            answer = CV_To_Standard(InfoMethod(CV).value(InfoMethod(CV).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = CV_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = CV_From_Standard(answer, ConvertTo)
        End If
    
    Case Dwater     'Diffusivity in Water
        If (External) Then
            answer = Diffusivity_To_Standard(InfoMethod(Dwater).value(InfoMethod(Dwater).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Diffusivity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Diffusivity_From_Standard(answer, ConvertTo)
        End If
    
    Case Dair       'Diffusivity in Air
        If (External) Then
            answer = Diffusivity_To_Standard(InfoMethod(Dair).value(InfoMethod(Dair).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Diffusivity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Diffusivity_From_Standard(answer, ConvertTo)
        End If
    
    Case ST25       'Surface Tension @ 25C
        If (External) Then
            answer = ST_To_Standard(InfoMethod(ST25).value(InfoMethod(ST25).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = ST_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = ST_From_Standard(answer, ConvertTo)
        End If
    
    Case ST         'Surface Tension as f(T)
        If (External) Then
            answer = ST_To_Standard(InfoMethod(ST).value(InfoMethod(ST).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = ST_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = ST_From_Standard(answer, ConvertTo)
        End If
    
    Case VV         'Vapor Viscosity as f(T)
        If (External) Then
            answer = Viscosity_To_Standard(InfoMethod(VV).value(InfoMethod(VV).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Viscosity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Viscosity_From_Standard(answer, ConvertTo)
        End If
    
    Case LV         'Liquid Viscosity as f(T)
        If (External) Then
            answer = Viscosity_To_Standard(InfoMethod(LV).value(InfoMethod(LV).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Viscosity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Viscosity_From_Standard(answer, ConvertTo)
        End If
    
    Case LTC        'Liquid Thermal Conductivity as f(T)
        If (External) Then
            answer = TC_To_Standard(InfoMethod(LTC).value(InfoMethod(LTC).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = TC_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = TC_From_Standard(answer, ConvertTo)
        End If
    
    Case VTC        'Vapor Thermal Conductivity as f(T)
        If (External) Then
            answer = TC_To_Standard(InfoMethod(VTC).value(InfoMethod(VTC).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = TC_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = TC_From_Standard(answer, ConvertTo)
        End If
    
    Case UFL        'Upper Flammability Limit
        If (External) Then
            answer = FL_To_Standard(InfoMethod(UFL).value(InfoMethod(UFL).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = FL_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = FL_From_Standard(answer, ConvertTo)
        End If
    
    Case LFL        'Lower Flammability Limit
        If (External) Then
            answer = FL_To_Standard(InfoMethod(LFL).value(InfoMethod(LFL).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = FL_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = FL_From_Standard(answer, ConvertTo)
        End If
    
    Case FP         'Flash Point
        If (External) Then
            answer = Temp_From_Standard(InfoMethod(FP).value(InfoMethod(FP).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Temp_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Temp_From_Standard(answer, ConvertTo)
        End If
        
    Case AIT        'Autoignition Temperature
        If (External) Then
            answer = Temp_From_Standard(InfoMethod(AIT).value(InfoMethod(AIT).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Temp_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Temp_From_Standard(answer, ConvertTo)
        End If
        
    Case Hcomb      'Heat of Combustion
        If (External) Then
            answer = Heat_To_Standard(InfoMethod(Hcomb).value(InfoMethod(Hcomb).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Heat_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Heat_From_Standard(answer, ConvertTo)
        End If
    
    Case ThODcarb   'Carbonaceous ThOD
        If (External) Then
            answer = Odemand_To_Standard(InfoMethod(ThODcarb).value(InfoMethod(ThODcarb).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Odemand_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Odemand_From_Standard(answer, ConvertTo)
        End If
    
    Case ThODcomb   'Combined ThOD
        If (External) Then
            answer = Odemand_To_Standard(InfoMethod(ThODcomb).value(InfoMethod(ThODcomb).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Odemand_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Odemand_From_Standard(answer, ConvertTo)
        End If
    
    Case COD        'Chemical Oxygen Demand
        If (External) Then
            answer = Odemand_To_Standard(InfoMethod(COD).value(InfoMethod(COD).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Odemand_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Odemand_From_Standard(answer, ConvertTo)
        End If
    
    Case BOD        'Biochemical Oxygen Demand
        If (External) Then
            answer = Odemand_To_Standard(InfoMethod(BOD).value(InfoMethod(BOD).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Odemand_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Odemand_From_Standard(answer, ConvertTo)
        End If
    
    Case ACwater    'Infinite Dilution Activity Coefficient of Water in Chemical
        If (External) Then
            answer = NoUnits_To_Standard(InfoMethod(ACwater).value(InfoMethod(ACwater).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = NoUnits_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = NoUnits_From_Standard(answer, ConvertTo)
        End If
    
    Case HC         'Henry's Constant
            found = do_HC_convert(value, answer, ConvertFrom, ConvertTo)
            If found <> False Then
                Convert = answer
            Else
                Convert = ERROR_FLAG
            End If
            Exit Function

    Case ACchem     'Infinite Dilution Activity Coefficient of Chemical in Water
        If (External) Then
            answer = NoUnits_To_Standard(InfoMethod(ACchem).value(InfoMethod(ACchem).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = NoUnits_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = NoUnits_From_Standard(answer, ConvertTo)
        End If
    
    Case logKow     'log Kow
        If (External) Then
            answer = NoUnits_To_Standard(InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = NoUnits_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = NoUnits_From_Standard(answer, ConvertTo)
        End If
    
    Case logKoc     'log Koc
        If (External) Then
            answer = LogKOC_To_Standard(InfoMethod(logKoc).value(InfoMethod(logKoc).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = LogKOC_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = LogKOC_From_Standard(answer, ConvertTo)
        End If
    
    Case BCF        'Bioconcentration Factor
        If (External) Then
            answer = NoUnits_To_Standard(InfoMethod(BCF).value(InfoMethod(BCF).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = NoUnits_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = NoUnits_From_Standard(answer, ConvertTo)
        End If
    
    Case Schem      'Solubility Limit of Chemical in Water
            found = do_sol_convert(value, answer, ConvertFrom, ConvertTo, Schem)
            If found <> False Then
              Convert = answer
            Else
               Convert = ERROR_FLAG
           End If
           Exit Function
           
    Case Swater     'Solubility Limit of Water in Chemical
            found = do_sol_convert(value, answer, ConvertFrom, ConvertTo, Swater)
            If found <> False Then
              Convert = answer
            Else
               Convert = ERROR_FLAG
           End If
           Exit Function
           
   Case Fat48E     'Fathead minnow 48h, ec50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Fat48E).value(InfoMethod(Fat48E).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
           
   Case Fat96E     'Fathead minnow 96h, ec50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Fat96E).value(InfoMethod(Fat96E).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Fat24L     'Fathead minnow 24h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Fat24L).value(InfoMethod(Fat24L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Fat48L     'Fathead minnow 48h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Fat48L).value(InfoMethod(Fat48L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Fat96L     'Fathead minnow 96h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Fat96L).value(InfoMethod(Fat96L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Sal24L     'Salmonidae 24h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Sal24L).value(InfoMethod(Sal24L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Sal48L     'Salmonidae 48h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Sal48L).value(InfoMethod(Sal48L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Sal96L     'Salmonidae 96h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Sal96L).value(InfoMethod(Sal96L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Daph24E     'Daphnia 24h, ec50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Daph24E).value(InfoMethod(Daph24E).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Daph48E     'Daphnia 24h, ec50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Daph48E).value(InfoMethod(Daph48E).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Daph24L     'Daphnia 24h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Daph24L).value(InfoMethod(Daph24L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Daph48L     'Daphnia 48h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Daph48L).value(InfoMethod(Daph48L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case Mysid96L     'Mysid 96h, lc50
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(Mysid96L).value(InfoMethod(Mysid96L).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
    
   Case AltSpecies     'other species
        If (External) Then
            answer = Toxicity_To_Standard(InfoMethod(AltSpecies).value(InfoMethod(AltSpecies).CurMethod), ConvertTo)
        Else
            If Not (FromStan) Then answer = Toxicity_To_Standard(answer, ConvertFrom)
            If Not (ToStan) And (answer <> err_flag) Then answer = Toxicity_From_Standard(answer, ConvertTo)
        End If
        
   Case Else       'if something is wrong
        answer = ERROR_FLAG
        MsgBox "could not convert units"
End Select

Convert = answer
Exit Function

err_convert:
    Convert = ERROR_FLAG
End Function
Public Function MW_To_Standard(value As Double, Units_To As String) As Double
' This function converts Molecular Weight values to different units

' Note: currently MW units are all equal, this is here just in case
' of future changes
Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_MW_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "kg/kmol"
        result = value
    Case "lb/lbmol"
        result = value
    Case "g/gmol"
        result = value
    Case Else
        result = ERROR_FLAG
End Select

MW_To_Standard = result
Exit Function

err_MW_To_Standard:
    result = ERROR_FLAG
End Function

Public Function Heat_To_Standard(value As Double, Units_To As String) As Double
' This function converts Heat values to different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    Heat_To_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Heat_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "J/kmol"
        result = value
    Case "kJ/mol"
        result = value / 1000000#
    Case "kJ/kmol"
        result = value / 1000#
    Case "J/mol"
        result = value / 1000#
    Case "cal/mol"
        result = value * 0.00023901
    Case "kcal/mol"
        result = value * 0.00000023901
    Case "cal/lbmol"
        result = value * 0.1084132
    Case "cal/g"
        result = value * 0.00023901 / MW_val
    Case "kcal/g"
        result = value * 0.00000023901 / MW_val
    Case "J/g"
        result = value / 1000 / MW_val
    Case "kJ/kg"
        result = value / 1000 / MW_val
    Case "Btu/lb"
        result = value * 0.0004299 / MW_val
    Case Else
        result = ERROR_FLAG
End Select

Heat_To_Standard = result
Exit Function

err_Heat_To_Standard:
    result = ERROR_FLAG
End Function

Public Function Odemand_To_Standard(value As Double, Units_To As String) As Double
' This function converts Heat values to different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    Odemand_To_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Odemand_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "g O2/g chem"
        result = value
    Case "mg O2/g chem"
        result = value / 1000#
    Case "mol O2/g chem"
        result = value * 32
    Case Else
        result = ERROR_FLAG
End Select

Odemand_To_Standard = result
Exit Function

err_Odemand_To_Standard:
    result = ERROR_FLAG
End Function


Public Function Odemand_From_Standard(value As Double, Units_To As String) As Double
' This function converts Heat values to different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    Odemand_From_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Odemand_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "g O2/g chem"
        result = value
    Case "mg O2/g chem"
        result = value * 1000#
    Case "mol O2/g chem"
        result = value / 32
    Case Else
        result = ERROR_FLAG
End Select

Odemand_From_Standard = result
Exit Function

err_Odemand_From_Standard:
    result = ERROR_FLAG
End Function



Public Function Heat_From_Standard(value As Double, Units_To As String) As Double
' This function converts Heat values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    Heat_From_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Heat_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "J/kmol"
        result = value
    Case "kJ/mol"
        result = value / 1000000#
    Case "kJ/kmol"
        result = value / 1000#
    Case "J/mol"
        result = value / 1000#
    Case "cal/mol"
        result = value * 0.00023901
    Case "kcal/mol"
        result = value * 0.00000023901
    Case "cal/lbmol"
        result = value * 0.1084132
    Case "cal/g"
        result = value * 0.00023901 / MW_val
    Case "kcal/g"
        result = value * 0.00000023901 / MW_val
    Case "J/g"
        result = value / 1000 / MW_val
    Case "kJ/kg"
        result = value / 1000 / MW_val
    Case "Btu/lb"
        result = value * 0.0004299 / MW_val
    Case Else
        result = ERROR_FLAG
End Select

Heat_From_Standard = result
Exit Function

err_Heat_From_Standard:
    result = ERROR_FLAG
End Function
Public Function LD_To_Standard(value As Double, Units_To As String) As Double
' This function converts LD values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    LD_To_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_LD_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "kmol/m3"
        result = value
    Case "g/cm3"    'val / tempMW / 1000# * 1000000#
        result = value / MW_val * 1000#
    Case "g/mL"     'val/ tempMW / 1000# * 1000000#
        result = value / MW_val * 1000#
    Case "g/L"      'val  / tempMW/ 1000#* 1000#
        result = value / MW_val
    Case "mg/mL"    'val / tempMW / 1000# / 1000# * 1000000
        result = value / MW_val
    Case "g/m3"     'val / tempMW / 1000#
        result = value / MW_val / 1000#
    Case "mol/L"     'val / 1000# * 1000#
        result = value
    Case "lb/gal"     'val / 2.20462 / tempMW * 264.17
        result = value / 2.20462
        result = result / MW_val
        result = result * 264.17
    Case "mmol/L"   'val / 1000000# * 1000#
        result = value / 1000#
    Case "ng/L"     'val / 1000000000# * 1000 / 1000 / tempMW
        result = value / 1000000000# / MW_val
    Case "mg/L"     'val  / tempMW/ 1000#* 1000#/ 1000#
        result = value / MW_val / 1000#
    Case "mol/cm3"  'val / 1000# * 1000000
        result = value * 1000
    Case "kg/L"     'val / tempMW * 1000# / 1000# * 1000#
        result = value / MW_val * 1000#
    Case "kg/m3"    'val / tempMW * 1000# / 1000#
        result = value / MW_val
    Case "lb/ft3"   'val / 0.028317 / tempMW
        result = value / 16.0181 / MW_val
    Case "mol/m3"   'val / 1000#
        result = value / 1000#
    Case "mol/dm3"      'val / 1000# * 1000#
        result = value
    Case "gmol/L"       'val * 1000# / 1000#
        result = value
    Case Else
        result = ERROR_FLAG
End Select

LD_To_Standard = result
Exit Function

err_LD_To_Standard:
    result = ERROR_FLAG
End Function

Public Function LD25_To_Standard(value As Double, Units_From As String) As Double
' This function converts LD values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    LD25_To_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_From = Trim(Units_From)

On Error GoTo err_LD25_To_Standard

'determine units to go from
Select Case (Units_From)
    Case "kmol/m3"
        result = value * MW_val
    Case "g/cm3"    'val /1000# * 1000000
        result = value * 1000#
    Case "g/mL"     'val / 1000# * 1000000#
        result = value * 1000#
    Case "g/L"      'val / 1000# * 1000#
        result = value
    Case "mg/mL"    'val / 1000000# * 1000000#
        result = value
    Case "g/m3"
        result = value / 1000#
    Case "mol/L"     'val * tempMW / 1000# * 1000#
        result = value * MW_val
    Case "lb/gal"
        result = value / 2.20462
        result = result * 264.17
    Case "mmol/L"   'val / 1000# * tempMW /1000# * 1000#
        result = value / 1000#
        result = result * MW_val
    Case "ng/L"     'val /1000000000# * 1000# / 1000#
        result = value / 1000000000#
    Case "mg/L"     'val / 1000000# * 1000#
        result = value / 1000#
    Case "mol/cm3"  'val * tempMW /1000# *1000000#
        result = value * MW_val
        result = result * 1000#
    Case "kg/L"
        result = value * 1000#
    Case "kg/m3"
        result = value
    Case "lb/ft3"
        result = value * 16.0181
    Case "mol/m3"
        result = value * MW_val
        result = result / 1000#
    Case "mol/dm3"      'val * tempMW /1000# * 1000#
        result = value * MW_val
    Case "gmol/L"       'val * tempMW / 1000# * 1000#
        result = value * MW_val
    Case Else
        result = ERROR_FLAG
End Select

LD25_To_Standard = result
Exit Function

err_LD25_To_Standard:
    result = ERROR_FLAG
End Function


Public Function LD_From_Standard(value As Double, Units_To As String) As Double
' This function converts LD values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    LD_From_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_LD_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "kmol/m3"
        result = value
    Case "g/cm3"    '(val * tempMW* 1000#/ 1000000#)
        result = value * MW_val / 1000#
    Case "g/mL"     'val * tempMW * 1000# / 1000000#
        result = value * MW_val / 1000#
    Case "g/L"      'val * tempMW * 1000# / 1000#
        result = value * MW_val
    Case "mg/mL"    'val* tempMW * 1000# * 1000# / 1000000
        result = value * MW_val
    Case "g/m3"     'val * tempMW * 1000#
        result = value * MW_val * 1000#
    Case "mol/L"    'val * 1000# / 1000#
        result = value
    Case "lb/gal"   'val * 2.20462 * tempMW / 264.17
        result = value * 2.20462
        result = result * MW_val
        result = result / 264.17
    Case "mmol/L"   'val * 1000000# / 1000#
        result = value * 1000#
    Case "ng/L"     ' val* 1000000000# / 1000 * 1000 *tempMW
        result = value * 1000000000# * MW_val
    Case "mg/L"     'val * tempMW * 1000# / 1000# * 1000#
        result = value * MW_val * 1000#
    Case "mol/cm3"     ' val * 1000# / 1000000
        result = value / 1000#
    Case "kg/L"         'val * tempMW / 1000# * 1000# / 1000#
        result = value * MW_val / 1000#
    Case "kg/m3"        'val * tempMW / 1000# * 1000#
        result = value * MW_val
    Case "lb/ft3"       'val * 2.20462 / 0.028317 * tempMW
        result = value * MW_val * 2.20462 * 0.028317
    Case "mol/m3"   'val * 1000#
        result = value * 1000#
    Case "mol/dm3"      'val * 1000# / 1000#
        result = value
    Case "gmol/L"       'val * 1000# / 1000#
        result = value
    Case Else
        result = ERROR_FLAG
End Select

LD_From_Standard = result
Exit Function

err_LD_From_Standard:
    result = ERROR_FLAG
End Function
Public Function LD25_From_Standard(value As Double, Units_To As String) As Double
' This function converts LD values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    LD25_From_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_LD25_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "kmol/m3"
        result = value / MW_val
    Case "g/cm3"    '(val * 1000# / 1000000#)
        result = value / 1000#
    Case "g/mL"     '(val* 1000#) / 1000000#
        result = value / 1000#
    Case "g/L"      '(val * 1000#) / 1000#
        result = value
    Case "mg/mL"    '(val * 1000000#) / 1000000#
        result = value
    Case "g/m3"
        result = value * 1000#
    Case "mol/L"    '((val / tempMW) * 1000#) / 1000#
        result = value / MW_val
    Case "lb/gal"
        result = value * 2.20462
        result = result / 264.17
    Case "mmol/L"   '(((val * 1000#)/ tempMW)* 1000#) / 1000#
        result = value * 1000#
        result = result / MW_val
    Case "ng/L"     ' val * 1000000000# / 1000# * 1000#
        result = value * 1000000000#
    Case "mg/L"     'val * 1000000# / 1000#
        result = value * 1000#
    Case "mol/cm3"     ' val / tempMW * 1000# / 1000000#
        result = value / MW_val
        result = result / 1000#
    Case "kg/L"
        result = value / 1000#
    Case "kg/m3"
        result = value
    Case "lb/ft3"
        result = value * 2.20462
        result = result * 0.028317
    Case "mol/m3"
        result = value / MW_val
        result = result * 1000#
    Case "mol/dm3"      'val * tempMW * 1000# / 1000#
        result = value / MW_val
    Case "gmol/L"       'val * tempMW * 1000# / 1000#
        result = value / MW_val
    Case Else
        result = ERROR_FLAG
End Select

LD25_From_Standard = result
Exit Function

err_LD25_From_Standard:
    result = ERROR_FLAG
End Function

Public Function HeatCapacity_From_Standard(value As Double, Units_To As String) As Double
' This function converts LD values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    HeatCapacity_From_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_HeatCapacity_From_Standard

'determine units to go from
Select Case (Units_To)
'msh    Case "J/kmol*K"
    Case "J/kmol·K", "J/kmol*K"
        result = value
    Case "Btu/lb*F"
        result = value / MW_val * 0.0002388
    Case "kJ/kmol*K"
        result = value / 1000#
    Case "cal/g*C"
        result = value / MW_val * 0.00023901
    Case "kJ/kg*K"
        result = value / MW_val / 1000#
    Case Else
        result = ERROR_FLAG
End Select

HeatCapacity_From_Standard = result
Exit Function

err_HeatCapacity_From_Standard:
    result = ERROR_FLAG
End Function
Public Function CV_From_Standard(value As Double, Units_To As String) As Double
' This function converts LD values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    CV_From_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_CV_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "m3/kmol"
        result = value
    Case "m3/kg"
        result = value / MW_val
    Case "m3/g"
        result = value / MW_val / 1000#
    Case "cm3/g"
        result = value * 1000# / MW_val
    Case "L/g"
        result = value / MW_val
    Case "L/kg"
        result = value * 1000# / MW_val
    Case "L/mg"
        result = value / 1000# / MW_val
    Case "mL/mg"
        result = value / MW_val
    Case "mL/g"
        result = value * 1000# / MW_val
    Case "L/ng"
        result = value / 1000000000# / MW_val
    Case "L/mol"
        result = value ' L/mol = m3/kmol
    Case "L/mmol"
        result = value / 1000#
    Case "gal/lb"
        result = value * 119.83223 / MW_val
    Case "ft3/lb"
        result = value / MW_val / 16.017
    Case Else
        result = ERROR_FLAG
End Select

CV_From_Standard = result
Exit Function

err_CV_From_Standard:
    result = ERROR_FLAG
End Function

Public Function CV_To_Standard(value As Double, Units_To As String) As Double
' This function converts LD values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    CV_To_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_CV_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "m3/kmol"
        result = value
    Case "m3/kg"
        result = value * MW_val
    Case "m3/g"
        result = value * 1000# * MW_val
    Case "cm3/g"
        result = value * MW_val / 1000#
    Case "L/g"
        result = value * MW_val
    Case "L/kg"
        result = value * MW_val / 1000#
    Case "L/mg"
        result = value * 1000# * MW_val
    Case "mL/mg"
        result = value * MW_val
    Case "mL/g"
        result = value * MW_val / 1000#
    Case "L/ng"
        result = value * 1000000000# * MW_val
    Case "L/mol"
         result = value ' L/mol = m3/kmol
    Case "L/mmol"
         result = value / 1000#
    Case "gal/lb"
        result = value * MW_val / 119.83223
    Case "ft3/lb"
        result = value * 0.06243 * MW_val
    Case Else
        result = ERROR_FLAG
End Select

CV_To_Standard = result
Exit Function

err_CV_To_Standard:
    result = ERROR_FLAG
End Function


Public Function HeatCapacity_To_Standard(value As Double, Units_To As String) As Double
' This function converts LD values from different units

Dim result As Double
Dim MW_val As Double
Dim MW_Units As String

'get molecular weight for calcs
MW_val = InfoMethod(MW).value(InfoMethod(MW).CurMethod)
MW_Units = InfoMethod(MW).Unit

'make sure its in right units
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    MW_val = get_MW("kg/kmol", Cur_Info.CAS)
Else
    MW_val = Convert(MW_val, MW, MW_Units, "kg/kmol", True)
End If

'double check
If MW_val = 0 Or MW_val = ERROR_FLAG Then
    HeatCapacity_To_Standard = ERROR_FLAG
    Exit Function
End If

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_HeatCapacity_To_Standard

'determine units to go from
Select Case (Units_To)
'msh    Case "J/kmol*K"
    Case "J/kmol·K ", "J/kmol*K"
        result = value
    Case "Btu/lb*F"
        result = value * MW_val / 0.0002388
    Case "kJ/kmol*K"
        result = value * 1000#
    Case "cal/g*C"
        result = value * MW_val / 0.00023901
    Case "kJ/kg*K"
        result = value * MW_val * 1000#
    Case Else
        result = ERROR_FLAG
End Select

HeatCapacity_To_Standard = result
Exit Function

err_HeatCapacity_To_Standard:
    result = ERROR_FLAG
End Function
Public Function Temp_To_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values to internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Temp_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "K"
        result = value
    Case "C"
        result = value + 273.15
    Case "F"
        result = (value + 459.67) / 1.8
    Case "R"
        result = value / 1.8
    Case Else
        result = ERROR_FLAG
End Select

Temp_To_Standard = result
Exit Function

err_Temp_To_Standard:
    result = ERROR_FLAG
End Function
Public Function Temp_From_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values to internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Temp_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "K"
        result = value
    Case "C"
        result = value - 273.15
    Case "F"
        result = value * 1.8 - 459.67
    Case "R"
        result = value * 1.8
    Case Else
        result = ERROR_FLAG
End Select

Temp_From_Standard = result
Exit Function

err_Temp_From_Standard:
    result = ERROR_FLAG
End Function

Public Function Diffusivity_To_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values from internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Diffusivity_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "cm2/s"
        result = value
    Case "m2/s"
        result = value * 10000#
    Case "in2/s"
        result = value / 0.1549996
    Case "ft2/s"
        result = value / 0.0010763
    Case Else
        result = ERROR_FLAG
End Select

Diffusivity_To_Standard = result
Exit Function

err_Diffusivity_To_Standard:
    result = ERROR_FLAG
End Function
Public Function ST_To_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values from internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_ST_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "N/m"
        result = value
    Case "erg/cm2"
        result = value / 1000#
    Case "lbf/ft"
        result = value / 0.0685156
    Case "lbf/in"
        result = value / 0.00570996414
    Case "dynes/cm"
        result = value / 1000#
    Case Else
        result = ERROR_FLAG
End Select

ST_To_Standard = result
Exit Function

err_ST_To_Standard:
    result = ERROR_FLAG
End Function
Public Function Viscosity_To_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values from internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Viscosity_To_Standard

'determine units to go from
Select Case (Units_To)
'msh    Case "Pa*s"
    Case "Pa·s", "Pa*s"
        result = value
    Case "cp"
        result = value / 1000#
    Case "kg/m*h"
        result = value / 3600#
    Case "kg/m*s"
        result = value
    Case "lb/ft*hr"
        result = value / 2419.1
    Case "lb/ft*s"
        result = value / 0.6719722
    Case Else
        result = ERROR_FLAG
End Select

Viscosity_To_Standard = result
Exit Function

err_Viscosity_To_Standard:
    result = ERROR_FLAG
End Function


Public Function FL_From_Standard(value As Double, Units_To As String) As Double
' This function returns the value it is sent as there is
' only one unit currently that the Flammibility properties
' can be in

FL_From_Standard = value
End Function

Public Function HC_From_Standard(value As Double, Units_To As String) As Double
End Function

Public Function HC_To_Standard(value As Double, Units_To As String) As Double
End Function

Public Function Sol_To_Standard(value As Double, Units_To As String) As Double
End Function

Public Function Sol_From_Standard(value As Double, Units_To As String) As Double
End Function
Public Function NoUnits_From_Standard(value As Double, Units_To As String) As Double
' This function returns the value it is sent as there is
' only unit-less is acceptable as a unit parameter

NoUnits_From_Standard = value
End Function

Public Function NoUnits_To_Standard(value As Double, Units_To As String) As Double
' This function returns the value it is sent as there is
' only unit-less is acceptable as a unit parameter

NoUnits_To_Standard = value
End Function
Public Function FL_To_Standard(value As Double, Units_To As String) As Double
' This function returns the value it is sent as there is
' only one unit currently that the Flammibility properties
' can be in

FL_To_Standard = value
End Function

Public Function LogKOC_To_Standard(value As Double, Units_To As String) As Double
' This function returns the value it is sent as there is
' only one unit that is acceptable as a unit parameter
' for log Koc

LogKOC_To_Standard = value
End Function

Public Function LogKOC_From_Standard(value As Double, Units_To As String) As Double
' This function returns the value it is sent as there is
' only one unit that is acceptable as a unit parameter
' for log Koc

LogKOC_From_Standard = value
End Function
Public Function TC_To_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values from internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo Err_TC_To_Standard

'determine units to go from
Select Case (Units_To)
'msh    Case "W/m*K"
    Case "W/m·K", "W/m*K"
        result = value
    Case "kcal/m*hr*C"
        result = value / 0.8604
    Case "cal/cm*s*C"
        result = value / 0.00239
    Case "Btu/ft*hr*F"
        result = value / 0.5777908
    Case Else
        result = ERROR_FLAG
End Select

TC_To_Standard = result
Exit Function

Err_TC_To_Standard:
    result = ERROR_FLAG
End Function

Public Function TC_From_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values from internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_TC_From_Standard

'determine units to go from
Select Case (Units_To)
'msh    Case "W/m*K"
    Case "W/m·K", "W/m*K"
        result = value
    Case "kcal/m*hr*C"
        result = value * 0.8604
    Case "cal/cm*s*C"
        result = value * 0.00239
    Case "Btu/ft*hr*F"
        result = value * 0.5777908
    Case Else
        result = ERROR_FLAG
End Select

TC_From_Standard = result
Exit Function

err_TC_From_Standard:
    result = ERROR_FLAG
End Function


Public Function Viscosity_From_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values from internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Viscosity_From_Standard

'determine units to go from
Select Case (Units_To)
'msh    Case "Pa*s"
    Case "Pa·s", "Pa*s"
        result = value
    Case "cp"
        result = value * 1000#
    Case "kg/m*h"
        result = value * 3600#
    Case "kg/m*s"
        result = value
    Case "lb/ft*hr"
        result = value * 2419.1
    Case "lb/ft*s"
        result = value * 0.6719722
    Case Else
        result = ERROR_FLAG
End Select

Viscosity_From_Standard = result
Exit Function

err_Viscosity_From_Standard:
    result = ERROR_FLAG
End Function


Public Function ST_From_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values from internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_ST_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "N/m"
        result = value
    Case "erg/cm2"
        result = value * 1000#
    Case "lbf/ft"
        result = value * 0.0685156
    Case "lbf/in"
        result = value * 0.00570996414
    Case "dynes/cm"
        result = value * 1000#
    Case Else
        result = ERROR_FLAG
End Select

ST_From_Standard = result
Exit Function

err_ST_From_Standard:
    result = ERROR_FLAG
End Function

Public Function Diffusivity_From_Standard(value As Double, Units_To As String) As Double
' This function converts Tempreture values from internal standards

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo err_Diffusivity_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "cm2/s"
        result = value
    Case "m2/s"
        result = value / 10000#
    Case "in2/s"
        result = value * 0.1549996
    Case "ft2/s"
        result = value * 0.0010763
    Case Else
        result = ERROR_FLAG
End Select

Diffusivity_From_Standard = result
Exit Function

err_Diffusivity_From_Standard:
    result = ERROR_FLAG
End Function
Public Function Press_To_Standard(value As Double, Units_To As String) As Double
' This function converts Pressure values to internal standard units

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo Err_Press_To_Standard

'determine units to go from
Select Case (Units_To)
    Case "Pa"   'internal standard
        result = value
    Case "mPa"
        result = value / 1000#
    Case "kPa"
        result = value * 1000#
    Case "MPa"
        result = value * 1000000#
    Case "psig"
        result = (value + 14.696) / 0.000145
    Case "lbf/ft2"
        result = value / 0.0208825
    Case "psia"
        result = value / 0.000145
    Case "atm"
        result = value * 101325#
    Case "cm Hg"
        result = value / 0.00075
    Case "kN/m2"
        result = value * 1000#
    Case "bar"
        result = value * 100000#
    Case "lbf/in2"
        result = value / 0.000145
    Case "mm Hg"
        result = value / 0.0075006
    Case "mbar"
        result = value * 100#
    Case "torr"
        result = value / 0.0075006
    Case "N/m2"
        result = value
    Case Else   'bad values
        result = ERROR_FLAG
End Select

Press_To_Standard = result
Exit Function

Err_Press_To_Standard:
    result = ERROR_FLAG

End Function
Public Function Press_From_Standard(value As Double, Units_To As String) As Double
' This function converts Pressure values from internal standard units

Dim result As Double

'remove whitespace
Units_To = Trim(Units_To)

On Error GoTo Err_Press_From_Standard

'determine units to go from
Select Case (Units_To)
    Case "Pa"   'internal standard
        result = value
    Case "mPa"
        result = value * 1000#
    Case "kPa"
        result = value / 1000#
    Case "MPa"
        result = value / 1000000#
    Case "psig"
        result = value * 0.000145 - 14.696
    Case "lbf/ft2"
        result = value * 0.0208825
    Case "psia"
        result = value * 0.000145
    Case "atm"
        result = value / 101325#
    Case "cm Hg"
        result = value * 0.00075
    Case "kN/m2"
        result = value / 1000#
    Case "bar"
        result = value / 100000#
    Case "lbf/in2"
        result = value * 0.000145
    Case "mm Hg"
        result = value * 0.0075006
    Case "mbar"
        result = value / 100#
    Case "torr"
        result = value * 0.0075006
    Case "N/m2"
        result = value
    Case Else   'bad values
        result = ERROR_FLAG
End Select

Press_From_Standard = result
Exit Function

Err_Press_From_Standard:
    result = ERROR_FLAG

End Function

Public Function MW_From_Standard(value As Double, Units_From As String) As Double
' This function converts Molecular Weight values from different units

' Note: currently MW units are all equal, this is here just in case
' of future changes
Dim result As Double

'remove whitespace
Units_To = Trim(Units_From)

On Error GoTo err_MW_From_Standard

'determine units to go from
Select Case (Units_From)
    Case "kg/kmol"
        result = value
    Case "lb/lbmol"
        result = value
    Case "g/gmol"
        result = value
'msh 2/16/99
'researhing...
'    Case "g/mol"
'        result = value
    Case Else
        result = ERROR_FLAG
End Select

MW_From_Standard = result
Exit Function

err_MW_From_Standard:
    result = ERROR_FLAG
End Function

Public Function IsDefault(unit_passed As String, property_passed As Integer) As Boolean
'this tell if the unit passed is in the interal standard unit or not for that property

Dim answer As Boolean
answer = False
    
      Select Case property_passed
            Case OptTemp
                If Trim(unit_passed) = "K" Then
                    answer = True
                End If
            Case OptPress
                If Trim(unit_passed) = "Pa" Then
                    answer = True
                End If
            Case MW
                If Trim(unit_passed) = "kg/kmol" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "kg*kmol"
                End If
            Case LD25
                If Trim(unit_passed) = "kg/m3" Then
                    answer = True
                End If
            Case LD
                If Trim(unit_passed) = "kmol/m3" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "kmol*m3"
                End If
            Case mp
                If Trim(unit_passed) = "K" Then
                    answer = True
                End If
            Case NBP
                If Trim(unit_passed) = "K" Then
                    answer = True
                End If
            Case VP25
                If Trim(unit_passed) = "Pa" Then
                    answer = True
                End If
            Case VP
                If Trim(unit_passed) = "Pa" Then
                    answer = True
                End If
            Case hfor
                If Trim(unit_passed) = "J/kmol" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "J*kmol"
                End If
            Case LHC
                If Trim(unit_passed) = "J/kmol*K" Then
                    answer = True
                ElseIf Trim(unit_passed) = "J/kmol-K" Then
                    answer = True
                End If
            Case VHC
                If Trim(unit_passed) = "J/kmol*K" Then
                    answer = True
                End If
            Case Hvap25
                If Trim(unit_passed) = "J/kmol" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "J*kmol"
                End If
            Case HvapNBP
                If Trim(unit_passed) = "J/kmol" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "J*kmol"
                End If
            Case Hvap
                If Trim(unit_passed) = "J/kmol" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "J*kmol"
                End If
            Case CT
                If Trim(unit_passed) = "K" Then
                    answer = True
                End If
            Case CP
                If Trim(unit_passed) = "Pa" Then
                    answer = True
                End If
            Case Dwater
                If Trim(unit_passed) = "cm2/s" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "cm2*s"
                End If
            Case Dair
                If Trim(unit_passed) = "cm2/s" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "cm2*s"
                End If
            Case ST25
                If Trim(unit_passed) = "N/m" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "N*m"
                End If
            Case ST
                If Trim(unit_passed) = "N/m" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "N*m"
                End If
            Case VV
                If Trim(unit_passed) = "Pa*s" Then
                    answer = True
                ElseIf Trim(unit_passed) = "Pa-s" Then
                    answer = True
                End If
            Case LV
                If Trim(unit_passed) = "Pa*s" Then
                    answer = True
                ElseIf Trim(unit_passed) = "Pa-s" Then
                    answer = True
                End If
            Case LTC
                If Trim(unit_passed) = "W/m*K" Then
                    answer = True
                ElseIf Trim(unit_passed) = "W/m-K" Then
                    answer = True
                End If
            Case VTC
                If Trim(unit_passed) = "W/m*K" Then
                    answer = True
                ElseIf Trim(unit_passed) = "W/m-K" Then
                    answer = True
                End If
            Case UFL
                If Trim(unit_passed) = "vol% in air" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "vol*"
                End If
            Case LFL
                If Trim(unit_passed) = "vol% in air" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "vol*"
                End If
            Case FP
                If Trim(unit_passed) = "K" Then
                    answer = True
                End If
            Case AIT
                If Trim(unit_passed) = "K" Then
                    answer = True
                End If
            Case Hcomb
                If Trim(unit_passed) = "J/kmol" Then
                    answer = True
                End If
            Case ThODcarb
                If Trim(unit_passed) = "g O2/g chem" Then
                    answer = True
                End If
            Case ThODcomb
                If Trim(unit_passed) = "g O2/g chem" Then
                    answer = True
                End If
            Case COD
                If Trim(unit_passed) = "g O2/g chem" Then
                    answer = True
                End If
            Case BOD
                If Trim(unit_passed) = "g O2/g chem" Then
                    answer = True
                End If
            Case ACwater
                If Trim(unit_passed) = "unit-less" Then
                    answer = True
                End If
            Case HC
                If Trim(unit_passed) = "kPa*mol/mol" Then
                    answer = True
                Else
                    answer = Trim(unit_passed) Like "kPa*mol/mol"
                End If
            Case ACchem
                If Trim(unit_passed) = "unit-less" Then
                    answer = True
                End If
            Case logKow
                If Trim(unit_passed) = "unit-less" Then
                    answer = True
                ElseIf Trim(unit_passed) = "Log Kow" Then
                    answer = True
                End If
            Case logKoc
                If Trim(unit_passed) = "cm3/g OC" Then
                    answer = True
                End If
            Case BCF
                If Trim(unit_passed) = "unit-less" Then
                    answer = True
                End If
            Case CV
                If Trim(unit_passed) = "m3/kmol" Then
                    answer = True
                End If
            Case Schem
                If Trim(unit_passed) = "ppm*" Or Trim(unit_passed) = "ppm" Then
                    answer = True
                End If
            Case Swater
                If Trim(unit_passed) = "ppm*" Or Trim(unit_passed) = "ppm" Then
                    answer = True
                End If
            Case Fat48E
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Fat96E
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Fat24L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Fat48L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
             Case Fat96L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Sal24L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Sal48L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Sal96L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Daph24E
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Daph48E
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Daph24L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Daph48L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case Mysid96L
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
            Case AltSpecies
                If Trim(unit_passed) = "mg/L" Then
                    answer = True
                End If
        End Select
        
        IsDefault = answer
End Function
