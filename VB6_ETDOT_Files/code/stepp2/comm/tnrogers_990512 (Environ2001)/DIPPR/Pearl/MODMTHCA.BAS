Attribute VB_Name = "modmthcalcs"
Option Explicit


Function CalcDwaterHaydukMinhas(T As Double, TUnits As String, method_units As String) As Double
' REVISIONS:  DMW  6/15/97  added unit parameter to check the units this is returned in
    Dim MV As Double
    Dim tempT As Double
    Dim sub_value As Double
    Dim sub_units As String
    
    sub_units = "cm2/s"
    On Error Resume Next
    
    If Trim(TUnits) <> "K" Then
        tempT = Convert(T, OptTemp, TUnits, "K", False)
    Else
        tempT = T
    End If
    MV = CalcMVNBPSchroeder("cm3/mol")
    
    If MV <> 0 Then
        sub_value = 0.0000000125 * (MV ^ -0.19 - 0.292) * tempT ^ 1.52
        If Trim(method_units) <> sub_units Then
            sub_value = Convert(sub_value, Dwater, sub_units, method_units, False)
        End If
    Else
        sub_value = 0
    End If
    CalcDwaterHaydukMinhas = sub_value
End Function

Function cosh(Number As Double) As Double
    
    cosh = (Exp(Number) + Exp(-Number)) / 2
    
End Function
Function ln(Number As Double)

    'Returns the natural log of a number
    'NOTE: In VB log = ln
    
    If (Number > 0) Then
    
        ln = Log(Number)
    Else
        ln = ERROR_FLAG
    End If

End Function

Function log10(Number As Double) As Double

    'Return the log base 10 of a number
    'NOTE: In VB log = ln
    ' REVISIONS     5/20/98  DMW fixed to return ERROR_FLAG if invalid
    If Number > 0 Then
        
        log10 = Log(Number) / Log(10#)
    Else
        log10 = ERROR_FLAG
    End If

End Function
Function sinh(Number As Double) As Double
    
    sinh = (Exp(Number) - Exp(-Number)) / 2

End Function


Function CalcFofT(Code As Integer, CNT As Integer) As Double
    'Calculate value using f(T) correlation
    
    Dim T As Double
    Dim Tr As Double
    Dim a As Double
    Dim b As Double
    Dim C As Double
    Dim D As Double
    Dim E As Double
    
    On Error Resume Next
    
    'Set calculation temperature
    T = InfoMethod(Code).TFT
    
    'Check to see T is in chemical temperature range
    
    If T < InfoMethod(Code).MinT(CNT) Or T > InfoMethod(Code).MaxT(CNT) Then Exit Function
    
    'Convert T to K if necessary
    If Trim(InfoMethod(CurProp).TFTUnit) <> "K" Then
        T = Convert(InfoMethod(CurProp).TFT, OptTemp, InfoMethod(CurProp).TFTUnit, "K", False)
    End If
        
    'Set equation number and f(T) coefficients
    a = InfoMethod(Code).Coeff(CNT, 1)
    b = InfoMethod(Code).Coeff(CNT, 2)
    C = InfoMethod(Code).Coeff(CNT, 3)
    D = InfoMethod(Code).Coeff(CNT, 4)
    E = InfoMethod(Code).Coeff(CNT, 5)
     
    'Calculate reduced pressure
    If InfoMethod(CT).Enabled(1) = True Then
        Tr = T / InfoMethod(CT).value(1)
    End If
    
    Select Case InfoMethod(Code).EqNum(CNT)
        Case 100:
            CalcFofT = a + b * T + C * T ^ 2 + D * T ^ 3 + E * T ^ 4
            If Code = LHC Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "J/kmol*K", Get_DefaultUnit(Code), False)
            ElseIf Code = ST Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "N/m", Get_DefaultUnit(Code), False)
            ElseIf Code = LTC Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "W/m*K", Get_DefaultUnit(Code), False)
            ElseIf Code = VHC Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "J/kmol*K", Get_DefaultUnit(Code), False)
            Else
                MsgBox ("don't know code, eq num is 100, code is " & Code)
            End If
        Case 101:
            CalcFofT = Exp(a + (b / T) + (C * ln(T)) + D * T ^ E)
            If Code = LV Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "Pa*s", Get_DefaultUnit(Code), False)
            ElseIf Code = VP Then    ' it's VP ??
                CalcFofT = Convert(CalcFofT, CLng(Code), "Pa", Get_DefaultUnit(Code), False)
            Else
                MsgBox ("don't know code, equation is 101, Code is " & Code)
            End If
                
        Case 102:
            CalcFofT = a * T ^ b / (1 + C / T + D / T ^ 2)
            If Code = VV Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "Pa*s", Get_DefaultUnit(Code), False)
            ElseIf Code = Hvap Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "J/kmol", Get_DefaultUnit(Code), False)
            ElseIf Code = VTC Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "W/m*K", Get_DefaultUnit(Code), False)
            Else
                MsgBox ("don't know code Eq is 102, Code is " & Code)
            End If
        Case 105:
            CalcFofT = a / (b ^ (1 + (1 - (T / C)) ^ D))    ' LD
            If Code = LD Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "kmol/m3", Get_DefaultUnit(Code), False)
            Else
                MsgBox ("don't know code, eq is 105, code is " & Code)
            End If
        Case 106:
            CalcFofT = a * (1 - Tr) ^ (b + C * Tr + D * Tr ^ 2 + E * Tr ^ 3)
            If Code = ST Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "N/m", Get_DefaultUnit(Code), False)
            ElseIf Code = Hvap Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "J/kmol", Get_DefaultUnit(Code), False)
            Else
                MsgBox ("don't know code, eq is 106, code is " & Code)
            End If
        Case 107:
            CalcFofT = a + b * ((C / T) / sinh(C / T)) ^ 2 + D * ((E / T) / cosh(E / T)) ^ 2
            If Code = VHC Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "J/kmol*K", Get_DefaultUnit(Code), False)
            Else
                MsgBox ("don't know code, eq is 107, code is " & Code)
            End If
        Case 114:
            CalcFofT = (a ^ 2 / (1 - Tr)) + b - (2 * a * C * (1 - Tr)) - (a * D * (1 - Tr) ^ 2) _
                        - ((C ^ 2 * (1 - Tr) ^ 3) / 3) - ((C * D * (1 - Tr) ^ 4) / 2) - ((D ^ 2 * (1 - Tr) ^ 5) / 5)
            CalcFofT = Convert(CalcFofT, CLng(Code), "J/kmol*K", Get_DefaultUnit(Code), False)
        Case 115:
            CalcFofT = 2.718282 ^ (a + (b / T) + (C * ln(T)) + (D * T ^ 2) + (E / T ^ 2))
            CalcFofT = Convert(CalcFofT, CLng(Code), "Pa", Get_DefaultUnit(Code), False)
        Case 116:
            CalcFofT = a + (b * (1 - Tr) ^ 0.35) + (C * (1 - Tr) ^ (2 / 3)) + (D * (1 - Tr)) + (E * (1 - Tr) ^ (4 / 3))
            CalcFofT = Convert(CalcFofT, CLng(Code), "kmol/m3", Get_DefaultUnit(Code), False)
        Case 200:
            MsgBox ("don't know code, eq is 200, code is " & Code)
            CalcFofT = a + b * T + C * T ^ 2 * ln(T) + D * T ^ 2.5 + E * T ^ 3
            CalcFofT = Convert(CalcFofT, CLng(Code), Get_DefaultUnit(Code), Get_DefaultUnit(Code), False)
        Case 201:
        MsgBox ("don't know code, eq is 201, code is " & Code)
            CalcFofT = a + b * T ^ 2 * ln(T) + C * T ^ 2.5 + D * T ^ 3
            CalcFofT = Convert(CalcFofT, CLng(Code), Get_DefaultUnit(Code), Get_DefaultUnit(Code), False)
        Case 202:
        
            T = T - 273.15
            CalcFofT = (Exp(a + b / (T + C))) * 133.33
            If Code = VP Then
                CalcFofT = Convert(CalcFofT, CLng(Code), "Pa", Get_DefaultUnit(Code), False)
            Else
                MsgBox ("code is " & Code & " eq num is 202")
            End If
    End Select
                
    'CalcFofT = Convert(CLng(Code), ConvertToDefault(Code), ConvertToDefault(Code), CalcFofT)
                
End Function


Function CalcMPTaftStarek(Tc As Double, Tc_Units As String, Tb As Double, Tb_Units As String, default_units As String) As Double
'Calculation for Melting Point (K)
'
'Method: Taft and Starek (1930)
'
'Equation Inputs: Tc (K)    - Critical Temperature
'                 Tb (K)    - Boiling Point
'
' Modified 6/18/97 BGH : Added Unit parameters
Dim tempTc As Double
Dim tempTb As Double
Dim value As Double
Dim method_units As String

method_units = "K"  'method computes answer in K

    On Error Resume Next
        
    ' Check inputs for correct units
    If Trim(Tc_Units) = "K" Then
        tempTc = Tc
    Else
        tempTc = Convert(Tc, CT, Tc_Units, "K", False)
    End If
    
    If Trim(Tb_Units) = "K" Then
        tempTb = Tb
    Else
        tempTb = Convert(Tb, NBP, Tb_Units, "K", False)
    End If
    
    value = tempTc - tempTb
    
    ' Covert the answer if it is in the wrong units
    If Trim(method_units) = Trim(default_units) Then
        CalcMPTaftStarek = value
    Else
        CalcMPTaftStarek = Convert(value, FP, method_units, default_units, False)
    End If
    
End Function

Function CalcMPLorenzHerz(Tb As Double, Tb_Units As String, default_units As String) As Double
'Calculation for Melting Point (K)
'
'Method: Lorenz and Herz (1922)
'
'Equation Inputs: Tb (K)    - Boiling Point
'
'Modified 6/18/97 BGH: Added unit parameters
Dim tempTb As Double
Dim value As Double
Dim method_units As String

method_units = "K"

    On Error Resume Next
    
'check inputs for correct units
If Trim(Tb_Units) = "K" Then
    tempTb = Tb
Else
    tempTb = Convert(Tb, NBP, Tb_Units, "K", False)
End If

value = 0.5839 * tempTb

'make sure answer is in correct units
If Trim(method_units) = Trim(default_units) Then
    CalcMPLorenzHerz = value
Else
    CalcMPLorenzHerz = Convert(value, FP, method_units, default_units, False)
End If

End Function

Function CalcBCFKobayshi(log10Kow As Double, default_units As String) As Double
 'Calculation for Bioconcentration Factor (unit-less)
'
'Method: Kobayshi (1981)
'
'Equation Inputs: log Kow (unit-less)   - log(Octanol/Water Partition Coefficient)
'
' Modified 6/25/97 BGH: added unit parameters

Dim method_units As String
Dim value As Double

method_units = "unit-less"

    On Error Resume Next
        
    value = 10 ^ (0.74 * log10Kow - 0.77)
    
    'make sure answer is in correct units
    If Trim(method_units) = Trim(default_units) Then
        CalcBCFKobayshi = value
    Else
        CalcBCFKobayshi = Convert(value, BCF, method_units, default_units, False)
    End If
    
End Function

Function CalcBCFKenagaGoring(S As Double, S_Units As String, default_units As String) As Double
'Calculation for Bioconcentration Factor (unit-less)
'
'Method: Kenaga and Goring (1980)
'
'Equation Inputs: M (g/mol)     - Molecular Weight  ** removed 6/25/97 **
'                 S (ppm(wt))   - Solubility of chemical in water
'
' Modified 6/25/97 BGH: added unit paramaters and removed MW as a parameter
'   since its only purpose was to convert S
    
    Dim value As Double
    Dim tempS As Double
    Dim method_units As String
    
    method_units = "unit-less"
    On Error Resume Next
    
    'Convert S to ppm(wt)
    'S = S * M * 1000
    
    'check inputs for correct units
    If Trim(S_Units) Like "ppm*" Then
        tempS = S
    Else
        tempS = Convert(S, Schem, S_Units, "ppm(wt)", False)
    End If
    
    value = 10 ^ (-0.55 * log10(tempS) + 2.83)
    
    'make sure answer is in correct units
    If Trim(method_units) = Trim(default_units) Then
        CalcBCFKenagaGoring = value
    Else
        CalcBCFKenagaGoring = Convert(value, BCF, method_units, default_units, False)
    End If

    'Convert S to kmol/m3 water
    'S = S / (M * 1000)

End Function

Function CalcThODcarbBaker(Formula As String, default_units As String) As Double
'Calculation for Carbonaceous ThOD (unit-less)
'
'Method: Baker (1994)
'
'Equation Inputs: Formula
'
' Modified 6/19/97 BGH: Added unit parameter
    
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
    Dim value As Double
    
    method_units = "unit-less"
    On Error Resume Next
    
    If Formula = "" Then Exit Function

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
    value = cnum + ((HNum - (BrNum + FNum + ClNum + INum) - (3 * NNum) - (2 * SNum) - (3 * PNum)) / 4) - (ONum / 2) + (2 * SNum) + (2 * PNum)
    
    'make sure answer is in correct units
    If Trim(method_units) = Trim(default_units) Then
        CalcThODcarbBaker = value
    Else
        CalcThODcarbBaker = Convert(value, ThODcarb, method_units, default_units, False)
    End If

End Function

Function CalcThODcombBaker(Formula As String, default_units As String) As Double
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
    Dim value As Double
    Dim method_units As String
    
    method_units = "unit-less"
    On Error Resume Next
    
    If Formula = "" Then Exit Function

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
    value = cnum + ((HNum - (BrNum + FNum + ClNum + INum) - NNum - (2 * SNum) - (3 * PNum)) / 4) - (ONum / 2) + ((3 * NNum) / 2) + (2 * SNum) + (2 * PNum)
    
    'check to make sure answer is in correct units
    If Trim(method_units) = Trim(default_units) Then
        CalcThODcombBaker = value
    Else
        CalcThODcombBaker = Convert(value, ThODcomb, method_units, default_units, False)
    End If
    
End Function

Function CalcACchemUNIFAC(T As Double, T_Units As String, BIPCode As Integer, default_units As String) As Double
'Calculation for Infinite Dilution Activity Coefficient of Chemical in Water (unit-less)
'
'Method: UNIFAC
'
'Equation Inputs: T (C)     - Temperature
'                 BIPCode   - BIP database
'
' Modified 6/24/97 BGH: added unit parameters

    Dim i As Integer
    Dim j As Integer
    Dim GAM As Double
    Dim GAMSS As Long
    Dim GAMLS As Long
    Dim GAMErr As Long
    Dim GAMTemp As Double
    Dim FGRPErr As Long
    Dim MX As Long
    Dim AI(1 To 58, 1 To 58) As Double
    Dim MST(1 To 10, 1 To 10, 1 To 2) As Long
    Dim tempT As Double
    Dim value As Double
    Dim method_units As String
    
    method_units = "unit-less"

    On Error Resume Next
    
    'make sure inputs are in correct units
    If T_Units = "C" Then
        tempT = T
    Else
        tempT = Convert(T, OptTemp, T_Units, "C", False)
    End If
    
    'Convert T to C
    'T = T - 273.15
    
    MX = Cur_Info.MaxGroups
            
    If MX <= 0 Then Exit Function

    For i = 1 To 10
        For j = 1 To 10
            MST(i, j, 1) = 0
            MST(i, j, 2) = 0
        Next j
    Next i

    For i = 1 To 10
        MST(2, i, 1) = Cur_Info.Grp(i)
        MST(2, i, 2) = Cur_Info.NumGrp(i)
    Next i
       
    GAM = 0     'Returned Value
    GAMSS = 0   'Not Important
    GAMLS = 0   'Not Important
    GAMErr = 0  'Not Important
    GAMTemp = 0 'Not Important
    FGRPErr = 0 'Not Important
    
    If BIPCode <> 4 Then
        For i = 1 To 58
            For j = 1 To 58
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    Else
        For i = 1 To 32
            For j = 1 To 32
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    End If
    
    Call ACCALL(GAM, GAMSS, GAMLS, GAMErr, GAMTemp, tempT, FGRPErr, MX, MST(1, 1, 1), MGSG(1), AI(1, 1), RI(1), QI(1), MWS(1), MVS(1))

    If GAMErr = -1 Or GAM = 0 Then
        CalcACchemUNIFAC = ERROR_FLAG
        Exit Function
    Else
        value = GAM
        ' check to make sure answer is in correct units
        If Trim(method_units) = Trim(default_units) Then
            CalcACchemUNIFAC = value
        Else
            CalcACchemUNIFAC = Convert(value, ACchem, method_units, default_units, False)
        End If
        
    End If
    
    'Convert T to K
    'T = T + 273.15

    Exit Function
        
End Function

Function CalcACwaterUNIFAC(T As Double, T_Units As String, BIPCode As Integer, default_units As String) As Double
'Calculation for Infinite Dilution Activity Coefficient of Water in Chemical (unit-less)
'
'Method: UNIFAC
'
'Equation Inputs: T (C)     - Temperature
'                 BIPCode   - BIP database
'
' Modified 6/19/97 BGH: Added Unit Parameters
    Dim i As Integer
    Dim j As Integer
    Dim GAM As Double
    Dim GAMSS As Long
    Dim GAMLS As Long
    Dim GAMErr As Long
    Dim GAMTemp As Double
    Dim FGRPErr As Long
    Dim MX As Long
    Dim AI(1 To 58, 1 To 58) As Double
    Dim MST(1 To 10, 1 To 10, 1 To 2) As Long
    Dim tempT As Double
    Dim value As Double
    Dim method_units As String
    
    method_units = "unit-less"
    On Error Resume Next

    'Convert T to C
    'T = T - 273.15
    
    'check to see if T needs to be converted to C
    If Trim(T_Units) = "C" Then
        tempT = T
    Else
        tempT = Convert(T, OptTemp, T_Units, "C", False)
    End If
    
    MX = Cur_Info.MaxGroups
    
    If MX <= 0 Then Exit Function
   
    For i = 1 To 10
        For j = 1 To 10
            MST(i, j, 1) = 0
            MST(i, j, 2) = 0
        Next j
    Next i

    For i = 1 To 10
        MST(2, i, 1) = Cur_Info.Grp(i)
        MST(2, i, 2) = Cur_Info.NumGrp(i)
    Next i

    GAM = 0     'Returned Value
    GAMSS = 0   'Not Important
    GAMLS = 0   'Not Important
    GAMErr = 0  'Not Important
    GAMTemp = 0 'Not Important
    FGRPErr = 0 'Not Important

    If BIPCode <> 4 Then
        For i = 1 To 58
            For j = 1 To 58
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    Else
        For i = 1 To 32
            For j = 1 To 32
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    End If
    
    Call ACCALL2(GAM, GAMSS, GAMLS, GAMErr, GAMTemp, tempT, FGRPErr, MX, MST(1, 1, 1), MGSG(1), AI(1, 1), RI(1), QI(1), MWS(1), MVS(1))

    If GAMErr = -1 Or GAM = 0 Then
        CalcACwaterUNIFAC = ERROR_FLAG
        Exit Function
    Else
        CalcACwaterUNIFAC = GAM
    End If

    'Convert T to K
    'T = T + 273.15
    
    Exit Function
    
End Function

Function CalcACHansch(log10Kow As Double, default_units As String) As Double
'Calculation for Infinite Dilution Activity Coefficient of Chemical in Water (unit-less)
'
'Method: Hansch (1968)
'
'Equation Inputs: log10Kow (unit-less)  - log(Octanol/Water Partition Coefficient)
'
' Modified 6/24/97 BGH : Added unit parameter

Dim value As Double

    On Error Resume Next
            
    value = 10 ^ (log10Kow + 0.94)

If Trim(default_units) = "unit-less" Then
    CalcACHansch = value
Else
    CalcACHansch = Convert(value, ACchem, "unit-less", default_units, False)
End If


End Function

'Calculation for Gas Diffusivity (m2/s)
'
'Method: Wilke Lee
'
'Equation Inputs: M  (g/mol)    - Molecular Weight
'                 T  (K)        - Temperature
'                 P  (bar)      - Pressure
'                 Tb (K)        - Boiling Point
'
'  REVISIONS:  DMW  6/15/97  added unit parameter to check the units this is returned in
Function CalcDairWilkeLeeMod(T As Double, TUnits As String, P As Double, Punits As String, M As Double, Munits As String, Tb As Double, Tbunits As String, method_units As String) As Double
   
    Dim M12 As Double
    Dim MV As Double
    Dim sigma As Double
    Dim sigma12 As Double
    Dim omegaD As Double
    Dim eoverk As Double
    Dim eoverk12 As Double
    Dim Tp As Double
    
    Dim tempP As Double
    Dim tempT As Double
    Dim tempM As Double
    Dim tempTb As Double
    Dim sub_value As Double
    Dim sub_units As String
    
    sub_units = "m2/s"
    On Error Resume Next
    
    If Trim(TUnits) <> "K" Then
        tempT = Convert(T, OptTemp, TUnits, "K", False)
    Else
        tempT = T
    End If
    If Trim(Tbunits) <> "K" Then
        tempTb = Convert(Tb, NBP, Tbunits, "K", False)
    Else
        tempTb = Tb
    End If
    If Trim(Munits) <> "g/mol" Then
        tempM = Convert(M, MW, Munits, "g/mol", False)
    Else
        tempM = M
    End If
    MV = CalcMVNBPSchroeder("cm3/mol")
    
    If MV <> 0 Then
        'Convert P to bar if necessary
        If Trim(Punits) <> "bar" Then
           
            tempP = P * (101325 / 769)
        Else
            tempP = P
        End If
        sigma = 1.18 * MV ^ (1 / 3)
        eoverk = 1.15 * tempTb
        eoverk12 = (97 * eoverk) ^ 0.5
        sigma12 = (3.62 + sigma) / 2
        M12 = 2 * (1 / 28.84 + 1 / tempM) ^ -1
        Tp = tempT / eoverk12
        omegaD = 1.06036 / Tp ^ 0.1561 + 0.193 / Exp(0.47635 * Tp) + 1.03587 / Exp(1.52996 * Tp) + 1.76474 / Exp(3.89411 * Tp)
        sub_value = (3.03 - 0.98 / M12 ^ 0.5) * 10 ^ -3 * tempT ^ (3 / 2) / tempP * M12 ^ 0.5 * sigma12 ^ 2 * omegaD
        'Convert P to bar
        'P = P * (101325 / 760)
        If Trim(method_units) <> sub_units Then
            sub_value = Convert(sub_value, Dair, sub_units, method_units, False)
        End If
    Else
        sub_value = 0
    End If
    
    CalcDairWilkeLeeMod = sub_value
    
    
End Function

Function CalcHCUNIFAC(AC As Double, Pv As Double, method_units As String) As Double
'Calculation for Henry's Law Constant (Pa*mol/mol)
'
'Method: UNIFAC
'
'Equation Inputs: AC (unit-less)    - Activity Coefficient
'                 Pv (Pa)           - Vapor Pressure
'
' REVISIONS:  6/6/97  DMW  :  added a unit parameter since the units used here are no
'                               longer the default units

    Dim result As Double
    Dim sub_units As String
    
    sub_units = "Pa*mol/mol"
    On Error Resume Next
    
    result = AC * Pv
    If Trim(method_units) <> Trim(sub_units) Then
        result = Convert(result, HC, sub_units, method_units, False)
    End If
    If result = 0 Then
        result = ERROR_FLAG
    End If
    CalcHCUNIFAC = result
End Function

Function CalcHvapWatson(HVb As Double, HVb_Units As String, Tb As Double, Tb_Units As String, Tc As Double, Tc_Units As String, default_units As String) As Double
'Calculation for Heat of Vaporization (J/kmol)
'
'Method: Watson
'
'Equation Inputs: HVb (J/kmol) - Heat of Vaporization at NBP
'                 Tb  (K)      - Boiling Point
'                 Tc  (K)      - Critical Temperature
'
' Modified on 6/18/97 BGH : Added Unit parameters
    
    Dim Tr1 As Double
    Dim Tr2 As Double
    Dim tempHVb As Double
    Dim tempTb As Double
    Dim tempTc As Double
    Dim value As Double
    Dim method_units As String
    
    method_units = "J/kmol"
    
    On Error Resume Next
    
    'check inputs for correct units
    If Trim(HVb_Units) = "J/kmol" Then
        tempHVb = HVb
    Else
        tempHVb = Convert(HVb, HvapNBP, HVb_Units, "J/kmol", False)
    End If
    
    If Trim(Tb_Units) = "K" Then
        tempTb = Tb
    Else
        tempTb = Convert(Tb, NBP, Tb_Units, "K", False)
    End If
            
    If Trim(Tc_Units) = "K" Then
        tempTc = Tc
    Else
        tempTc = Convert(Tc, CT, Tc_Units, "K", False)
    End If
    
'compute value
    If tempTc = 0 Then
        tempTc = STANDARD_K_TEMP
    End If
    Tr1 = (tempTb / tempTc)
    Tr2 = (298.15 / tempTc)
    If (1 - Tr1) <> 0# Then
        value = HVb * ((1 - Tr2) / (1 - Tr1)) ^ 0.38
    Else
        CalcHvapWatson = ERROR_FLAG
        Exit Function
    End If
' check to make sure value is in default units
    If Trim(method_units) = Trim(default_units) Then
        CalcHvapWatson = value
    Else
        CalcHvapWatson = Convert(value, Hvap25, method_units, default_units, False)
    End If
    
End Function

Function CalcHvapNBPKlein(Tb As Double, Tb_Units As String, Pc As Double, Pc_Units As String, Tc As Double, Tc_Units As String, default_units As String) As Double
'Calculation for Heat of Vaporization at NBP (J/kmol)
'
'Method: Klein (1949)
'
'Equation Inputs: Tb (K)    - Boiling Point
'                 Pc (atm)  - Critical Pressure
'                 Tc (K)    - Critical Temperature
'
' Modified 6/19/97 BGH : Added Unit parameters

    Dim R As Double
    Dim Kkl As Double
    Dim temp1 As Double
    Dim temp2 As Double
    Dim Temp3 As Double
    Dim tempTb As Double
    Dim tempPc As Double
    Dim tempTc As Double
    Dim value As Double
    Dim method_units As String
    
    method_units = "J/kmol"
    On Error Resume Next
    
    'check inputs for correct units
    If Trim(Tb_Units) = "K" Then
        tempTb = Tb
    Else
        tempTb = Convert(Tb, NBP, Tb_Units, "K", False)
    End If
    
    If Trim(Tc_Units) = "K" Then
        tempTc = Tc
    Else
        tempTc = Convert(Tc, CT, Tc_Units, "K", False)
    End If
    
    If Trim(Pc_Units) = "atm" Then
        tempPc = Pc
    Else
        tempPc = Convert(Pc, CP, Pc_Units, "atm", False)
    End If

    'Set universal gas constant (cal/mol*K)
    R = 1.9872
    
    'Convert Pc to atm
    ' taken care of above BGH 6/19/97
    'Pc = Pc / 101325

    'Set the Klein Constant based on Tb
    If tempTb < 200 Then
        Kkl = 1.02
    ElseIf tempTb <= 300 Then
        Kkl = 1.04
    ElseIf tempTb > 300 Then
        Kkl = 1.045
    End If
    ' try to go ahead with the calculation even if we don't have these
    If tempTc = 0# Then
        tempTc = STANDARD_K_TEMP
    End If
    If tempPc <= 0 Then
        tempPc = STANDARD_ATM_PRESSURE
    End If
    Temp3 = (1 - 1 / (tempPc * (tempTb / tempTc)) ^ 3) ^ 0.5
    temp2 = 1 - (tempTb / tempTc)
    temp1 = R * Kkl * tempTb * ln(tempPc)
    
    If temp2 <> 0 Then
        value = (temp1 * Temp3 / temp2) * 4183.9254
    Else
        CalcHvapNBPKlein = ERROR_FLAG
        Exit Function
    End If
    
    ' make sure answer is in correct units
    If Trim(method_units) = Trim(default_units) Then
        CalcHvapNBPKlein = value
    Else
        CalcHvapNBPKlein = Convert(value, HvapNBP, method_units, default_units, False)
    End If

    'Convert Pc to Pa
    ' Pc was put in a temp_variable so it has not been changed
    'Pc = Pc * 101325

End Function


Function CalclogKowHansch(S As Double, S_Units As String, default_units As String) As Double
 'Calculation for log(Octanol/Water Partition Coefficient) (unit-less)
'
'Method: Hansch (1968)
'
'Equation Inputs: S (mol/L)     - Solubility of chemical in water
'
' Modified 6/25/97 BGH: added unit parameters

' DENISE check the units for S since we updated default units for Sol

    'Dim M As Double
    Dim tempS As Double
    Dim value As Double
    Dim method_units As String
    
    method_units = "unit-less"
    On Error Resume Next
       
    ' for now just convert S back to mol/L
    'M = infomethod(MW).value(infomethod(MW).CurMethod)
    'S = S / (M * 1000)
    ' should be taken care of by following if/convert block
    
    If Trim(S_Units) = "mol/L" Then
        tempS = S
    Else
        tempS = Convert(S, Schem, S_Units, "mol/L", False)
    End If
    
    If tempS <= 0 Then
        CalclogKowHansch = ERROR_FLAG
        Exit Function
    End If
    If (1 / tempS) <= 0 Then
        CalclogKowHansch = ERROR_FLAG
        Exit Function
    End If
    value = 0.747 * log10(1 / tempS) + 0.73
    
    If Trim(method_units) = Trim(default_units) Then
        CalclogKowHansch = value
    Else
        CalclogKowHansch = Convert(value, logKow, method_units, default_units, False)
    End If
    
    'S = S * M * 1000
End Function
Function CalclogKocBaker(log10Kow As Double, default_units As String) As Double
 'Calculation for log(Organic Carbon/Water Partition Coefficient) (cm3/g OC)
'
'Method: Baker (1994)
'
'Input Equations: log10Kow (unit-less)  - log(Octanol/Water Partition Coefficient)
'
' Modified 6/25/97 BGH: added unit parameter

Dim value As Double
Dim method_units As String

method_units = "cm3/g OC"

    On Error Resume Next
    
    value = 0.904 * log10Kow + 0.086

    'make sure answer is in correct units
    If Trim(method_units) = Trim(default_units) Then
        CalclogKocBaker = value
    Else
        CalclogKocBaker = Convert(value, logKoc, method_units, default_units, False)
    End If
    
End Function

Function CalclogKowUNIFAC(T As Double, T_Units As String, BIPCode As Integer, default_units As String) As Double
'Calculation for log Kow (unit-less)
'
'Method: UNIFAC
'
'Equation Inputs: T (C)     - Temperature
'                 BIPCode   - BIP database
'
' Modified 6/24/97 BGH: added unit parameters

    Dim i As Integer
    Dim j As Integer
    Dim Kow As Double
    Dim KowSS As Long
    Dim KowLS As Long
    Dim KowErr As Long
    Dim KowTemp As Double
    Dim FGRPErr As Long
    Dim MX As Long
    Dim AI(1 To 58, 1 To 58) As Double
    Dim MST(1 To 10, 1 To 10, 1 To 2) As Long
    Dim value As Double
    Dim method_units As String
    Dim tempT As Double
    
    method_units = "unit-less"
    
    On Error Resume Next
    
    'check to see if inputs are in correct units
    If Trim(T_Units) = "C" Then
        tempT = T
    Else
        tempT = Convert(T, OptTemp, T_Units, "C", False)
    End If
    
    'Convert T to C
    'T = T - 273.15
    
    MX = Cur_Info.MaxGroups
    
    If MX <= 0 Then Exit Function

    For i = 1 To 10
        For j = 1 To 10
            MST(i, j, 1) = 0
            MST(i, j, 2) = 0
        Next j
    Next i

    For i = 1 To MX
        MST(2, i, 1) = Cur_Info.Grp(i)
        MST(2, i, 2) = Cur_Info.NumGrp(i)
    Next i
    For j = i To 10
        MST(2, i, 1) = 0
        MST(2, i, 2) = 0
    Next j

    Kow = 0     'Returned Value
    KowSS = 0   'Not Important
    KowLS = 0   'Not Important
    KowErr = 0  'Not Important
    KowTemp = 0 'Not Important
    FGRPErr = 0 'Not Important
    
    If BIPCode <> 4 Then
        For i = 1 To 58
            For j = 1 To 58
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    Else
        For i = 1 To 32
            For j = 1 To 32
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    End If
    
    Call KOWCALL(Kow, KowSS, KowLS, KowErr, KowTemp, tempT, FGRPErr, MX, MST(1, 1, 1), MGSG(1), AI(1, 1), RI(1), QI(1), MWS(1), MVS(1))
    
    If KowErr = -1 Or Kow = 0 Then
        CalclogKowUNIFAC = ERROR_FLAG
        Exit Function
    Else
        value = Kow
        'make sure answer is in correct units
        If Trim(default_units) = Trim(method_units) Then
            CalclogKowUNIFAC = value
        Else
            CalclogKowUNIFAC = Convert(value, logKow, method_units, default_units, False)
        End If
        
    End If
    
    'Convert T to K
    'T = T + 273.15
    
End Function

Function CalclogKowKenagaGoring(S As Double, S_Units As String, default_units As String) As Double
'Calculation for log(Octanol/Water Partition Coefficient) (unit-less)
'
'Method: Kenaga and Goring (1978)
'
'Equation Inputs: M (g/mol)     - Molecular Weight ** not used in method,
'                                   removed 6/25/97 BGH **
'                 S (ppm(wt))   - Solubility of chemical in water
'
' Modified 6/25/97 BGH: Added Unit parameters

    ' DENISE check units of S since we changed default
    On Error Resume Next
    
'Dim tempM As Double
Dim tempS As Double
Dim value As Double
Dim method_units As String

method_units = "unit-less"

'make sure inputs are in correct units
'If Trim(M_Units) = "g/mol" Then
'    tempM = M
'Else
'    tempM = Convert(MW, M_Units, "g/mol", M)
'End If

If Trim(S_Units) Like "ppm*" Then
    tempS = S
Else
    tempS = Convert(S, Schem, S_Units, "ppm(wt)", False)
End If

    
    'Convert S to ppm(wt) (no need to because Solubility now default in ppm)
    'S = S * M * 1000
    If tempS <= 0 Then
        CalclogKowKenagaGoring = ERROR_FLAG
        Exit Function
    End If
    value = -1.085 * log10(tempS) + 4.538
    
    'make sure answer is in correct units
    If Trim(method_units) = Trim(default_units) Then
        CalclogKowKenagaGoring = value
    Else
        CalclogKowKenagaGoring = Convert(value, logKow, method_units, default_units, False)
    End If
    
    'Convert S to kmol/m3 water
    'S = S / (M * 1000)
    
End Function
Function CalcLDBhiruds(M As Double, M_Units As String, Pc As Double, Pc_Units As String, T As Double, Tb As Double, Tb_Units As String, Tc As Double, Tc_Units As String, method_units As String) As Double

'Calculation for Liquid Density (kmol/m3)
'
'Method: Bhirud's (1978)
'
'Equation Inputs: M  (g/mol)    - Molecular Weight
'                 Pc (atm)      - Critical Pressure
'                 T  (K)        - Temperature
'                 Tb (K)        - Boiling Point
'                 Tc (K)        - Critcal Temperature
'
' REVISIONS:  6/6/97  DMW  :  added a unit parameter since the units used here are no
'                               longer the default units
'   added M_Units, Pc_Units, Tb_Units, Tc_Units so we could check to see if they where
'       in the correct units for this method. BGH 6/17/97
'               5/20/98  DMW  : added checks for divide by zero
    Dim calc_units As String
    Dim R As Double
    Dim Tr As Double
    Dim a As Double
    Dim b As Double
    Dim w As Double
    Dim result As Double
    Dim tempPc As Double
    Dim tempM As Double
    Dim tempTb As Double
    Dim tempTc As Double
    Dim TempValue As Double
    calc_units = "kmol/m3"
    On Error Resume Next

    'Check to make sure inputs to equation are in
    'correct units. BGH 6/17/97
    
    ' M_Units should be in g/mol, if not, convert to g/mol
    If Trim(M_Units) = "g/mol" Then
        tempM = M
    Else
        tempM = Convert(M, MW, M_Units, "g/mol", False)
    End If
        
    ' Pc_Units should be in atm, if not, convert to atm
    If Trim(Pc_Units) = "atm" Then
        tempPc = Pc
    Else
        
            tempPc = Convert(Pc, CP, Pc_Units, "atm", False)
       
    End If
    
    ' Check Tb_Units = K
    If Trim(Tb_Units) = "K" Then
        tempTb = Tb
    Else
        tempTb = Convert(Tb, OptTemp, Tb_Units, "K", False)
    End If
    
    ' Check Tc_Units = K
    If Trim(Tc_Units) = "K" Then
        tempTc = Tc
    Else
        tempTc = Convert(Tc, CT, Tc_Units, "K", False)
    End If
    ' make sure tempTc isn't 0, if nothing else put in standard K temp
    If tempTc = 0 Then
        tempTc = STANDARD_K_TEMP
    End If
    'Set universal gas constant (cm3*atm/mol*K)
    R = 82.04

    'Convert Pc to atm (and put in temp location for function)
    'Modified 6/17/97 BGH : conversion should have already taken place
    'tempPc = Pc / 101325

    Tr = T / tempTc
    
    a = 1.39644 - 24.076 * Tr + 102.615 * Tr ^ 2 - 255.719 * Tr ^ 3 + 355.805 * Tr ^ 4 - 256.671 * Tr ^ 5 + 75.1088 * Tr ^ 6
    b = 13.4412 - 135.7437 * Tr + 533.38 * Tr ^ 2 - 1091.453 * Tr ^ 3 + 1231.43 * Tr ^ 4 - 728.227 * Tr ^ 5 + 176.737 * Tr ^ 6
    
    w = CalcAFPitzer(tempTc, "K", tempTb, "K", tempPc, "atm")    'Calculates Acentric Factor (unit-less)
    TempValue = (R * T * Exp(a + w * b))
    If TempValue <> 0 Then
        result = (tempM * tempPc / TempValue) * 1000
    Else
        '  indicate failure
        CalcLDBhiruds = ERROR_FLAG
        Exit Function
    End If
    'Convert Pc to Pa (why are we doing this here? DENISE check this)
    'Pc = Pc * 101325 (made it tempPc so it would not have to be changed back. BGH 6/16/97
       
    'now if calc_units not same as default units, do conversion
        If Trim(method_units) <> Trim(calc_units) Then
            CalcLDBhiruds = Convert(result, LD, calc_units, method_units, False)
        Else
            CalcLDBhiruds = result
        End If
   
   
End Function

Function CalcLDModRackett(M As Double, M_Units As String, Tb As Double, Tb_Units As String, Tc As Double, Tc_Units As String, Pc As Double, Pc_Units As String, T As Double, method_units As String) As Double
'Calculation for Liquid Density (kg/m3)
'
'Method: Modified Rackett (1978)
'
'Equation Inputs: M  (g/mol)    - Molecular Weight
'                 Tb (K)        - Boiling Point
'                 Tc (K)        - Critical Temperature
'                 Pc (Pa)       - Critical Pressure
'                 T  (K)        - Temperature
'
' REVISIONS:  6/6/97  DMW  :  added a unit parameter since the units used here are no
'                               longer the default units
'       6/17/97 BGH : added unit patameters to all inputs
'       5/20/98  DMW  : added checks for divide by zero
    Dim tempM As Double
    Dim tempTb As Double
    Dim tempTc As Double
    Dim tempPc As Double
    Dim R As Double
    Dim w As Double
    Dim Tr As Double
    Dim Zra As Double
    Dim result As Double
    Dim ExpValue As Double
    Dim calc_units As String
    Dim temp_value As Double
    
    calc_units = "kg/m3"
    On Error Resume Next
    
    'Check for incorrect units in the inputs 6/17/97 BGH
    ' M_Units should be in g/mol, if not, convert to g/mol
    If Trim(M_Units) = "g/mol" Then
        tempM = M
    Else
        tempM = Convert(M, MW, M_Units, "g/mol", False)
    End If
        
    ' Pc_Units should be in Pa, if not, convert to Pa
    If Trim(Pc_Units) = "Pa" Then
        tempPc = Pc
    Else
        tempPc = Convert(Pc, CP, Pc_Units, "Pa", False)
    End If
    
    ' Check Tb_Units = K
    If Trim(Tb_Units) = "K" Then
        tempTb = Tb
    Else
        tempTb = Convert(Tb, NBP, Tb_Units, "K", False)
    End If
    
    ' Check Tc_Units = K
    If Trim(Tc_Units) = "K" Then
        tempTc = Tc
    Else
        tempTc = Convert(Tc, CT, Tc_Units, "K", False)
    End If

    If tempTc = 0 Then
        tempTc = STANDARD_K_TEMP
    End If
    
    'Set universal gas constant (Pa*m3/kmol*K)
    R = 8314
                   
    Tr = T / tempTc ' Fix CalcAFPitzer Later, Pc needs to be in atm for it to work now
    w = CalcAFPitzer(tempTc, "K", tempTb, "K", tempPc, "Pa")
    Zra = 0.29056 - 0.08775 * w
        
    ExpValue = 1 + (1 - Tr) ^ (2 / 7)
    
    'Equation modified by BGH 6/17/97 : missing MW in calculation
    temp_value = (tempPc * tempM)
    If temp_value <> 0 Then
        temp_value = (((R * tempTc) / temp_value) * Zra ^ ExpValue)
        If temp_value <> 0 Then
            result = 1 / temp_value
        Else
            CalcLDModRackett = ERROR_FLAG
            Exit Function
        End If
    Else
        CalcLDModRackett = ERROR_FLAG
        Exit Function
    End If
    
    'now if calc_units not same as default units, do conversion
        If Trim(method_units) <> Trim(calc_units) Then
           CalcLDModRackett = Convert(result, LD, calc_units, method_units, False)
        Else
            CalcLDModRackett = result
        End If
   
End Function

Function CalcDwaterHaydukLaudie(method_units As String) As Double
'Calculation for Liquid Diffusivity (cm2/s)
'
'Method: Hayduk and Laudie
'
'Equation Inputs: None
'
' REVISIONS:  DMW  6/15/97  added unit parameter to check the units this is returned in
    
    Dim MV As Double
    Dim sub_value As Double
    Dim sub_units As String
    
    sub_units = "cm2/s"
    
    On Error Resume Next
    
    MV = CalcMVNBPSchroeder("cm3/mol")
    
    If MV <> 0 Then
        sub_value = 0.0001326 / MV ^ 0.589
        If Trim(method_units) <> sub_units Then
            sub_value = Convert(sub_value, Dwater, sub_units, method_units, False)
        End If
    Else
        sub_value = 0
    End If
    CalcDwaterHaydukLaudie = sub_value
End Function

Function CalcDwaterPolson(M As Double, Munits As String, method_units As String) As Double
'Calculation for Liquid Diffusivity (cm2/s)
'
'Method: Polson
'
'Equation Inputs: M (g/mol)     - Molecular Weight
'
' REVISIONS:  DMW  6/15/97  added unit parameter to check the units this is returned in
    Dim tempM As Double
    Dim sub_value As Double
    Dim sub_units As String
    
    sub_units = "cm2/s"
    If Trim(Munits) <> "g/mol" Then
        tempM = Convert(M, MW, Munits, "g/mol", False)
    Else
        tempM = M
    End If
    On Error Resume Next

    sub_value = 0.0000274 * (tempM ^ (-1 / 3))
    
    If sub_units <> Trim(method_units) Then
        sub_value = Convert(sub_value, Dwater, sub_units, method_units, False)
    End If
    CalcDwaterPolson = sub_value
    
End Function

Function CalcDwaterWilkeChang(T As Double, TUnits As String, method_units As String) As Double
'Calculation for Liquid Diffusivity (cm2/s)
'
'Method: Wilke and Chang
'
'Equation Inputs: T (K)     - Temperature
'
' REVISIONS:  DMW  6/15/97  added unit parameter to check the units this is returned in
    Dim MV As Double
    Dim sub_value As Double
    Dim tempT As Double
    Dim sub_units As String
    
    sub_units = "cm2/s"
    On Error Resume Next
    If Trim(TUnits) <> "K" Then
        tempT = Convert(T, OptTemp, TUnits, "K", False)
    Else
        tempT = T
    End If
    
    MV = CalcMVNBPSchroeder("cm3/mol")
    
    If MV <> 0 Then
        sub_value = (7.4856519E-16 * tempT * MV ^ 0.6 / 0.001) * 10000
        If Trim(method_units) <> sub_units Then
            sub_value = Convert(sub_value, Dwater, sub_units, method_units, False)
        End If
    Else
        sub_value = 0
    End If
    CalcDwaterWilkeChang = sub_value
End Function

Function CalcMVNBPSchroeder(method_units As String)
'Calculation for Molar Volume at NBP (cm3/mol)
'
'Method: Schroeder
'
'Equation Inputs: Group contribution fragments
'
' REVISIONS :  DMW  6/17/97  added a conversion at end so value is returned in requested units (method_units)
   
    Dim i As Integer
    Dim MV As Double
    Dim result As Double
    
    On Error Resume Next

    For i = 1 To Cur_Info.MaxGroups
        MV = MV + MVS(Cur_Info.Grp(i)) * Cur_Info.NumGrp(i)
    Next i

    result = (MV - (Cur_Info.NumRings * 7)) / 1000

    ' convert to pearls default units if necessary
    If Trim(method_units) <> "cm3/mol" Then
        result = Convert(result, OptPress, "cm3/mol", method_units, False) 'FIX, -2 is nonsense, but there is no code for MVNBP
    End If
    CalcMVNBPSchroeder = result
End Function

Function CalcMWUNIFAC(default_units As String)

'Calculation for Molecular Weight (kg/kmol)
'
'Method: UNIFAC
'
'Equation Inputs: Group contribution fragments
'
'Modified 6/17/97 BGH: added defualt_units parameter so
'   that we could make sure the computed value was in the
'   correct units

    Dim i As Integer
    Dim M As Double
    Dim method_units As String
    
    method_units = "kg/kmol"
    On Error Resume Next

    M = 0

    For i = 1 To Cur_Info.MaxGroups
        M = M + MWS(Cur_Info.Grp(i)) * Cur_Info.NumGrp(i)
    Next i
    
    ' If the Method's result is not in the default units
    ' convert them. *Convert takes parameters (units_from,
    ' units_to, Value)
    If M = 0 Then
        CalcMWUNIFAC = ERROR_FLAG
        Exit Function
    End If
    
    If Trim(method_units) = Trim(default_units) Then
        CalcMWUNIFAC = M
    Else
        CalcMWUNIFAC = Convert(M, MW, method_units, default_units, False)
    End If
End Function

Function CalcAFPitzer(Tc As Double, Tc_Units As String, Tb As Double, Tb_Units As String, Pc As Double, Pc_Units As String) As Double
'Calculate Acentric Factor (unit-less)
'
'Method: Pitzer
'
'Equation Inputs: Tc (K)    - Critical Temperature
'                 Tb (K)    - Boiling Point
'                 Pc (atm)  - Critical Pressure
'
' Modified 6/25/97 BGH: added unit parameters

    Dim Tbr As Double
    Dim tempTc As Double
    Dim tempTb As Double
    Dim tempPc As Double
    Dim TempValue As Double
    On Error Resume Next
    
    'make sure inputs are in correct units
    If Trim(Tc_Units) = "K" Then
        tempTc = Tc
    Else
        tempTc = Convert(Tc, CT, Tc_Units, "K", False)
    End If
    
    If Trim(Tb_Units) = "K" Then
        tempTb = Tb
    Else
        tempTb = Convert(Tb, NBP, Tb_Units, "K", False)
    End If
    
    If Trim(Pc_Units) = "atm" Then
        tempPc = Pc
    Else
        tempPc = Convert(Pc, CP, Pc_Units, "atm", False)
    End If
    ' make sure TempTc isn't 0, if so make it standard temp in K
    If tempTc = 0 Then
        tempTc = STANDARD_K_TEMP
    End If
    Tbr = tempTb / tempTc
    TempValue = 1 - Tbr
    If TempValue <> 0 And tempPc > 0 Then
        CalcAFPitzer = (3 / 7) * (Tbr / TempValue) * log10(tempPc) - 1
    Else
        CalcAFPitzer = 0
    End If

End Function

Function CalcSchemUNIFAC(T As Double, TUnits As String, M As Double, BIPCode As Integer, method_units As String) As Double
'Calculation for Solubility in Water (kmol/m3 water)
'
'Method: UNIFAC
'
'Equation Inputs: T (C)         - Temperature
'                 M (g/mol)     - Molecular Weight
'                 BIPCode       - BIP database
'
' REVISIONS:  6/6/97  DMW  :  added a unit parameter since the units used here are no
'                               longer the default units

    Dim i As Integer
    Dim j As Integer
    Dim Sol As Double   ' the solubility
    Dim SOLSS As Long
    Dim SOLLS As Long
    Dim SOLErr As Long
    Dim SOLTemp As Double
    Dim MX As Long
    Dim AI(1 To 58, 1 To 58) As Double
    Dim MST(1 To 10, 1 To 10, 1 To 2) As Long
    Dim XMW(1 To 10) As Double
    Dim result As Double
    Dim sub_units As String
    Dim tempT As Double
    
    
    ' DENISE check should this have water?
    sub_units = "kmol/m3"
    On Error Resume Next
    
    'Convert T to C
    tempT = T
    If TUnits <> "C" Then
        tempT = Convert(tempT, OptTemp, TUnits, "C", False)
    End If
    
    MX = Cur_Info.MaxGroups
    
    If MX <= 0 Then
        CalcSchemUNIFAC = ERROR_FLAG
        Exit Function
    End If
    XMW(1) = 18.02
    XMW(2) = M

    For i = 3 To 10
        XMW(i) = 0
    Next i

    For i = 1 To 10
        For j = 1 To 10
            MST(i, j, 1) = 0
            MST(i, j, 2) = 0
        Next j
    Next i

    For i = 1 To 10
        MST(2, i, 1) = Cur_Info.Grp(i)
        MST(2, i, 2) = Cur_Info.NumGrp(i)
    Next i

    Sol = 0     'Returned Value
    SOLSS = 0   'Not Important
    SOLLS = 0   'Not Important
    SOLErr = 0  'Not Important
    SOLTemp = 0 'Not Important
    
    If BIPCode <> 4 Then
        For i = 1 To 58
            For j = 1 To 58
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    Else
        For i = 1 To 32
            For j = 1 To 32
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    End If

    Call AQSCALL(Sol, SOLSS, SOLLS, SOLErr, SOLTemp, tempT, MX, MST(1, 1, 1), XMW(1), MGSG(1), AI(1, 1), RI(1), QI(1), MWS(1), MVS(1))

    If SOLErr = -1 Then
        result = 0
        CalcSchemUNIFAC = ERROR_FLAG
        Exit Function
    Else
        If M <> 0 Then
            result = Sol / (M * 1000)
        Else
            CalcSchemUNIFAC = ERROR_FLAG
            Exit Function
        End If
    End If
    
    'Convert C To T
'    T = T + 273.15
    ' do the conversion if necessary to return in correct units
    CalcSchemUNIFAC = Convert(result, Schem, sub_units, Get_DefaultUnit(Schem), False)

End Function

Function CalcSwaterUNIFAC(T As Double, M As Double, BIPCode As Integer, method_units As String) As Double

'Calculation for Solubility in Chemical (kmol/m3 chem)
'
'Method: UNIFAC
'
'Equation Inputs: T (C)         - Temperature
'                 M (g/mol)     - Molecular Weight
'                 BIPCode       - BIP database
'
' REVISIONS:  6/6/97  DMW  :  added a unit parameter since the units used here are no
'                               longer the default units

    Dim i As Integer
    Dim j As Integer
    Dim Sol As Double
    Dim SOLSS As Long
    Dim SOLLS As Long
    Dim SOLErr As Long
    Dim SOLTemp As Double
    Dim MX As Long
    Dim AI(1 To 58, 1 To 58) As Double
    Dim MST(1 To 10, 1 To 10, 1 To 2) As Long
    Dim XMW(1 To 10) As Double
    Dim result As Double
    Dim sub_units As String
    Dim tempT As Double
    
    sub_units = "kmol/m3"
    On Error Resume Next
    
    'Convert T to C
    tempT = T
    If Cur_Info.OpTUnit <> "C" Then
        tempT = Convert(tempT, OptTemp, Cur_Info.OpTUnit, "C", False)
    End If
    
    MX = Cur_Info.MaxGroups
    
    If MX <= 0 Then Exit Function
   
    XMW(1) = 18.02
    XMW(2) = M

    For i = 3 To 10
        XMW(i) = 0
    Next i

    For i = 1 To 10
        For j = 1 To 10
            MST(i, j, 1) = 0
            MST(i, j, 2) = 0
        Next j
    Next i

    For i = 1 To 10
        If Cur_Info.Grp(i) <> 0 And Cur_Info.Grp(i) <> -1 Then
            MST(2, i, 1) = Cur_Info.Grp(i)
            MST(2, i, 2) = Cur_Info.NumGrp(i)
        End If
    Next i

    Sol = 0     'Returned Value
    SOLSS = 0   'Not Important
    SOLLS = 0   'Not Important
    SOLErr = 0  'Not Important
    SOLTemp = 0 'Not Important

    If BIPCode <> 4 Then
        For i = 1 To 58
            For j = 1 To 58
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    Else
        For i = 1 To 32
            For j = 1 To 32
                AI(i, j) = BIP(BIPCode, i, j)
            Next j
        Next i
    End If
    
    Call AQSCALL2(Sol, SOLSS, SOLLS, SOLErr, SOLTemp, tempT, MX, MST(1, 1, 1), XMW(1), MGSG(1), AI(1, 1), RI(1), QI(1), MWS(1), MVS(1))
   
    If SOLErr = -1 Then
        result = 0
        CalcSwaterUNIFAC = ERROR_FLAG
        Exit Function
    Else
        If M <> 0# Then
            result = Sol / (M * 1000)
        Else
            CalcSwaterUNIFAC = ERROR_FLAG
            Exit Function
        End If
    End If
    
    'Convert T to K
    'T = T + 273.15
    CalcSwaterUNIFAC = Convert(result, Swater, sub_units, Get_DefaultUnit(Swater), False)
End Function

Function CalcSchemYalkowsky(Tmp As Double, TmpUnits As String, log10Kow As Double, KowUnits As String, T As Double, TUnits As String, method_units As String) As Double
'Calculation for Solubility in Water (mol/L water)
'
'Method: Yalkowsky (1990)
'
'Equation Inputs: Tmp      (C)          - Melting Point
'                 log10Kow (unit-less)  - log(Octanol/Water Partition Coefficient)
'                 T        (C)          - Temperature
'
' REVISIONS:  6/6/97  DMW  :  added a unit parameter since the units used here are no
'                               longer the default units
        
    Dim tempT As Double
    Dim tempTmp As Double
    Dim tempKow As Double
    Dim result As Double
    Dim sub_units As String
    
    ' DENISE check should this have water?
    sub_units = "mol/L"
    On Error Resume Next
    
    'Convert T and Tmp to C and logKow to unit-less
    
    tempTmp = Convert(Tmp, OptTemp, TmpUnits, "C", False)
    tempT = Convert(T, OptTemp, TUnits, "C", False)
    tempKow = Convert(log10Kow, logKow, KowUnits, "unit-less", False)
    
    If tempTmp < tempT Then tempT = 25
    
    result = 10 ^ (-0.01 * tempTmp - tempKow + 0.8)

    If result = 0# Then
        CalcSchemYalkowsky = ERROR_FLAG
        Exit Function
    End If
    ' do the conversion if necessary
    CalcSchemYalkowsky = Convert(result, Schem, sub_units, Get_DefaultUnit(Schem), False)

End Function

Function CalcSTBrockBird(Pc As Double, PcUnit As String, Tc As Double, TcUnit As String, Tb As Double, TbUnit As String, T As Double, TUNIT As String, method_units As String) As Double
'Calculation for Surface Tension (N/m)
'
'Method: Brock and Bird (1983)
'
'Equation Inputs: Pc (Pa)   - Critical Pressure
'                 Tc (K)    - Critical Temperature
'                 Tb (K)    - Boiling Point
'                 T  (K)    - Temperature
'
    
    Dim q As Double
    Dim Tr As Double
    Dim Tbr As Double
    Dim TempVal As Double

    Dim tempPc As Double
    Dim tempTc As Double
    Dim tempTb As Double
    Dim tempT As Double
    Dim sub_value As Double
    Dim sub_units As String
    Dim temp_value As Double
    
    On Error Resume Next
        
    If PcUnit <> "Pa" Then
        tempPc = Convert(Pc, CP, PcUnit, "Pa", False)
    Else
        tempPc = Pc
    End If
    If TcUnit <> "K" Then
        tempTc = Convert(Tc, CT, TcUnit, "K", False)
    Else
        tempTc = Tc
    End If
    If TbUnit <> "K" Then
        tempTb = Convert(Tb, NBP, TbUnit, "K", False)
    Else
        tempTb = Tb
    End If
    sub_units = "N/m"
    
    If tempTc = 0# Then
        tempTc = STANDARD_K_TEMP
    End If
    Tr = tempT / tempTc
    Tbr = tempTb / tempTc
    temp_value = 1 - Tbr
    If temp_value <> 0# And tempPc >= 0# Then
        q = 0.1207 * (1 + Tbr * (ln(tempPc) - 11.526) / (1 - Tbr)) - 0.281
    Else
        CalcSTBrockBird = ERROR_FLAG
        Exit Function
    End If
             
    sub_value = (0.00046 * (tempPc ^ (2 / 3) * tempTc ^ (1 / 3) * q * (1 - Tr) ^ (11 / 9))) / 1000
    If Trim(method_units) <> sub_units Then
        sub_value = Convert(sub_value, ST, sub_units, method_units, False)
    End If
    CalcSTBrockBird = sub_value
End Function




Function CalcHvapNBPChenPitzer(Tb As Double, Tb_Units As String, Tc As Double, Tc_Units As String, Pc As Double, Pc_Units As String, default_units As String)
'Calculation for Heat of Vaporization at NBP (J/kmol)
'
'Method: Chen and Pitzer (1965)
'
'Equation Inputs: Tb (K)    - Boiling Point
'                 Tc (K)    - Critical Temperature
'                 Pc (atm)  - Critical Pressure
'
' Modified 6/19/97 BGH: Added unit parameters

    Dim Tbr As Double
    Dim tempTb As Double
    Dim tempTc As Double
    Dim tempPc As Double
    Dim value As Double
    Dim method_units As String
    
    ' check to make sure inputs are in corrrect units
    If Trim(Tb_Units) = "K" Then
        tempTb = Tb
    Else
        tempTb = Convert(Tb, NBP, Tb_Units, "K", False)
    End If
    
    If Trim(Tc_Units) = "K" Then
        tempTc = Tc
    Else
        tempTc = Convert(Tc, CT, Tc_Units, "K", False)
    End If
    
    If Trim(Pc_Units) = "atm" Then
        tempPc = Pc
    Else
        tempPc = Convert(Pc, CP, Pc_Units, "atm", False)
    End If
    
    method_units = "J/kmol"
    On Error Resume Next
    
    'Convert Pc to atm
    ' taken care of above
    'Pc = Pc / 101325
    If tempTc = 0# Then
        tempTc = STANDARD_K_TEMP
    End If
    Tbr = tempTb / tempTc
    If (tempPc > 0) And (1.07 - Tbr > 0) Then
        value = (tempTb * (7.11 * log10(tempPc) - 7.82 + 7.9 * Tbr) / (1.07 - Tbr)) * 4183.9254
    Else
        CalcHvapNBPChenPitzer = 0#
        Exit Function
    End If
    
    'check to see if answer is in correct units
    If Trim(method_units) = Trim(default_units) Then
        CalcHvapNBPChenPitzer = value
    Else
        CalcHvapNBPChenPitzer = Convert(value, HvapNBP, method_units, default_units, False)
    End If
        
    
    'Convert Pc to Pa
    'Pc = Pc * 101325
    
End Function

