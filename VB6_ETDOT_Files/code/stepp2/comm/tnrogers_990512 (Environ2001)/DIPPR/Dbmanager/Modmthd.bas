Attribute VB_Name = "MODMthCalcs"
Option Explicit

Function ln(Number As Double) As Double
'Returns the natural log of a number
'NOTE: In VB log = ln
    
    If (Number = 0) Then Exit Function
    ln = Log(Abs(Number))
End Function

Function log10(Number As Double) As Double
'Return the log base 10 of a number
'NOTE: In VB log = ln

    If Number <= 0 Then Exit Function
    log10 = Log(Number) / Log(10#)
End Function


Sub CalcBCFKobayshi(log10Kow As Double, Value As Double)
' Calculation for Bioconcentration Factor (unit-less)
'
' Method: Kobayshi (1981)
' Equation Inputs: log10Kow (unit-less) - log(Octanol/Water Partition Coefficient)
' Modified 6/25/97 BGH: added unit parameters

Dim method_units As String
    On Error Resume Next
    method_units = "unit-less"
    Value = 10 ^ (0.74 * log10Kow - 0.77)
    
    'make sure answer is in correct units
    'If Trim(method_units) = Trim(default_units) Then
    '    CalcBCFKobayshi = value
    'Else
    '    CalcBCFKobayshi = Convert(BCF, method_units, default_units, value)
    'End If
End Sub

Sub CalclogKocBaker(log10Kow As Double, Value As Double)
' Calculation for log(Organic Carbon/Water Partition Coefficient) (cm3/g OC)
'
' Method: Baker (1994)
' Input Equations: log10Kow (unit-less) - log(Octanol/Water Partition Coefficient)
' Modified 6/25/97 BGH: added unit parameter

Dim method_units As String
    On Error Resume Next
    method_units = "cm3/g OC"
    Value = 0.904 * log10Kow + 0.086

    'make sure answer is in correct units
    'If Trim(method_units) = Trim(default_units) Then
    '    CalclogKocBaker = value
    'Else
    '    CalclogKocBaker = Convert(logKoc, method_units, default_units, value)
    'End If
End Sub

Sub CalcACwaterUNIFAC(T As Double, Value As Double)
' Calculation for Infinite Dilution Activity Coefficient of Water in Chemical (unit-less)
'
' Method: UNIFAC
' Equation Inputs:
'       T (C)     - Temperature
'       BIPCode   - BIP database
' Modified 1/3/98 for dbman from PEARLS

    Dim i As Integer
    Dim j As Integer
    Dim GAM As Double
    Dim GAMSS As Long
    Dim GAMLS As Long
    Dim GAMErr As Long
    Dim GAMTemp As Double
    Dim FGRPErr As Long
    Dim MX As Long
    Dim Ai(1 To 58, 1 To 58) As Double
    Dim MST(1 To 10, 1 To 10, 1 To 2) As Long
    Dim tempT As Double
    Dim method_units As String
    Dim MGSG(1 To 116) As Long
    Dim RI(1 To 116) As Double
    Dim QI(1 To 116) As Double
    Dim MWS(1 To 116) As Double
    Dim MVS(1 To 116) As Double
    method_units = "unit-less"

    On Error Resume Next
    
    'make sure inputs are in correct units
    'If T_Units = "C" Then
        tempT = T
    'Else
    '    tempT = simple_t_convert(T_Units, "C", T)
    'End If
    
    'Convert T to C
    'T = T - 273.15
    
    'MX = Cur_Info.MaxGroups
            
    'If MX <= 0 Then Exit Function

    For i = 1 To 10
        For j = 1 To 10
            MST(i, j, 1) = 0
            MST(i, j, 2) = 0
        Next j
    Next i

    For i = 1 To 10
        If num_cur_chem_groups(i - 1) > 0 Then
            MST(2, i, 1) = cur_chem_groups(i - 1)
            MST(2, i, 2) = num_cur_chem_groups(i - 1)
        Else
            Exit For
        End If
    Next i
    MX = i
    
    GAM = 0     'Returned Value
    GAMSS = 0   'Not Important
    GAMLS = 0   'Not Important
    GAMErr = 0  'Not Important
    GAMTemp = 0 'Not Important
    FGRPErr = 0 'Not Important
    
    For i = 0 To 3
        If load_BIP_data(BIPCode(i), Ai()) = True Then
            Exit For
        End If
    Next i
    
    If load_UNIFAC_data(MGSG(), RI(), QI(), MWS(), MVS()) = False Then
        MsgBox ("UNIFAC data not available")
        Exit Sub
    End If
    
    '9999 --- This doesn't work ---
    'Call ACCALL2(GAM, GAMSS, GAMLS, GAMErr, GAMTemp, tempT, FGRPErr, MX, MST(1, 1, 1), MGSG(1), Ai(1, 1), RI(1), QI(1), MWS(1), MVS(1))

    If GAMErr <> -1 Then
        Value = GAM
    
    End If

    'Convert T to K
    'T = T + 273.15
    
    Exit Sub
    
End Sub

Sub CalcSwaterUNIFAC(T As Double, M As Double, Value As Double)
' Calculation for Solubility in Chemical (kmol/m3 chem)
'
' Method: UNIFAC
' Equation Inputs:
'       T (C)         - Temperature
'       M (g/mol)     - Molecular Weight
'       BIPCode       - BIP database
' REVISIONS:  6/6/97  DMW :
'       added a unit parameter since the units used here are no
'       longer the default units

    Dim i As Integer
    Dim j As Integer
    Dim Sol As Double
    Dim SOLSS As Long
    Dim SOLLS As Long
    Dim SOLErr As Long
    Dim SOLTemp As Double
    Dim MX As Long
    Dim Ai(1 To 58, 1 To 58) As Double
    Dim MST(1 To 10, 1 To 10, 1 To 2) As Long
    Dim XMW(1 To 10) As Double
    Dim result As Double
    Dim sub_units As String
    Dim tempT As Double
    Dim MGSG(1 To 116) As Long
    Dim RI(1 To 116) As Double
    Dim QI(1 To 116) As Double
    Dim MWS(1 To 116) As Double
    Dim MVS(1 To 116) As Double
    
    sub_units = "kmol/m3 chem"
    On Error Resume Next
    
    'Convert T to C
    'tempT = T - 273.15
    
    'MX = Cur_Info.MaxGroups
    
    'If MX <= 0 Then Exit Function
   
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
        If num_cur_chem_groups(i - 1) > 0 Then
            MST(2, i, 1) = cur_chem_groups(i - 1)
            MST(2, i, 2) = num_cur_chem_groups(i - 1)
        Else
            Exit For
        End If
    Next i
    MX = i
    
    Sol = 0     'Returned Value
    SOLSS = 0   'Not Important
    SOLLS = 0   'Not Important
    SOLErr = 0  'Not Important
    SOLTemp = 0 'Not Important

    For i = 0 To 3
        If load_BIP_data(BIPCode(i), Ai()) = True Then
            Exit For
        End If
    Next i
    
    If load_UNIFAC_data(MGSG(), RI(), QI(), MWS(), MVS()) = False Then
        MsgBox ("UNIFAC data not available")
        Exit Sub
    End If
    
    '9999 --- This doesn't work ---
    'Call AQSCALL2(Sol, SOLSS, SOLLS, SOLErr, SOLTemp, tempT, MX, MST(1, 1, 1), XMW(1), MGSG(1), Ai(1, 1), RI(1), QI(1), MWS(1), MVS(1))
   
    If SOLErr = -1 Then
        result = 0
    Else
        result = Sol / (M * 1000)
    End If
    
    'Convert T to K
    'T = T + 273.15
    'CalcSwaterUNIFAC = Convert(Swater, sub_units, ConvertToDefault(Swater), result)
    Value = Swater
End Sub


