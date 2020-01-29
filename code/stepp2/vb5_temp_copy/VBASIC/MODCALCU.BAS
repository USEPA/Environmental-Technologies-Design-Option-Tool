Attribute VB_Name = "modcalculate"
Option Explicit






Sub PredCalc(Code As Integer, CNT As Integer)
    
    'Home of the largest if block (check out case 35!)
    ' this function gets the units in the pearls default units
    ' REVISIONS  DMW    6/5/97  Adding unit parameters to method calls -
    '                   6/15/97  fixed case 16 and 17
    Dim i As Integer
    
    Select Case Code
        Case 0      'Molecular Weight
            InfoMethod(Code).value(CNT) = CalcMWUNIFAC(Get_DefaultUnit(MW))
            InfoMethod(Code).Unit = Get_DefaultUnit(MW)
            If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                InfoMethod(Code).Enabled(CNT) = True
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 1      'Liquid Density @ 25C
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcLDBhiruds(InfoMethod(MW).value(InfoMethod(MW).CurMethod), InfoMethod(MW).Unit, InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, 298.15, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Bhiruds (1978)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 1) = CalcLDModRackett(InfoMethod(MW).value(InfoMethod(MW).CurMethod), InfoMethod(MW).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, 298.15, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT + 1) = "Modified Rackett (1978)"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
        Case 3          'Melting Point
            If InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcMPTaftStarek(InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Taft and Starek (1930)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 1) = CalcMPLorenzHerz(InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT + 1) = "Lorenz and Herz (1922)"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
        Case 6              ' Vapor Pressure as f(t)
            ' this isn't a good way to do this but...need to indicate a null value until the user chooses to calculate
            'infomethod(Code).value(CNT) = 0
            'infomethod(Code).Unit = ""
            InfoMethod(Code).MethodName(CNT) = "Antoine"
            InfoMethod(Code).Enabled(CNT) = True
            
        Case 10             'Heat of Vaporization @ 25C
            If InfoMethod(HvapNBP).Enabled(InfoMethod(HvapNBP).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcHvapWatson(InfoMethod(HvapNBP).value(InfoMethod(HvapNBP).CurMethod), InfoMethod(HvapNBP).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Watson"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 11         'Heat of Vaporization @ NBP
            If InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True And InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcHvapNBPKlein(InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Klein (1949)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 1) = CalcHvapNBPChenPitzer(InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT + 1) = "Chen and Pitzer (1965)"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
        Case 15         'Diffusivity in Water
            InfoMethod(Code).value(CNT) = CalcDwaterHaydukMinhas(Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Dwater))
            InfoMethod(Code).Unit = Get_DefaultUnit(Dwater)
            If InfoMethod(Code).value(CNT) <> 0 Then
                InfoMethod(Code).MethodName(CNT) = "Hayduk and Minhas (1982)"
                InfoMethod(Code).Enabled(CNT) = True
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            InfoMethod(Code).value(CNT + 1) = CalcDwaterHaydukLaudie(Get_DefaultUnit(Dwater))
            InfoMethod(Code).Unit = Get_DefaultUnit(Dwater)
            If InfoMethod(Code).value(CNT + 1) <> 0 Then
                InfoMethod(Code).MethodName(CNT + 1) = "Hayduk and Laudie (1974)"
                InfoMethod(Code).Enabled(CNT + 1) = True
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 2) = CalcDwaterPolson(InfoMethod(MW).value(InfoMethod(MW).CurMethod), InfoMethod(MW).Unit, Get_DefaultUnit(Dwater))
                InfoMethod(Code).Unit = Get_DefaultUnit(Dwater)
                If InfoMethod(Code).value(CNT + 2) <> 0 Then
                    InfoMethod(Code).MethodName(CNT + 2) = "Polson"
                    InfoMethod(Code).Enabled(CNT + 2) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 2 And InfoMethod(Code).Enabled(CNT + 2) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 2) = True Then
                InfoMethod(Code).CurMethod = CNT + 2
            End If
            InfoMethod(Code).value(CNT + 3) = CalcDwaterWilkeChang(Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Dwater))
            InfoMethod(Code).Unit = Get_DefaultUnit(Dwater)
            If InfoMethod(Code).value(CNT + 3) <> 0 Then
                InfoMethod(Code).MethodName(CNT + 3) = "Wilke and Chang"
                InfoMethod(Code).Enabled(CNT + 3) = True
            End If
            If InfoMethod(Code).CurMethod = CNT + 3 And InfoMethod(Code).Enabled(CNT + 3) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 3) = True Then
                InfoMethod(Code).CurMethod = CNT + 3
            End If
        Case 16         'Diffusivity in Air
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcDairWilkeLeeMod(Cur_Info.OpT, Cur_Info.OpTUnit, Cur_Info.OpP, Cur_Info.OpPUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), InfoMethod(MW).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, Get_DefaultUnit(Dair))
                InfoMethod(Code).Unit = Get_DefaultUnit(Dair)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Wilke and Lee Modification"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 17         'Surface Tension @ 25C
            If InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcSTBrockBird(InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(ST25))
                InfoMethod(Code).Unit = Get_DefaultUnit(ST25)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Brock and Bird (1983)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 28         'Carbonaceous ThOD
            If Cur_Info.Formula <> "" Then
                InfoMethod(Code).value(CNT) = CalcThODcarbBaker(Cur_Info.Formula, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Baker (1994)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 29         'Combined ThOD
            If Cur_Info.Formula <> "" Then
                InfoMethod(Code).value(CNT) = CalcThODcombBaker(Cur_Info.Formula, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Baker (1994)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 30         'Chemical Oxygen Demand
            InfoMethod(Code).value(CNT) = InfoMethod(ThODcomb).value(InfoMethod(ThODcomb).CurMethod)
            InfoMethod(Code).Unit = Get_DefaultUnit(Code)
            If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                InfoMethod(Code).MethodName(CNT) = "MTU DIPPR"
                InfoMethod(Code).Enabled(CNT) = True
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 32         'Infinite Dilution Activity Coefficient of Water in Chemical
            InfoMethod(Code).value(CNT) = CalcACwaterUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPIndex(1), Get_DefaultUnit(Code))
            InfoMethod(Code).Unit = Get_DefaultUnit(Code)
            If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                InfoMethod(Code).Enabled(CNT) = True
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 33         'Henry's Constant
            If InfoMethod(ACchem).Enabled(InfoMethod(ACchem).CurMethod) = True And InfoMethod(VP).Enabled(InfoMethod(VP).CurMethod) = True Then
                InfoMethod(Code).value(12) = CalcHCUNIFAC(InfoMethod(ACchem).value(InfoMethod(ACchem).CurMethod), InfoMethod(VP).value(InfoMethod(VP).CurMethod), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(12) <> ERROR_FLAG And InfoMethod(Code).value(12) <> 0# Then
                    Select Case BIPIndex(1)
                        Case 1
                            InfoMethod(Code).MethodName(12) = "PEARLS"  ' PEARLS (AGLB)
                        Case 2
                            InfoMethod(Code).MethodName(12) = "PEARLS"  ' PEARLS (AVLE)
                        Case 3
                            InfoMethod(Code).MethodName(12) = "PEARLS"  ' PEARLS (AENV)
                        Case 4
                            InfoMethod(Code).MethodName(12) = "PEARLS"  ' PEARLS (ALLE)
                    End Select
                    InfoMethod(Code).Enabled(12) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = 12 And InfoMethod(Code).Enabled(12) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(12) = True Then
                InfoMethod(Code).CurMethod = 12
            End If
        Case 34         'Infinite Dilution Activity Coefficient of Chemical in Water
            If InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcACHansch(InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 Then
                    InfoMethod(Code).MethodName(CNT) = "Hansch (1968)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If BIPIndex(1) = 0 Then
                For i = 1 To 4
                    InfoMethod(Code).value(CNT + 1) = CalcACchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPHierarchy(1, i), Get_DefaultUnit(Code))
                    If InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 1) <> 0# Then
                        BIPIndex(1) = BIPHierarchy(1, i)
                        InfoMethod(Code).MethodName(CNT + 1) = "UNIFAC"
                        InfoMethod(Code).Enabled(CNT + 1) = True
                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                            InfoMethod(Code).CurMethod = CNT + 1
                        End If
                        Exit Sub
                    End If
                Next i
            Else
                InfoMethod(Code).value(CNT + 1) = CalcACchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPIndex(1), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 1) <> 0# Then
                    InfoMethod(Code).MethodName(CNT + 1) = "UNIFAC"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                    If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                        InfoMethod(Code).CurMethod = CNT + 1
                    End If
                    Exit Sub
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
        Case 35         'log Kow
            If BIPIndex(2) = 0 Then
                For i = 1 To 4
                    If InfoMethod(Code).Enabled(CNT) = False Then
                        InfoMethod(Code).value(CNT) = CalclogKowUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPHierarchy(2, i), Get_DefaultUnit(Code))
                        InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                        If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                            BIPIndex(2) = BIPHierarchy(2, i)
                            InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                            InfoMethod(Code).Enabled(CNT) = True
                            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                                InfoMethod(Code).CurMethod = CNT
                            End If
                            If InfoMethod(Schem).Enabled(5) = False Then
                                If InfoMethod(mp).Enabled(InfoMethod(mp).CurMethod) = True And InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                                    InfoMethod(Code).value(5) = CalcSchemYalkowsky(InfoMethod(mp).value(InfoMethod(mp).CurMethod), InfoMethod(mp).Unit, InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), InfoMethod(logKow).Unit, Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Code))
                                    InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                                    If InfoMethod(Code).value(5) <> ERROR_FLAG And InfoMethod(Code).value(5) <> 0# Then
                                        InfoMethod(Code).MethodName(5) = "Yalkowsky (1990)"
                                        InfoMethod(Code).Enabled(5) = True
                                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(5) = True Then
                                            InfoMethod(Code).CurMethod = 5
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            Else
                InfoMethod(Code).value(CNT) = CalclogKowUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPIndex(2), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                    InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                    InfoMethod(Code).Enabled(CNT) = True
                    If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                        InfoMethod(Code).CurMethod = CNT
                    End If
                    If InfoMethod(Schem).Enabled(5) = False Then
                        If InfoMethod(mp).Enabled(InfoMethod(mp).CurMethod) = True And InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                            InfoMethod(Code).value(5) = CalcSchemYalkowsky(InfoMethod(mp).value(InfoMethod(mp).CurMethod), InfoMethod(mp).Unit, InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), InfoMethod(logKow).Unit, Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Code))
                            InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                            If InfoMethod(Code).value(5) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 1) <> 0# Then
                                InfoMethod(Code).MethodName(5) = "Yalkowsky (1990)"
                                InfoMethod(Code).Enabled(5) = True
                                If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(5) = True Then
                                    InfoMethod(Code).CurMethod = 5
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Schem).CurMethod <> 5 Then
                If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(Schem).Enabled(InfoMethod(Schem).CurMethod) = True Then
                    InfoMethod(Code).value(CNT + 1) = CalclogKowKenagaGoring(InfoMethod(Schem).value(InfoMethod(Schem).CurMethod), InfoMethod(Schem).Unit, Get_DefaultUnit(Code))
                    InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                    If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                        InfoMethod(Code).MethodName(CNT + 1) = "Kenaga and Goring (1978)"
                        InfoMethod(Code).Enabled(CNT + 1) = True
                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                            InfoMethod(Code).CurMethod = CNT + 1
                        End If
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Schem).CurMethod <> 5 Then
                If InfoMethod(Schem).Enabled(InfoMethod(Schem).CurMethod) = True Then
                    InfoMethod(Code).value(CNT + 2) = CalclogKowHansch(InfoMethod(Schem).value(InfoMethod(Schem).CurMethod), InfoMethod(Schem).Unit, Get_DefaultUnit(Code))
                    InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                    If InfoMethod(Code).value(CNT + 2) <> 0 And InfoMethod(Code).value(CNT + 2) <> ERROR_FLAG Then
                        InfoMethod(Code).MethodName(CNT + 2) = "Hansch KOW (1968)"
                        InfoMethod(Code).Enabled(CNT + 2) = True
                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 2) = True Then
                            InfoMethod(Code).CurMethod = CNT + 2
                        End If
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 2 And InfoMethod(Code).Enabled(CNT + 2) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
        Case 36         'log Koc
            If InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalclogKocBaker(InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Baker (1994)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 37         'Bioconcentration Factor
            If InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcBCFKobayshi(InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Kobayshi (1981)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(Schem).Enabled(InfoMethod(Schem).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 1) = CalcBCFKenagaGoring(InfoMethod(Schem).value(InfoMethod(Schem).CurMethod), InfoMethod(Schem).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT + 1) = "Kenaga and Goring (1980)"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
        Case 39         'Solubility Limit of Chemical in Water
            If InfoMethod(mp).Enabled(InfoMethod(mp).CurMethod) = True And InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 2) = CalcSchemYalkowsky(InfoMethod(mp).value(InfoMethod(mp).CurMethod), InfoMethod(mp).Unit, InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), InfoMethod(logKow).Unit, Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 2) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 2) <> 0# Then
                    InfoMethod(Code).MethodName(CNT + 2) = "Yalkowsky (1990)"
                    InfoMethod(Code).Enabled(CNT + 2) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 2 And InfoMethod(Code).Enabled(CNT + 2) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 2) = True Then
                InfoMethod(Code).CurMethod = CNT + 2
            End If
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True Then
                If BIPIndex(3) = 0 Then
                    For i = 1 To 4
                        InfoMethod(Code).value(CNT + 3) = CalcSchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPHierarchy(3, i), Get_DefaultUnit(Code))
                        If InfoMethod(Code).value(CNT + 3) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 3) <> 0# Then
                            BIPIndex(3) = BIPHierarchy(3, i)
                            InfoMethod(Code).MethodName(CNT + 3) = "UNIFAC"
                            InfoMethod(Code).Enabled(CNT + 3) = True
                            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 3) = True Then
                                InfoMethod(Code).CurMethod = CNT + 3
                            End If
                            Exit Sub
                        End If
                    Next i
                Else
                    InfoMethod(Code).value(CNT + 3) = CalcSchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPIndex(3), Get_DefaultUnit(Code))
                    InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                    If InfoMethod(Code).value(CNT + 3) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 3) <> 0# Then
                        InfoMethod(Code).MethodName(CNT + 3) = "UNIFAC"
                        InfoMethod(Code).Enabled(CNT + 3) = True
                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 3) = True Then
                            InfoMethod(Code).CurMethod = CNT + 3
                        End If
                        Exit Sub
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 3 And InfoMethod(Code).Enabled(CNT + 3) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
        Case 40         'Solubility Limit of Water in Chemical
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcSwaterUNIFAC(Cur_Info.OpT, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPIndex(3), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                    InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
    End Select
       
End Sub



    


Sub PredCalcMethod(Code As Integer, CNT As Integer)
        
    'Home of the copy of the largest if block (check out case 35!)
    ' this function gets the units in the pearls default units
    ' REVISIONS  DMW    6/5/97  Adding unit parameters to method calls -
    '                   6/15/97  fixed case 16 and 17
    Dim i As Integer
    
    Select Case Code
        Case 0      ' Molecular Weight
            InfoMethod(Code).value(CNT) = CalcMWUNIFAC(Get_DefaultUnit(MW))
            InfoMethod(Code).Unit = Get_DefaultUnit(MW)
            If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                'infomethod(Code).value(CNT) = Convert(CLng(Code), Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                InfoMethod(Code).Enabled(CNT) = True
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 1      ' Liquid Density
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcLDBhiruds(InfoMethod(MW).value(InfoMethod(MW).CurMethod), InfoMethod(MW).Unit, InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, 298.15, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    'infomethod(Code).VALUE(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).VALUE(CNT))
                    InfoMethod(Code).MethodName(CNT) = "Bhiruds (1978)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 1) = CalcLDModRackett(InfoMethod(MW).value(InfoMethod(MW).CurMethod), InfoMethod(MW).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, 298.15, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                    'infomethod(Code).VALUE(CNT + 1) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).VALUE(CNT + 1))
                    InfoMethod(Code).MethodName(CNT + 1) = "Modified Rackett (1978)"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
        Case 3      'Melting Point
            If InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcMPTaftStarek(InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0# And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                    InfoMethod(Code).MethodName(CNT) = "Taft and Starek (1930)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 1) = CalcMPLorenzHerz(InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                    'infomethod(Code).value(CNT + 1) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT + 1))
                    InfoMethod(Code).MethodName(CNT + 1) = "Lorenz and Herz (1922)"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
        Case 6              ' Vapor Pressure as f(t)
            ' this isn't a good way to do this but...need to indicate a null value until the user chooses to calculate
            'infomethod(Code).value(CNT) = 0
            'infomethod(Code).Unit = ""
            InfoMethod(Code).MethodName(CNT) = "Antoine"
            InfoMethod(Code).Enabled(CNT) = True
        Case 10
            If InfoMethod(HvapNBP).Enabled(InfoMethod(HvapNBP).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcHvapWatson(InfoMethod(HvapNBP).value(InfoMethod(HvapNBP).CurMethod), InfoMethod(HvapNBP).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                    InfoMethod(Code).MethodName(CNT) = "Watson"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 11
            If InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True And InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcHvapNBPKlein(InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                    InfoMethod(Code).MethodName(CNT) = "Klein (1949)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 1) = CalcHvapNBPChenPitzer(InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> 0# And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                    'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT + 1))
                    InfoMethod(Code).MethodName(CNT + 1) = "Chen and Pitzer (1965)"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
        Case 15
            'infomethod(Code).value(CNT) = CalcDwaterHaydukMinhas(Cur_Info.OpT, Cur_Info.OpT.Unit, Get_DefaultUnit(Dwater))
            'FIX : this did not like "Cur_Info.OpT.Unit"
            InfoMethod(Code).value(CNT) = CalcDwaterHaydukMinhas(Cur_Info.OpT, "K", Get_DefaultUnit(Dwater))
            InfoMethod(Code).Unit = Get_DefaultUnit(Dwater)
            If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                InfoMethod(Code).MethodName(CNT) = "Hayduk and Minhas (1982)"
                InfoMethod(Code).Enabled(CNT) = True
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            InfoMethod(Code).value(CNT + 1) = CalcDwaterHaydukLaudie(Get_DefaultUnit(Dwater))
            InfoMethod(Code).Unit = Get_DefaultUnit(Dwater)
            If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                'infomethod(Code).value(CNT + 1) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT + 1))
                InfoMethod(Code).MethodName(CNT + 1) = "Hayduk and Laudie (1974)"
                InfoMethod(Code).Enabled(CNT + 1) = True
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 2) = CalcDwaterPolson(InfoMethod(MW).value(InfoMethod(MW).CurMethod), InfoMethod(MW).Unit, Get_DefaultUnit(Dwater))
                InfoMethod(Code).Unit = Get_DefaultUnit(Dwater)
                If InfoMethod(Code).value(CNT + 2) <> 0 And InfoMethod(Code).value(CNT + 2) <> ERROR_FLAG Then
                    'infomethod(Code).value(CNT + 2) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT + 2))
                    InfoMethod(Code).MethodName(CNT + 2) = "Polson"
                    InfoMethod(Code).Enabled(CNT + 2) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 2 And InfoMethod(Code).Enabled(CNT + 2) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 2) = True Then
                InfoMethod(Code).CurMethod = CNT + 2
            End If
            InfoMethod(Code).value(CNT + 3) = CalcDwaterWilkeChang(Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Dwater))
            InfoMethod(Code).Unit = Get_DefaultUnit(Dwater)
            If InfoMethod(Code).value(CNT + 3) <> 0 And InfoMethod(Code).value(CNT + 3) <> ERROR_FLAG Then
                'infomethod(Code).value(CNT + 3) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT + 3))
                InfoMethod(Code).MethodName(CNT + 3) = "Wilke and Chang"
                InfoMethod(Code).Enabled(CNT + 3) = True
            End If
            If InfoMethod(Code).CurMethod = CNT + 3 And InfoMethod(Code).Enabled(CNT + 3) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 3) = True Then
                InfoMethod(Code).CurMethod = CNT + 3
            End If
        Case 16
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcDairWilkeLeeMod(Cur_Info.OpT, Cur_Info.OpTUnit, Cur_Info.OpP, Cur_Info.OpPUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), InfoMethod(MW).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, Get_DefaultUnit(Dair))
                InfoMethod(Code).Unit = Get_DefaultUnit(Dair)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                    InfoMethod(Code).MethodName(CNT) = "Wilke and Lee Modification"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 17
            If InfoMethod(CP).Enabled(InfoMethod(CP).CurMethod) = True And InfoMethod(CT).Enabled(InfoMethod(CT).CurMethod) = True And InfoMethod(NBP).Enabled(InfoMethod(NBP).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcSTBrockBird(InfoMethod(CP).value(InfoMethod(CP).CurMethod), InfoMethod(CP).Unit, InfoMethod(CT).value(InfoMethod(CT).CurMethod), InfoMethod(CT).Unit, InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), InfoMethod(NBP).Unit, Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(ST25))
                InfoMethod(Code).Unit = Get_DefaultUnit(ST25)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    ' really need this?  DENISE check
                    'infomethod(Code).VALUE(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).VALUE(CNT))
                    InfoMethod(Code).MethodName(CNT) = "Brock and Bird (1983)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 28
            If Cur_Info.Formula <> "" Then
                InfoMethod(Code).value(CNT) = CalcThODcarbBaker(Cur_Info.Formula, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 Then
                    'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                    InfoMethod(Code).MethodName(CNT) = "Baker (1994)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 29
            If Cur_Info.Formula <> "" Then
                InfoMethod(Code).value(CNT) = CalcThODcombBaker(Cur_Info.Formula, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0# And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                    InfoMethod(Code).MethodName(CNT) = "Baker (1994)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 30
            InfoMethod(Code).value(CNT) = InfoMethod(ThODcomb).value(InfoMethod(ThODcomb).CurMethod)
            InfoMethod(Code).Unit = Get_DefaultUnit(Code)
            If InfoMethod(Code).value(CNT) <> 0# And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                'infomethod(Code).value(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).value(CNT))
                InfoMethod(Code).MethodName(CNT) = "MTU DIPPR"
                InfoMethod(Code).Enabled(CNT) = True
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 32
            InfoMethod(Code).value(CNT) = CalcACwaterUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPIndex(1), Get_DefaultUnit(Code))
            InfoMethod(Code).Unit = Get_DefaultUnit(Code)
            If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                InfoMethod(Code).Enabled(CNT) = True
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 33
            If InfoMethod(ACchem).Enabled(InfoMethod(ACchem).CurMethod) = True And InfoMethod(VP).Enabled(InfoMethod(VP).CurMethod) = True Then
                InfoMethod(Code).value(12) = CalcHCUNIFAC(InfoMethod(ACchem).value(InfoMethod(ACchem).CurMethod), InfoMethod(VP).value(InfoMethod(VP).CurMethod), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(12) <> ERROR_FLAG And InfoMethod(Code).value(12) <> 0# Then
                    'infomethod(Code).VALUE(12) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).VALUE(12))
                    Select Case BIPIndex(1)
                        Case 1
                            InfoMethod(Code).MethodName(12) = "PEARLS"  ' PEARLS (AGLB)
                        Case 2
                            InfoMethod(Code).MethodName(12) = "PEARLS"  ' PEARLS (AVLE)
                        Case 3
                            InfoMethod(Code).MethodName(12) = "PEARLS"  ' PEARLS (AENV)
                        Case 4
                            InfoMethod(Code).MethodName(12) = "PEARLS"  ' PEARLS (ALLE)
                    End Select
                    InfoMethod(Code).Enabled(12) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = 12 And InfoMethod(Code).Enabled(12) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(12) = True Then
                InfoMethod(Code).CurMethod = 12
            End If
        Case 34
            If InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcACHansch(InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Hansch (1968)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If BIPIndex(1) = 0 Then
                For i = 1 To 4
                    InfoMethod(Code).value(CNT + 1) = CalcACchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPHierarchy(1, i), Get_DefaultUnit(Code))
                    If InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 1) <> 0# Then
                        BIPIndex(1) = BIPHierarchy(1, i)
                        InfoMethod(Code).MethodName(CNT + 1) = "UNIFAC"
                        InfoMethod(Code).Enabled(CNT + 1) = True
                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                            InfoMethod(Code).CurMethod = CNT + 1
                        End If
                        Exit Sub
                    End If
                Next i
            Else
                InfoMethod(Code).value(CNT + 1) = CalcACchemUNIFAC(Cur_Info.OpT, "K", BIPIndex(1), Get_DefaultUnit(Swater))
                If InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 1) <> 0# Then
                    InfoMethod(Code).MethodName(CNT + 1) = "UNIFAC"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                    If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                        InfoMethod(Code).CurMethod = CNT + 1
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
        Case 35
            If BIPIndex(2) = 0 Then
                For i = 1 To 4
                    If InfoMethod(Code).Enabled(CNT) = False Then
                        InfoMethod(Code).value(CNT) = CalclogKowUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPHierarchy(2, i), Get_DefaultUnit(Code))
                        InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                        If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                            BIPIndex(2) = BIPHierarchy(2, i)
                            InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                            InfoMethod(Code).Enabled(CNT) = True
                            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                                InfoMethod(Code).CurMethod = CNT
                            End If
                            If InfoMethod(Schem).Enabled(5) = False Then
                                If InfoMethod(mp).Enabled(InfoMethod(mp).CurMethod) = True And InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                                    InfoMethod(Code).value(5) = CalcSchemYalkowsky(InfoMethod(mp).value(InfoMethod(mp).CurMethod), InfoMethod(mp).Unit, InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), InfoMethod(logKow).Unit, Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Code))
                                    InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                                    If InfoMethod(Code).value(5) <> ERROR_FLAG And InfoMethod(Code).value(5) <> 0# Then
                                        InfoMethod(Code).MethodName(5) = "Yalkowsky (1990)"
                                        InfoMethod(Code).Enabled(5) = True
                                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(5) = True Then
                                            InfoMethod(Code).CurMethod = 5
                                        End If
                                    End If
                                End If
                            End If
                            Exit Sub
                        End If
                    End If
                Next i
            Else
                InfoMethod(Code).value(CNT) = CalclogKowUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, BIPIndex(2), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                    InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                    InfoMethod(Code).Enabled(CNT) = True
                    If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                        InfoMethod(Code).CurMethod = CNT
                    End If
                    If InfoMethod(Schem).Enabled(5) = False Then
                        If InfoMethod(mp).Enabled(InfoMethod(mp).CurMethod) = True And InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                            InfoMethod(Code).value(5) = CalcSchemYalkowsky(InfoMethod(mp).value(InfoMethod(mp).CurMethod), InfoMethod(mp).Unit, InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), InfoMethod(logKow).Unit, Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Code))
                            InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                            If InfoMethod(Code).value(5) <> ERROR_FLAG And InfoMethod(Code).value(5) <> 0# Then
                                InfoMethod(Code).MethodName(5) = "Yalkowsky (1990)"
                                InfoMethod(Code).Enabled(5) = True
                                If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(5) = True Then
                                    InfoMethod(Code).CurMethod = 5
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Schem).CurMethod <> 5 Then
                If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(Schem).Enabled(InfoMethod(Schem).CurMethod) = True Then
                    InfoMethod(Code).value(CNT + 1) = CalclogKowKenagaGoring(InfoMethod(Schem).value(InfoMethod(Schem).CurMethod), InfoMethod(Schem).Unit, Get_DefaultUnit(Code))
                    InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                    If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                        InfoMethod(Code).MethodName(CNT + 1) = "Kenaga and Goring (1978)"
                        InfoMethod(Code).Enabled(CNT + 1) = True
                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                            InfoMethod(Code).CurMethod = CNT + 1
                        End If
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Schem).CurMethod <> 5 Then
                If InfoMethod(Schem).Enabled(InfoMethod(Schem).CurMethod) = True Then
                    InfoMethod(Code).value(CNT + 2) = CalclogKowHansch(InfoMethod(Schem).value(InfoMethod(Schem).CurMethod), InfoMethod(Schem).Unit, Get_DefaultUnit(Code))
                    InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                    If InfoMethod(Code).value(CNT + 2) <> 0 And InfoMethod(Code).value(CNT + 2) <> ERROR_FLAG Then
                        InfoMethod(Code).MethodName(CNT + 2) = "Hansch KOW (1968)"
                        InfoMethod(Code).Enabled(CNT + 2) = True
                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 2) = True Then
                            InfoMethod(Code).CurMethod = CNT + 2
                        End If
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 2 And InfoMethod(Code).Enabled(CNT + 2) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
        Case 36
            If InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalclogKocBaker(InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0# And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Baker (1994)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
        Case 37
            If InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcBCFKobayshi(InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> 0 And InfoMethod(Code).value(CNT) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT) = "Kobayshi (1981)"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True And InfoMethod(Schem).Enabled(InfoMethod(Schem).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 1) = CalcBCFKenagaGoring(InfoMethod(Schem).value(InfoMethod(Schem).CurMethod), InfoMethod(Schem).Unit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 1) <> 0 And InfoMethod(Code).value(CNT + 1) <> ERROR_FLAG Then
                    InfoMethod(Code).MethodName(CNT + 1) = "Kenaga and Goring (1980)"
                    InfoMethod(Code).Enabled(CNT + 1) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 1 And InfoMethod(Code).Enabled(CNT + 1) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 1) = True Then
                InfoMethod(Code).CurMethod = CNT + 1
            End If
        Case 39
            If InfoMethod(mp).Enabled(InfoMethod(mp).CurMethod) = True And InfoMethod(logKow).Enabled(InfoMethod(logKow).CurMethod) = True Then
                InfoMethod(Code).value(CNT + 2) = CalcSchemYalkowsky(InfoMethod(mp).value(InfoMethod(mp).CurMethod), InfoMethod(mp).Unit, InfoMethod(logKow).value(InfoMethod(logKow).CurMethod), InfoMethod(logKow).Unit, Cur_Info.OpT, Cur_Info.OpTUnit, Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT + 2) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 2) <> 0# Then
                    'infomethod(Code).VALUE(CNT + 2) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).VALUE(CNT + 2))
                    InfoMethod(Code).MethodName(CNT + 2) = "Yalkowsky (1990)"
                    InfoMethod(Code).Enabled(CNT + 2) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 2 And InfoMethod(Code).Enabled(CNT + 2) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 2) = True Then
                InfoMethod(Code).CurMethod = CNT + 2
            End If
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True Then
                If BIPIndex(3) = 0 Then
                    For i = 1 To 4
                        InfoMethod(Code).value(CNT + 3) = CalcSchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPHierarchy(3, i), Get_DefaultUnit(Code))
                        InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                        If InfoMethod(Code).value(CNT + 3) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 3) <> 0# Then
                            'infomethod(Code).VALUE(CNT + 3) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).VALUE(CNT + 3))
                            BIPIndex(3) = BIPHierarchy(3, i)
                            InfoMethod(Code).MethodName(CNT + 3) = "UNIFAC"
                            InfoMethod(Code).Enabled(CNT + 3) = True
                            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 3) = True Then
                                InfoMethod(Code).CurMethod = CNT + 3
                            End If
                            Exit Sub
                        End If
                    Next i
                Else
                    InfoMethod(Code).value(CNT + 3) = CalcSchemUNIFAC(Cur_Info.OpT, Cur_Info.OpTUnit, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPIndex(3), Get_DefaultUnit(Code))
                    InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                    If InfoMethod(Code).value(CNT + 3) <> ERROR_FLAG And InfoMethod(Code).value(CNT + 3) <> 0# Then
                        'infomethod(Code).VALUE(CNT + 3) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).VALUE(CNT + 3))
                        InfoMethod(Code).MethodName(CNT + 3) = "UNIFAC"
                        InfoMethod(Code).Enabled(CNT + 3) = True
                        If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT + 3) = True Then
                            InfoMethod(Code).CurMethod = CNT + 3
                        End If
                        Exit Sub
                    End If
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT + 3 And InfoMethod(Code).Enabled(CNT + 3) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
        Case 40
            If InfoMethod(MW).Enabled(InfoMethod(MW).CurMethod) = True Then
                InfoMethod(Code).value(CNT) = CalcSwaterUNIFAC(Cur_Info.OpT, InfoMethod(MW).value(InfoMethod(MW).CurMethod), BIPIndex(3), Get_DefaultUnit(Code))
                InfoMethod(Code).Unit = Get_DefaultUnit(Code)
                If InfoMethod(Code).value(CNT) <> ERROR_FLAG And InfoMethod(Code).value(CNT) <> 0# Then
                    'infomethod(Code).VALUE(CNT) = Convert(Get_DefaultUnit(Code), Get_DefaultUnit(Code), infomethod(Code).VALUE(CNT))
                    InfoMethod(Code).MethodName(CNT) = "UNIFAC"
                    InfoMethod(Code).Enabled(CNT) = True
                End If
            End If
            If InfoMethod(Code).CurMethod = CNT And InfoMethod(Code).Enabled(CNT) = False Then
                InfoMethod(Code).CurMethod = 0
            End If
            If InfoMethod(Code).CurMethod = 0 And InfoMethod(Code).Enabled(CNT) = True Then
                InfoMethod(Code).CurMethod = CNT
            End If
    End Select
                
End Sub


Sub RecalcPred()

    ' REVISIONS  DMW    6/5/97  when the OpT and OpP values get converted to defaults - set the units as well
    '                           this conversion is in most (if not all) cases redundant since the methods do it also
    Dim i As Integer
    Dim J As Integer
    Dim TempOpTVal As Double
    Dim TempOpPVal As Double
    Dim TempOpPUnits As String
    Dim TempOpTUnits As String
                                                         
    'Recalculate only the predicted properties, this
    'is used only for the method screen 'Accept' button
    
    'Set operating conditions to K(OpT) and Pa(OpP) for calculations
    TempOpTVal = Cur_Info.OpT
    TempOpTUnits = Cur_Info.OpTUnit
    TempOpPVal = Cur_Info.OpP
    TempOpPUnits = Cur_Info.OpPUnit
    If Not IsDefault(Trim(Cur_Info.OpTUnit), OptTemp) Then
        Cur_Info.OpT = Convert(Cur_Info.OpT, OptTemp, Trim(Cur_Info.OpTUnit), "K", False)
        'Cur_Info.OpTUnit = ConvertToDefault(OpT) (gets error because OpT is not a valid code)
        'FIX OpT should be in default unit
    End If
    If Not IsDefault(Trim(Cur_Info.OpPUnit), OptPress) Then
        Cur_Info.OpP = Convert(Cur_Info.OpP, OptPress, Trim(Cur_Info.OpPUnit), "Pa", False) 'FIX OpP has code
        'Cur_Info.OpPUnit = ConvertToDefault(OpP)
        'FIX OpP should also be in default unit (gets error because OpP is not a valid code)
    End If
                        
    'Load data from predictive models
    For i = 0 To NumProperties
        Call PredCalcMethod(CalcHierarchy(i), 4)
        
    Next i
    
    'Convert TFT minimum and maximum T's to proper units
    Call ConvertTFTUnits2("K", DefaultTFTUnit)
  
    'Reset operating conditions back to previous state
    Cur_Info.OpT = TempOpTVal
    Cur_Info.OpTUnit = TempOpTUnits
    Cur_Info.OpP = TempOpPVal
    Cur_Info.OpPUnit = TempOpPUnits
    
    'update cur_disp
    Call update_DisplayData
    
End Sub

Sub Recalculate()
    
    Dim i As Integer
    Dim J As Integer
    Dim M As Integer
    Dim ReLoad As Boolean
    Dim DBTbl As Recordset
    Dim TempOpTVal As Double
    Dim TempOpPVal As Double
    Dim TempOpTUnit As String
    Dim TempOpPUnit As String
    Dim TempCAS As Long
    On Error Resume Next
                                                           
    'Store last CAS number viewed
    Set DBTbl = DBJetMaster.OpenRecordset("Last CAS", dbOpenTable)
    DBTbl.Edit
    DBTbl("LastCAS") = Cur_Info.CAS
    DBTbl.Update
    DBTbl.Close
        
    ' make sure Cur_Info is also updated on the FRMMain (this is in case calculate is called from printing or graphing functions)
    
     
    frmmain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS
    frmmain.Refresh
    frmmain!TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    Cur_Info.name = frmmain!Data1.Recordset("Name")
    Cur_Info.Formula = frmmain!Data1.Recordset("Formula")
    Cur_Info.source = frmmain!Data1.Recordset("Source")
    Cur_Info.Family = frmmain!Data1.Recordset("Chemical Family")
    Cur_Info.SMILES = frmmain!Data1.Recordset("Smiles")
    
    'Reinitialize property enable flags
    For i = 0 To NumProperties
        InfoMethod(i).CurMethod = 0    ' the null value
        For J = 1 To NumMethods
            InfoMethod(i).EqNum(J) = 0
            InfoMethod(i).value(J) = 0
            InfoMethod(i).TFTUnit = DefaultTFTUnit
            InfoMethod(i).Enabled(J) = False
        Next J
    Next i
                                                         
    'Load user data if available (if it's in the save table)
    ReLoad = False
    If GetUserData = True Then
        ReLoad = True
    End If
                                                     
    'Set operating conditions to K(OpT) and Pa(OpP) for calculations
    TempOpTVal = Cur_Info.OpT
    TempOpTUnit = Cur_Info.OpTUnit
    TempOpPVal = Cur_Info.OpP
    TempOpPUnit = Cur_Info.OpPUnit
    If Not IsDefault(Trim(Cur_Info.OpTUnit), OptTemp) Then
        Cur_Info.OpT = Convert(Cur_Info.OpT, OptTemp, Trim(Cur_Info.OpTUnit), "K", False)
        Cur_Info.OpTUnit = "K"
    End If
    If Not IsDefault(Trim(Cur_Info.OpPUnit), OptPress) Then
        Cur_Info.OpP = Convert(Cur_Info.OpP, OptPress, Trim(Cur_Info.OpPUnit), "Pa", False) 'FIX OpP has no code
        Cur_Info.OpPUnit = "Pa"
    End If
                    
    For i = 0 To NumProperties
        InfoMethod(i).MethodName(1) = "801 Database"
        InfoMethod(i).TFT = Cur_Info.OpT
    Next i
    
    'Load 801 data
    Call Get801Data(1)
   
    'Load 911 data
    Call Get911Data(2)
               
        ' If CurMethod is 911 or 801 and they don't exist, set CurMethod to 0
   
    If ReLoad = True Then
        For i = 0 To NumProperties
            If InfoMethod(i).CurMethod = 2 Then
                If InfoMethod(i).Enabled(InfoMethod(i).CurMethod) = False Then
                    InfoMethod(i).CurMethod = 0
                End If
            ElseIf InfoMethod(i).CurMethod = 1 Then
                If InfoMethod(i).Enabled(InfoMethod(i).CurMethod) = False Then
                    InfoMethod(i).CurMethod = 0
                End If
            End If
        Next i
    End If
  
    'Load Master data
    Call GetMasterData(3)
    
    'do block 5 calculations if the database is there (this fills in holes for block 5 properties)
    On Error GoTo 0
    If PathBlock5 <> NULLPATH Then
        ' calculate block 5 properties, storing starting with method 4
        ' Block5Calc will set CurMethod if possible
        Call Block5Calc(4)
    End If
    On Error Resume Next
   
    ' In case CurMethod still hasn't been set (is still 0)
    ' now set CurMethod based on the chosen or available method in the order: 801, 911, Master
    For i = 0 To NumProperties
        If InfoMethod(i).Enabled(InfoMethod(i).CurMethod) = True Then
            GoTo next_I
        Else
            For J = 1 To NumMethods
                ' it's not a block 5 property, set to first available method (not based on a user preference list for now)
                If InfoMethod(i).Enabled(J) = True Then
                    InfoMethod(i).CurMethod = J
                    GoTo next_I
                End If
            Next J
        End If
next_I:
    Next i
    
    'Load data from predictive models (this will fill in holes that 911 and 801 and master didn't fill)
    For i = 0 To NumProperties
        If i < 23 Or i > 26 Then
            ' if it's not a block 5 property then call this function, starts with method 4
            ' predcalc will set CurMethod if it's not already set
            Call PredCalc(CalcHierarchy(i), 4)
        End If
    Next i
        
    'Convert values to user default values (from PEARLS default)
    For i = 0 To NumProperties
        CurProp = i
        InfoMethod(i).Unit = Get_DefaultUnit(i)
    Next i

    
   
    'Convert TFT minimum and maximum T's to proper units
    'Call ConvertTFTUnits2("K", DefaultTFTUnit)
  
    'Reset operating conditions back to previous state
    Cur_Info.OpT = TempOpTVal
    Cur_Info.OpTUnit = TempOpTUnit
    Cur_Info.OpP = TempOpPVal
    Cur_Info.OpPUnit = TempOpPUnit
            
    'the internal info changed so change display info to match new internal
    Call update_DisplayData
    
End Sub



Private Sub Block5Calc(CNT As Integer)

    ' this function calls the block 5 calculating function
    ' it sends -1 if it wants all 4 properties calculated at once
    ' it sends the index of the property if it only wants one property calculated
    ' for normal purposes have it do them all at once because
    ' every time it gets called it has to read from the
    ' block 5 database
    ' MODIFIED:
    '       5/30/97:  DMW - CNT is the index in the method array we're starting from (curtrently after master data)

Dim Code As Integer
Dim which_prop As Integer
Dim i As Integer

'which_prop = 1
Call do_block5_calcs(-1, CNT)   ' -1 indicates want to do all calculations at once
For Code = UFL To AIT
    which_prop = Code - 23
    If InfoMethod(Code).CurMethod = 0 Or InfoMethod(Code).Enabled(InfoMethod(Code).CurMethod) = False Then
            For i = 1 To NumMethods
                If InfoMethod(Code).Enabled(i) = True Then
                    InfoMethod(Code).CurMethod = B5Preference(which_prop, i - 1)
                    Exit For
                End If
            Next i
    End If
    'If infomethod(Code).CurMethod = 0 Or infomethod(Code).Enabled(infomethod(Code).CurMethod) = False Then
            'For I = 1 To NumMethods
             '   If infomethod(Code).Enabled(I) = True Then
              '      infomethod(Code).CurMethod = B5Preference(which_prop, I - 1)
              '      Exit For
              '  End If
            'Next I
    'End If
    
Next Code
End Sub
