Attribute VB_Name = "modDBAccess"
Option Explicit
Sub Get801Data(CNT As Integer)
    
    Dim i As Integer
    Dim DBTbl As Recordset
        
    'Check to see if the user owns the 801 database
    If DIPPR801 = False Or Path801 = NULLPATH Then Exit Sub

    
    On Error GoTo Get801DataError

    Set DBTbl = DBJetMaster.OpenRecordset("DIPPR801", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    If DBTbl.NoMatch Then
         DBTbl.Close
         Exit Sub
    End If
        
    InfoMethod(MW).value(CNT) = DBTbl("MW")
    InfoMethod(CT).value(CNT) = DBTbl("TC")
    InfoMethod(CP).value(CNT) = DBTbl("PC")
    InfoMethod(CV).value(CNT) = DBTbl("VC")
    InfoMethod(mp).value(CNT) = DBTbl("MP")
    InfoMethod(NBP).value(CNT) = DBTbl("NBP")
    InfoMethod(hfor).value(CNT) = DBTbl("HFOR")
    InfoMethod(Hcomb).value(CNT) = DBTbl("HCOM")
    InfoMethod(FP).value(CNT) = DBTbl("FP")
    InfoMethod(LFL).value(CNT) = DBTbl("FLML")
    InfoMethod(UFL).value(CNT) = DBTbl("FLMU")
    InfoMethod(AIT).value(CNT) = DBTbl("AIT")
    
    For i = 0 To NumProperties
        If InfoMethod(i).value(CNT) <> 0 Then
            InfoMethod(i).Enabled(CNT) = True
        End If
    Next i
    
    InfoMethod(LD).EqNum(CNT) = DBTbl("LDNEQN")
    InfoMethod(LD).Coeff(CNT, 1) = DBTbl("LDNA")
    InfoMethod(LD).Coeff(CNT, 2) = DBTbl("LDNB")
    InfoMethod(LD).Coeff(CNT, 3) = DBTbl("LDNC")
    InfoMethod(LD).Coeff(CNT, 4) = DBTbl("LDND")
    InfoMethod(LD).Coeff(CNT, 5) = DBTbl("LDNE")
    InfoMethod(LD).MinT(CNT) = DBTbl("LDNTMIN")
    InfoMethod(LD).MaxT(CNT) = DBTbl("LDNTMAX")
    InfoMethod(LD).value(CNT) = CalcFofT(LD, CNT)
    If InfoMethod(LD).value(CNT) <> 0 Then InfoMethod(LD).Enabled(CNT) = True
    
    InfoMethod(VP).EqNum(CNT) = DBTbl("VPEQN")
    InfoMethod(VP).Coeff(CNT, 1) = DBTbl("VPA")
    InfoMethod(VP).Coeff(CNT, 2) = DBTbl("VPB")
    InfoMethod(VP).Coeff(CNT, 3) = DBTbl("VPC")
    InfoMethod(VP).Coeff(CNT, 4) = DBTbl("VPD")
    InfoMethod(VP).Coeff(CNT, 5) = DBTbl("VPE")
    InfoMethod(VP).MinT(CNT) = DBTbl("VPTMIN")
    InfoMethod(VP).MaxT(CNT) = DBTbl("VPTMAX")
    InfoMethod(VP).value(CNT) = CalcFofT(VP, CNT)
    If InfoMethod(VP).value(CNT) <> 0 Then InfoMethod(VP).Enabled(CNT) = True
    
    InfoMethod(Hvap).EqNum(CNT) = DBTbl("HVPEQN")
    InfoMethod(Hvap).Coeff(CNT, 1) = DBTbl("HVPA")
    InfoMethod(Hvap).Coeff(CNT, 2) = DBTbl("HVPB")
    InfoMethod(Hvap).Coeff(CNT, 3) = DBTbl("HVPC")
    InfoMethod(Hvap).Coeff(CNT, 4) = DBTbl("HVPD")
    InfoMethod(Hvap).Coeff(CNT, 5) = DBTbl("HVPE")
    InfoMethod(Hvap).MinT(CNT) = DBTbl("HVPTMIN")
    InfoMethod(Hvap).MaxT(CNT) = DBTbl("HVPTMAX")
    InfoMethod(Hvap).value(CNT) = CalcFofT(Hvap, CNT)
    If InfoMethod(Hvap).value(CNT) <> 0 Then InfoMethod(Hvap).Enabled(CNT) = True
    
    InfoMethod(LHC).EqNum(CNT) = DBTbl("LCPEQN")
    InfoMethod(LHC).Coeff(CNT, 1) = DBTbl("LCPA")
    InfoMethod(LHC).Coeff(CNT, 2) = DBTbl("LCPB")
    InfoMethod(LHC).Coeff(CNT, 3) = DBTbl("LCPC")
    InfoMethod(LHC).Coeff(CNT, 4) = DBTbl("LCPD")
    InfoMethod(LHC).Coeff(CNT, 5) = DBTbl("LCPE")
    InfoMethod(LHC).MinT(CNT) = DBTbl("LCPTMIN")
    InfoMethod(LHC).MaxT(CNT) = DBTbl("LCPTMAX")
    InfoMethod(LHC).value(CNT) = CalcFofT(LHC, CNT)
    If InfoMethod(LHC).value(CNT) <> 0 Then InfoMethod(LHC).Enabled(CNT) = True
    
    InfoMethod(VHC).EqNum(CNT) = DBTbl("ICPEQN")
    InfoMethod(VHC).Coeff(CNT, 1) = DBTbl("ICPA")
    InfoMethod(VHC).Coeff(CNT, 2) = DBTbl("ICPB")
    InfoMethod(VHC).Coeff(CNT, 3) = DBTbl("ICPC")
    InfoMethod(VHC).Coeff(CNT, 4) = DBTbl("ICPD")
    InfoMethod(VHC).Coeff(CNT, 5) = DBTbl("ICPE")
    InfoMethod(VHC).MinT(CNT) = DBTbl("ICPTMIN")
    InfoMethod(VHC).MaxT(CNT) = DBTbl("ICPTMAX")
    InfoMethod(VHC).value(CNT) = CalcFofT(VHC, CNT)
    If InfoMethod(VHC).value(CNT) <> 0 Then InfoMethod(VHC).Enabled(CNT) = True
    
    InfoMethod(LV).EqNum(CNT) = DBTbl("LVSEQN")
    InfoMethod(LV).Coeff(CNT, 1) = DBTbl("LVSA")
    InfoMethod(LV).Coeff(CNT, 2) = DBTbl("LVSB")
    InfoMethod(LV).Coeff(CNT, 3) = DBTbl("LVSC")
    InfoMethod(LV).Coeff(CNT, 4) = DBTbl("LVSD")
    InfoMethod(LV).Coeff(CNT, 5) = DBTbl("LVSE")
    InfoMethod(LV).MinT(CNT) = DBTbl("LVSTMIN")
    InfoMethod(LV).MaxT(CNT) = DBTbl("LVSTMAX")
    InfoMethod(LV).value(CNT) = CalcFofT(LV, CNT)
    If InfoMethod(LV).value(CNT) <> 0 Then InfoMethod(LV).Enabled(CNT) = True
    
    InfoMethod(VV).EqNum(CNT) = DBTbl("VVSEQN")
    InfoMethod(VV).Coeff(CNT, 1) = DBTbl("VVSA")
    InfoMethod(VV).Coeff(CNT, 2) = DBTbl("VVSB")
    InfoMethod(VV).Coeff(CNT, 3) = DBTbl("VVSC")
    InfoMethod(VV).Coeff(CNT, 4) = DBTbl("VVSD")
    InfoMethod(VV).Coeff(CNT, 5) = DBTbl("VVSE")
    InfoMethod(VV).MinT(CNT) = DBTbl("VVSTMIN")
    InfoMethod(VV).MaxT(CNT) = DBTbl("VVSTMAX")
    InfoMethod(VV).value(CNT) = CalcFofT(VV, CNT)
    If InfoMethod(VV).value(CNT) <> 0 Then InfoMethod(VV).Enabled(CNT) = True
    
    InfoMethod(LTC).EqNum(CNT) = DBTbl("LTCEQN")
    InfoMethod(LTC).Coeff(CNT, 1) = DBTbl("LTCA")
    InfoMethod(LTC).Coeff(CNT, 2) = DBTbl("LTCB")
    InfoMethod(LTC).Coeff(CNT, 3) = DBTbl("LTCC")
    InfoMethod(LTC).Coeff(CNT, 4) = DBTbl("LTCD")
    InfoMethod(LTC).Coeff(CNT, 5) = DBTbl("LTCE")
    InfoMethod(LTC).MinT(CNT) = DBTbl("LTCTMIN")
    InfoMethod(LTC).MaxT(CNT) = DBTbl("LTCTMAX")
    InfoMethod(LTC).value(CNT) = CalcFofT(LTC, CNT)
    If InfoMethod(LTC).value(CNT) <> 0 Then InfoMethod(LTC).Enabled(CNT) = True
    
    InfoMethod(VTC).EqNum(CNT) = DBTbl("VTCEQN")
    InfoMethod(VTC).Coeff(CNT, 1) = DBTbl("VTCA")
    InfoMethod(VTC).Coeff(CNT, 2) = DBTbl("VTCB")
    InfoMethod(VTC).Coeff(CNT, 3) = DBTbl("VTCC")
    InfoMethod(VTC).Coeff(CNT, 4) = DBTbl("VTCD")
    InfoMethod(VTC).Coeff(CNT, 5) = DBTbl("VTCE")
    InfoMethod(VTC).MinT(CNT) = DBTbl("VTCTMIN")
    InfoMethod(VTC).MaxT(CNT) = DBTbl("VTCTMAX")
    InfoMethod(VTC).value(CNT) = CalcFofT(VTC, CNT)
    If InfoMethod(VTC).value(CNT) <> 0 Then InfoMethod(VTC).Enabled(CNT) = True
    
    InfoMethod(ST).EqNum(CNT) = DBTbl("STEQN")
    InfoMethod(ST).Coeff(CNT, 1) = DBTbl("STA")
    InfoMethod(ST).Coeff(CNT, 2) = DBTbl("STB")
    InfoMethod(ST).Coeff(CNT, 3) = DBTbl("STC")
    InfoMethod(ST).Coeff(CNT, 4) = DBTbl("STD")
    InfoMethod(ST).Coeff(CNT, 5) = DBTbl("STE")
    InfoMethod(ST).MinT(CNT) = DBTbl("STTMIN")
    InfoMethod(ST).MaxT(CNT) = DBTbl("STTMAX")
    InfoMethod(ST).value(CNT) = CalcFofT(ST, CNT)
    If InfoMethod(ST).value(CNT) <> 0 Then InfoMethod(ST).Enabled(CNT) = True
            
    DBTbl.Close
    
    Exit Sub

Get801DataError:
    
    If Err = 94 Then Resume Next
    MsgBox "Error loading data from 801 database", 48, "Error"
    DBTbl.Close

End Sub



Sub Get911Data(CNT As Integer)
    
    ' REVISIONS:  DMW  6/6/97  - fixed this so it reads in the units as well and converts to Pearls default if necessary since
    '                               911 db doesn't always store values in default units
    '             DMW  3/13/98 - modified to check whether we're dealing with master.mdb (real 911 info) or user database (user values from dbman)
    Dim DBTbl As Recordset
    Dim Code As Integer
    Dim i As Integer
    Dim units_read(NumProperties) As String
    Dim prev_rating As Integer
    Dim temp_rating As Integer
    
    'Check to see if the user owns the 911 database
    If DIPPR911 = False Or Path911 = NULLPATH Then Exit Sub

    
    On Error GoTo Get911DataError

    Set DBTbl = DBJetMaster.OpenRecordset("DIPPR911", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey1"
    DBTbl.Seek "=", Cur_Info.CAS
    
    If DBTbl.NoMatch Then
         DBTbl.Close
         Exit Sub
    End If
        
    Do While Not DBTbl.EOF And DBTbl("CAS #") = Cur_Info.CAS
        Code = DBTbl("PEARLS Code")
        If Code <> -1 Then
        
            ' commented this statement out so it'll get the best value instead of the first
            
            'If infomethod(Code).Enabled(CNT) = False Then
                If Code = 2 Or Code = 6 Or Code = 8 Or Code = 9 Or Code = 12 Or Code = 18 Or Code = 19 Or Code = 20 Or Code = 21 Or Code = 22 Then
                    InfoMethod(Code).MinT(CNT) = DBTbl("Value")
                    units_read(Code) = ""
                    On Error Resume Next
                    units_read(Code) = DBTbl("Units")
                    On Error GoTo Get911DataError
                    InfoMethod(Code).MaxT(CNT) = DBTbl("Temperature")
                    InfoMethod(Code).Coeff(CNT, 1) = DBTbl("Coef1")
                    InfoMethod(Code).Coeff(CNT, 2) = DBTbl("Coef2")
                    InfoMethod(Code).Coeff(CNT, 3) = DBTbl("Coef3")
                    InfoMethod(Code).Coeff(CNT, 4) = DBTbl("Coef4")
                    InfoMethod(Code).Coeff(CNT, 5) = DBTbl("Coef5")
                    InfoMethod(Code).EqNum(CNT) = DBTbl("Equation")
                    InfoMethod(Code).value(CNT) = CalcFofT(Code, CNT)
                    If LCase(Trim(MasterDBName)) = "master.mdb" Then
                        InfoMethod(Code).MethodName(CNT) = "911 Database"
                    Else
                        InfoMethod(Code).MethodName(CNT) = DBTbl("Desc/Method") & "(" & MasterDBName & ")"
                    End If
                    If InfoMethod(Code).value(CNT) <> 0 Then InfoMethod(Code).Enabled(CNT) = True
                Else
                
                    temp_rating = DBTbl("Rating")
                    If InfoMethod(Code).value(CNT) <> 0 And prev_rating < temp_rating Then
                        'we already have a better value for this property, go to next property
                        GoTo next_iteration
                    End If
                    InfoMethod(Code).value(CNT) = DBTbl("Value")
                    On Error Resume Next
                    units_read(Code) = ""
                    prev_rating = -1    ' dummy value in case there's no entry here
                    prev_rating = DBTbl("Rating")
                    units_read(Code) = DBTbl("Units")
                    
                    On Error GoTo Get911DataError
                    ' shouldn't need this stuff anymore because we check for it after we're done reading in data
                    If Code = 33 And Cur_Info.OpT <> 298.15 Then
                        InfoMethod(Code).Enabled(CNT) = False
                        InfoMethod(Code).value(CNT) = 0
                    End If
                    
                    If InfoMethod(Code).value(CNT) <> 0 Then
                        InfoMethod(Code).Enabled(CNT) = True
                        If LCase(Trim(MasterDBName)) = "master.mdb" Then
                            InfoMethod(Code).MethodName(CNT) = "911 Database"
                        Else
                            InfoMethod(Code).MethodName(CNT) = MasterDBName
                        End If
                    ElseIf Code <= 54 And Code >= 41 And DBTbl("comment") <> "" Then
                        InfoMethod(Code).Enabled(CNT) = True
                        If LCase(Trim(MasterDBName)) = "master.mdb" Then
                            InfoMethod(Code).MethodName(CNT) = "911 Database"
                        Else
                            InfoMethod(Code).MethodName(CNT) = MasterDBName
                        End If
                    End If
                    
                End If
            End If
        'End If
next_iteration:
        If Not DBTbl.EOF Then
            DBTbl.MoveNext
        Else
            GoTo close_db
        End If
    Loop
close_db:
    DBTbl.Close
    
    ' Do the conversions if necessary to pearls default (not user default)(
    ' starting with MW and LD and OpT since we need those to convert other units
    ' IsDefault returns Boolean if the unit read is the PEARLS default unit (not user default)
    If Not IsDefault(Trim(units_read(MW)), MW) And Trim(units_read(MW)) <> "" Then
        InfoMethod(MW).value(CNT) = Convert(InfoMethod(MW).value(CNT), MW, units_read(MW), Get_DefaultUnit(MW), False)
        InfoMethod(MW).Unit = Get_DefaultUnit(MW)
    End If
    If Not IsDefault(Trim(units_read(LD)), LD) And Trim(units_read(LD)) <> "" Then
        InfoMethod(LD).value(CNT) = Convert(InfoMethod(LD).value(CNT), LD, units_read(LD), Get_DefaultUnit(LD), False)
        InfoMethod(LD).Unit = Get_DefaultUnit(LD)
    End If
    ' not doing OpT for now or MW or LD or f(t) props
    For i = 0 To NumProperties
        If i <> MW And i <> LD Then
        
            If Not IsDefault(Trim(units_read(i)), i) And units_read(i) <> "" Then
                InfoMethod(i).value(CNT) = Convert(InfoMethod(i).value(CNT), CLng(i), units_read(i), Get_DefaultUnit(i), False)
                'infomethod(I).Unit = ConvertToDefault(I)
                'infomethod(I).OpTUnit = ConvertToDefault(CLng(I))
                'FIX set OpTUnits
            End If
            InfoMethod(i).Unit = Get_DefaultUnit(i)
        End If
    Next i
    Exit Sub

Get911DataError:
   'MsgBox ("Code is " & Code)
    If Err = 94 Then
        Resume Next
    ElseIf Err = 3021 Then
        GoTo close_db
    End If
    If LCase(Trim(MasterDBName)) = "master.mdb" Then
        MsgBox "Error loading data from 911 database", 48, "Error"
    Else
        MsgBox "Error loading data from " & MasterDBName, 48, "Error"
    End If
    DBTbl.Close

End Sub


Function GetDataPrint() As Boolean

    Dim i As Integer
    Dim J As Integer
    Dim DBTbl As Recordset
        
    On Error GoTo GetUserDataError
    
    Set DBTbl = DBJetUser.OpenRecordset("SaveTable1", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    If DBTbl.NoMatch Then
        GetDataPrint = False
        DBTbl.Close
        Exit Function
    End If
    
    GetDataPrint = True
    
    Cur_Info.OpT = DBTbl("OpT")
    Cur_Info.OpP = DBTbl("OpP")
    Cur_Info.OpTUnit = DBTbl("OpTUnit")
    Cur_Info.OpPUnit = DBTbl("OpPUnit")
    BIPIndex(1) = DBTbl("BIPIndex(ACchem)")
    BIPIndex(2) = DBTbl("BIPIndex(logKow)")
    BIPIndex(3) = DBTbl("BIPIndex(Schem)")
    
    Set DBTbl = DBJetUser.OpenRecordset("SaveTable2", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    For i = 0 To NumProperties
        For J = 1 To NumMethods
            InfoMethod(i).MethodName(J) = DBTbl("Method Name")
            InfoMethod(i).Enabled(J) = DBTbl("Method Enabled")
            InfoMethod(i).CurMethod = DBTbl("Current Method Index")
            InfoMethod(i).value(J) = DBTbl("Value")
            InfoMethod(i).Unit = DBTbl("Unit")
            InfoMethod(i).TFT = DBTbl("TFT")
            InfoMethod(i).TFTUnit = DBTbl("TFTUnit")
            InfoMethod(i).EqNum(J) = DBTbl("EqNum")
            InfoMethod(i).Coeff(J, 1) = DBTbl("Coeff A")
            InfoMethod(i).Coeff(J, 2) = DBTbl("Coeff B")
            InfoMethod(i).Coeff(J, 3) = DBTbl("Coeff C")
            InfoMethod(i).Coeff(J, 4) = DBTbl("Coeff D")
            InfoMethod(i).Coeff(J, 5) = DBTbl("Coeff E")
            InfoMethod(i).MinT(J) = DBTbl("MinT")
            InfoMethod(i).MaxT(J) = DBTbl("MaxT")
            DBTbl.MoveNext
        Next J
    Next i
    DBTbl.Close
        
    Exit Function
    
GetUserDataError:
    
    If Err = 13 Then Resume Next
    MsgBox "Error loading data from user database", 48, "Error"
    DBTbl.Close

End Function

Sub GetMasterData(CNT As Integer)
    
    ' REVISIONS  DMW 6/8/97  - The HC data in the tables is all in Pa mol/mol
    '                               quick fix is to change it to kPa mol/mol here.
    '                               This should be changed when time allows to represent
    '                               data in database in proper units or read in units and do conversion
    Dim i As Integer
    Dim mySMILES As String
    Dim SearchResult As Byte
    Dim SearchType As Byte
    Dim DBTbl As Recordset
    Dim intSF_ID(0 To 99) As Long, intSF_Quant(0 To 99) As Long
    Dim intMF_ID(0 To 20) As Long, intMF_Quant(0 To 20) As Long
    Dim local_smiles As String
    Dim local_file As String
    Dim message As String
    
    Dim units_read As String
    If PathMaster = NULLPATH Then
        Exit Sub
    End If
    
    On Error GoTo GetMasterDataError
        
'    Cur_Info.OpT = Cur_Info.OpT - 273.15
    Cur_Info.OpT = Convert(Cur_Info.OpT, OptTemp, Cur_Info.OpTUnit, "C", False)
    Cur_Info.OpTUnit = "C"
    ' first reset the groups so we're not getting garbage
    For i = 1 To 15
        Cur_Info.Grp(i) = 0
        Cur_Info.NumGrp(i) = 0
    Next i
    Cur_Info.NumRings = 0
    Set DBTbl = DBJetMaster.OpenRecordset("UNIFAC Groups", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
    
        Cur_Info.NumRings = DBTbl("RG")
        Cur_Info.MaxGroups = DBTbl("MX")
        Cur_Info.Grp(1) = DBTbl("G1")
        Cur_Info.NumGrp(1) = DBTbl("N1")
        If Cur_Info.Grp(1) <= 0 Then
            GoTo try_shredder
        End If
        Cur_Info.Grp(2) = DBTbl("G2")
        Cur_Info.NumGrp(2) = DBTbl("N2")
        Cur_Info.Grp(3) = DBTbl("G3")
        Cur_Info.NumGrp(3) = DBTbl("N3")
        Cur_Info.Grp(4) = DBTbl("G4")
        Cur_Info.NumGrp(4) = DBTbl("N4")
        Cur_Info.Grp(5) = DBTbl("G5")
        Cur_Info.NumGrp(5) = DBTbl("N5")
        Cur_Info.Grp(6) = DBTbl("G6")
        Cur_Info.NumGrp(6) = DBTbl("N6")
        Cur_Info.Grp(7) = DBTbl("G7")
        Cur_Info.NumGrp(7) = DBTbl("N7")
        Cur_Info.Grp(8) = DBTbl("G8")
        Cur_Info.NumGrp(8) = DBTbl("N8")
        Cur_Info.Grp(9) = DBTbl("G9")
        Cur_Info.NumGrp(9) = DBTbl("N9")
        Cur_Info.Grp(10) = DBTbl("G10")
        Cur_Info.NumGrp(10) = DBTbl("N10")
        DBTbl.Close
    Else
        ' if there's no unifac groups for this chem in the database, try running shredder
try_shredder:

        Screen.MousePointer = 11
        local_smiles = Trim(Cur_Info.SMILES)
        SearchType = 2
        local_file = AddBackSlash(App.path) & "dat\unifac.dat"  'the file we're reading groups from (ie unifac.dat)
        
        If Not FileExists(local_file) Then
            MsgBox "Mosdap dat file '" & local_file & "' : doesn't exist"
            Exit Sub
        End If
        If Not FileExists(AddBackSlash(App.path) & "dlls\Mosdap32.dll") Then
            MsgBox "file : 'Mosdap32.dll' doesn't exist"
            Exit Sub
        End If
        
        Call MOSDAP(local_smiles, 0, local_file, "", SearchType, SearchResult, intSF_ID(0), intSF_Quant(0), intMF_ID(0), intMF_Quant(0))
        
        Select Case SearchResult
            Case 0
                'Call clear_struct_groups_frame
                message = "Unable to disassemble " & local_smiles
            Case 1, 2
                'Call fill_struct_groups_frame
                message = "Successfully disassembled"
            Case 2
                'Call fill_struct_groups_frame
                message = "Partially disassembled"
            Case Else
                'Call clear_struct_groups_frame
                message = "An error occurred in the code while disassembling"
        End Select
        MsgBox message

        If SearchResult <> 0 Then
            For i = 0 To 98
                If intSF_ID(i) <= 0 Or intSF_Quant(i) <= 0 Then
                    Cur_Info.MaxGroups = i
                    Exit For
                End If
                Cur_Info.Grp(i) = intSF_ID(i)
                Cur_Info.NumGrp(i) = intSF_Quant(i)
            Next i
        End If
        Screen.MousePointer = 0
        
    End If
    
    
    Set DBTbl = DBJetMaster.OpenRecordset("VP Yaws", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        InfoMethod(VP).EqNum(CNT) = 202
        InfoMethod(VP).Coeff(CNT, 1) = DBTbl("AntA")
        InfoMethod(VP).Coeff(CNT, 2) = DBTbl("AntB")
        InfoMethod(VP).Coeff(CNT, 3) = DBTbl("AntC")
        InfoMethod(VP).MinT(CNT) = DBTbl("MINT") + 273.15
        InfoMethod(VP).MaxT(CNT) = DBTbl("MAXT") + 273.15
        InfoMethod(VP).value(CNT) = CalcFofT(VP, CNT) * (101325 / 760)
        InfoMethod(VP).MethodName(CNT) = "Yaws"
    End If
    DBTbl.Close
    
    If InfoMethod(VP).CurMethod = CNT And InfoMethod(VP).Enabled(CNT) = False Then
        InfoMethod(VP).CurMethod = 0
    End If
    If InfoMethod(VP).CurMethod = 0 And InfoMethod(VP).Enabled(CNT) = True Then
        InfoMethod(VP).CurMethod = CNT
    End If
    
    If Cur_Info.OpT = 298.15 Then
        Set DBTbl = DBJetMaster.OpenRecordset("VP@25 Superfund", dbOpenTable)
        DBTbl.Index = "PrimaryKey"
        DBTbl.Seek "=", Cur_Info.CAS
        If Not DBTbl.NoMatch Then
            InfoMethod(VP25).value(CNT) = DBTbl("VP") * (101325 / 760)
            If InfoMethod(VP25).value(CNT) <> 0 Then
                InfoMethod(VP25).MethodName(CNT) = "Superfund"
                InfoMethod(VP25).Enabled(CNT) = True
            End If
        End If
        DBTbl.Close
    End If
    
    If InfoMethod(VP25).CurMethod = CNT And InfoMethod(VP25).Enabled(CNT) = False Then
        InfoMethod(VP25).CurMethod = 0
    End If
    If InfoMethod(VP25).CurMethod = 0 And InfoMethod(VP25).Enabled(CNT) = True Then
        InfoMethod(VP25).CurMethod = CNT
    End If
    
    If Cur_Info.OpT = 298.15 Then
        Set DBTbl = DBJetMaster.OpenRecordset("Kow@25 Superfund", dbOpenTable)
        DBTbl.Index = "PrimaryKey"
        DBTbl.Seek "=", Cur_Info.CAS
        If Not DBTbl.NoMatch Then
            InfoMethod(logKow).value(CNT) = DBTbl("log Kow")
            If InfoMethod(logKow).value(CNT) <> 0 And InfoMethod(logKow).value(CNT) <> ERROR_FLAG Then
                InfoMethod(logKow).MethodName(CNT) = "Superfund"
                InfoMethod(logKow).Enabled(CNT) = True
            End If
        End If
        DBTbl.Close
    End If
    
    If InfoMethod(logKow).CurMethod = CNT And InfoMethod(logKow).Enabled(CNT) = False Then
        InfoMethod(logKow).CurMethod = 0
    End If
    If InfoMethod(logKow).CurMethod = 0 And InfoMethod(logKow).Enabled(CNT) = True Then
        InfoMethod(logKow).CurMethod = CNT
    End If
    
    i = CNT
    'This if statement was added temporarily to
    'facilitate my data comparison
    If Cur_Info.CAS = 108883 And Cur_Info.OpT <> 25 Then
    Set DBTbl = DBJetMaster.OpenRecordset("HC RTI", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
            
                    ' a temporary fix, converts to kPa*mol/mol
                    
                    'infomethod(HC).VALUE(I) = DBTbl("HC") * 101325 / 1000
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC") * 101325, HC, "kPa*mol/mol", Get_DefaultUnit(HC), False)
                'End If
                If InfoMethod(HC).value(i) <> 0# And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "RTI"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    End If
    
    Set DBTbl = DBJetMaster.OpenRecordset("HC Yaws", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                '    infomethod(HC).VALUE(I) = DBTbl("HC")
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC"), HC, units_read, Get_DefaultUnit(HC), False)
                'End If
                    
                If InfoMethod(HC).value(i) <> 0 And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "Yaws"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveNext
          Else
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    
    Set DBTbl = DBJetMaster.OpenRecordset("HC Superfund", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'units_read = "Pa*mol/mol"
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                    ' temporary fix, converts to Pa*mol/mol
                    'infomethod(HC).VALUE(I) = DBTbl("HC") / 1.7784357E-10 '/ 1000
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC") / 1.7784357E-10, HC, "Pa*mol/mol", Get_DefaultUnit(HC), False) '  / 1.7784357E-10)
                'End If
                If InfoMethod(HC).value(i) <> 0 And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "Superfund"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveNext
          Else
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    
    Set DBTbl = DBJetMaster.OpenRecordset("HC Ashworth", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                '    infomethod(HC).VALUE(I) = DBTbl("HC")
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC"), HC, units_read, Get_DefaultUnit(HC), False)
                'End If
                If InfoMethod(HC).value(i) <> 0 And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "Ashworth"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveNext
          Else
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
        
    Set DBTbl = DBJetMaster.OpenRecordset("HC Stephenson", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                '    infomethod(HC).VALUE(I) = DBTbl("HC")
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC"), HC, units_read, Get_DefaultUnit(HC), False)
                'End If
                If InfoMethod(HC).value(i) <> 0 And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "Stephenson"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveNext
          Else
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    
    Set DBTbl = DBJetMaster.OpenRecordset("HC Gmehling", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS") Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                '    infomethod(HC).VALUE(I) = DBTbl("HC")
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC"), HC, units_read, Get_DefaultUnit(HC), False)
                'End If
                If InfoMethod(HC).value(i) <> 0 And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "Gmehling"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveNext
          Else
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    
    If i = 10 Then i = i + 1
    Set DBTbl = DBJetMaster.OpenRecordset("HC Chen and Wagner", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                '    infomethod(HC).VALUE(I) = DBTbl("HC")
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC"), HC, units_read, Get_DefaultUnit(HC), False)
                'End If
                    
                If InfoMethod(HC).value(i) <> 0# And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "Chen and Wagner"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveNext
          Else
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    
    If i = 10 Then i = i + 1
    Set DBTbl = DBJetMaster.OpenRecordset("HC Persichetti", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                '    infomethod(HC).VALUE(I) = DBTbl("HC")
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC"), HC, units_read, Get_DefaultUnit(HC), False)
                'End If
                If InfoMethod(HC).value(i) <> 0 And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "Persichetti"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveNext
          Else
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    
    If i = 10 Then i = i + 1
    Set DBTbl = DBJetMaster.OpenRecordset("HC Hwang", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                '    infomethod(HC).VALUE(I) = DBTbl("HC")
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC"), HC, units_read, Get_DefaultUnit(HC), False)
                'End If
                If InfoMethod(HC).value(i) <> 0# And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "Hwang"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            End If
            DBTbl.MoveNext
          Else
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    
    If i = 10 Then i = i + 1
    Set DBTbl = DBJetMaster.OpenRecordset("HC ASPEN (AENV)", dbOpenTable)
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    If Not DBTbl.NoMatch Then
        While Not (DBTbl.EOF)
          If Cur_Info.CAS = DBTbl("CAS").value Then
            If Abs(DBTbl("HCT") - Cur_Info.OpT) < 0.5 Then
                units_read = ""
                units_read = DBTbl("Units")
                'If IsDefault(Trim(units_read), HC) Or units_read = "" Then
                 '   infomethod(HC).VALUE(I) = DBTbl("HC")
                'Else
                    InfoMethod(HC).value(i) = Convert(DBTbl("HC"), HC, units_read, Get_DefaultUnit(HC), False)
                'End If
                If InfoMethod(HC).value(i) <> 0 And InfoMethod(HC).value(i) <> ERROR_FLAG Then
                    InfoMethod(HC).MethodName(i) = "ASPEN (AENV)"
                    InfoMethod(HC).Enabled(i) = True
                    i = i + 1
                End If
            DBTbl.MoveNext
            End If
            DBTbl.MoveLast
            DBTbl.MoveNext
          End If
        Wend
    End If
    DBTbl.Close
    
    If InfoMethod(HC).CurMethod = i And InfoMethod(HC).Enabled(i) = False Then
        InfoMethod(HC).CurMethod = 0
    End If
    If InfoMethod(HC).CurMethod = 0 And InfoMethod(HC).Enabled(i) = True Then
        InfoMethod(HC).CurMethod = i
    End If
    
    If Cur_Info.OpT = 25 Then
        Set DBTbl = DBJetMaster.OpenRecordset("SB@25 Superfund", dbOpenTable)
        DBTbl.Index = "PrimaryKey"
        DBTbl.Seek "=", Cur_Info.CAS
        If Not DBTbl.NoMatch Then
            For i = 1 To 10
                If InfoMethod(MW).value(i) <> 0 Then
                    units_read = ""
                    units_read = DBTbl("Units")
                    'If IsDefault(Trim(units_read), Schem) Or units_read = "" Then
                        ' already in ppm(wt), no need for conversion ??
                       ' infomethod(Schem).VALUE(CNT) = DBTbl("Sol")
                    'Else
                        InfoMethod(Schem).value(CNT) = Convert(DBTbl("Sol"), Schem, units_read, Get_DefaultUnit(Schem), False)
                        'infomethod(Schem).VALUE(CNT) = DBTbl("Sol") / (infomethod(MW).VALUE(I) * 1000)
                    'End If
                End If
            Next i
            If InfoMethod(Schem).value(CNT) <> 0 And InfoMethod(Schem).value(CNT) <> ERROR_FLAG Then
                InfoMethod(Schem).MethodName(CNT) = "Superfund"
                InfoMethod(Schem).Enabled(CNT) = True
            End If
        End If
        DBTbl.Close
    End If
    
    If InfoMethod(Schem).CurMethod = CNT And InfoMethod(Schem).Enabled(CNT) = False Then
        InfoMethod(Schem).CurMethod = 0
    End If
    If InfoMethod(Schem).CurMethod = 0 And InfoMethod(Schem).Enabled(CNT) = True Then
        InfoMethod(Schem).CurMethod = CNT
    End If
    
    If Cur_Info.OpT = 25 Then
        Set DBTbl = DBJetMaster.OpenRecordset("SB@25 Yaws", dbOpenTable)
        DBTbl.Index = "PrimaryKey"
        DBTbl.Seek "=", Cur_Info.CAS
        If Not DBTbl.NoMatch Then
            For i = 1 To 10
                If InfoMethod(MW).value(i) <> 0 Then
                    ' already in ppm(wt), no need for conversion ??
                    units_read = ""
                    units_read = DBTbl("Units")
                    'If IsDefault(Trim(units_read), Schem) Or units_read = "" Then
                    '    infomethod(Schem).VALUE(CNT + 1) = DBTbl("Sol")
                    'Else
                        InfoMethod(Schem).value(CNT + 1) = Convert(DBTbl("Sol"), Schem, units_read, Get_DefaultUnit(Schem), False)
                    'infomethod(Schem).VALUE(CNT + 1) = DBTbl("Sol") / (infomethod(MW).VALUE(I) * 1000)
                    'End If
                End If
            Next i
            If InfoMethod(Schem).value(CNT + 1) <> 0 And InfoMethod(Schem).value(CNT + 1) <> ERROR_FLAG Then
                InfoMethod(Schem).MethodName(CNT + 1) = "Yaws"
                InfoMethod(Schem).Enabled(CNT + 1) = True
            End If
        End If
        DBTbl.Close
    End If
    
    If InfoMethod(Schem).CurMethod = CNT + 1 And InfoMethod(Schem).Enabled(CNT + 1) = False Then
        InfoMethod(Schem).CurMethod = 0
    End If
    If InfoMethod(Schem).CurMethod = 0 And InfoMethod(Schem).Enabled(CNT + 1) = True Then
        InfoMethod(Schem).CurMethod = CNT + 1
    End If
    
    ' DENISE 6/4/97: I think this stuff is outdated by block 5 stuff
    'Set DBTbl = DBJetMaster.OpenRecordset("MTU Fire and Explosion", dbOpenTable)
    'DBTbl.Index = "PrimaryKey"
    'DBTbl.Seek "=", Cur_Info.CAS
    'If Not DBTbl.NoMatch Then
       ' infomethod(FP).VALUE(CNT) = DBTbl("FP_MTU912")
       ' If infomethod(FP).VALUE(CNT) <> 0 Then
          '  infomethod(FP).MethodName(CNT) = "MTU DIPPR"
          '  infomethod(FP).Enabled(CNT) = True
       'End If
        'infomethod(LFL).VALUE(CNT) = DBTbl("LFL_MTU912")
        'If infomethod(LFL).VALUE(CNT) <> 0 Then
          '  infomethod(LFL).MethodName(CNT) = "MTU DIPPR"
          '  infomethod(LFL).Enabled(CNT) = True
        'End If
        'infomethod(UFL).VALUE(CNT) = DBTbl("UFL_MTU")
        'If infomethod(UFL).VALUE(CNT) <> 0 Then
         '   infomethod(UFL).MethodName(CNT) = "MTU DIPPR"
         '   infomethod(UFL).Enabled(CNT) = True
        'End If
    'End If
    'DBTbl.Close
    
    'If infomethod(FP).CurMethod = CNT And infomethod(FP).Enabled(CNT) = False Then
    '    infomethod(FP).CurMethod = 0
    'End If
    'If infomethod(FP).CurMethod = 0 And infomethod(FP).Enabled(CNT) = True Then
    '    infomethod(FP).CurMethod = CNT
    'End If
        
    'If infomethod(LFL).CurMethod = CNT And infomethod(LFL).Enabled(CNT) = False Then
    '    infomethod(LFL).CurMethod = 0
    'End If
    'If infomethod(LFL).CurMethod = 0 And infomethod(LFL).Enabled(CNT) = True Then
    '    infomethod(LFL).CurMethod = CNT
    'End If
        
    'If infomethod(UFL).CurMethod = CNT And infomethod(UFL).Enabled(CNT) = False Then
    '    infomethod(UFL).CurMethod = 0
    'End If
    'If infomethod(UFL).CurMethod = 0 And infomethod(UFL).Enabled(CNT) = True Then
    '    infomethod(UFL).CurMethod = CNT
   'End If
        
    'Cur_Info.OpT = Cur_Info.OpT + 273.15
    Cur_Info.OpT = Convert(Cur_Info.OpT, OptTemp, Cur_Info.OpTUnit, "K", False)
    Cur_Info.OpTUnit = "K"
        
    Exit Sub
    
GetMasterDataError:
    ' Err = 3021 NO longer handled due to vb5 different handling
    
    If Err = 94 Then Resume Next
    MsgBox "Error loading data from Master database", 48, "Error"
    DBTbl.Close
    
'    Cur_Info.OpT = Cur_Info.OpT + 273.15
    Cur_Info.OpT = Convert(Cur_Info.OpT, OptTemp, Cur_Info.OpTUnit, "K", False)
    Cur_Info.OpTUnit = "K"

End Sub











Function GetPropGroup(Code As Integer) As Integer

    Select Case Code
        Case MW
            GetPropGroup = 1
        Case LD25
            GetPropGroup = 1
        Case LD
            GetPropGroup = 1
        Case mp
            GetPropGroup = 1
        Case NBP
            GetPropGroup = 1
        Case VP25
            GetPropGroup = 1
        Case VP
            GetPropGroup = 1
        Case hfor
            GetPropGroup = 1
        Case LHC
            GetPropGroup = 2
        Case VHC
            GetPropGroup = 2
        Case Hvap25
            GetPropGroup = 2
        Case HvapNBP
            GetPropGroup = 2
        Case Hvap
            GetPropGroup = 2
        Case CT
            GetPropGroup = 2
        Case CP
            GetPropGroup = 2
        Case CV
            GetPropGroup = 2
        Case Dwater
            GetPropGroup = 3
        Case Dair
            GetPropGroup = 3
        Case ST25
            GetPropGroup = 3
        Case ST
            GetPropGroup = 3
        Case VV
            GetPropGroup = 3
        Case LV
            GetPropGroup = 3
        Case LTC
            GetPropGroup = 3
        Case VTC
            GetPropGroup = 3
        Case UFL
            GetPropGroup = 4
        Case LFL
            GetPropGroup = 4
        Case FP
            GetPropGroup = 4
        Case AIT
            GetPropGroup = 4
        Case Hcomb
            GetPropGroup = 4
        Case ThODcarb
            GetPropGroup = 5
        Case ThODcomb
            GetPropGroup = 5
        Case COD
            GetPropGroup = 5
        Case BOD
            GetPropGroup = 5
        Case ACwater
            GetPropGroup = 6
        Case HC
            GetPropGroup = 6
        Case ACchem
            GetPropGroup = 6
        Case logKow
            GetPropGroup = 6
        Case logKoc
            GetPropGroup = 6
        Case BCF
            GetPropGroup = 6
        Case Schem
            GetPropGroup = 6
        Case Swater
            GetPropGroup = 6
    End Select
    
End Function
Function GetPropName(Code As Integer) As String

    Select Case Code
        Case MW
            GetPropName = "Molecular Weight"
        Case LD25
            GetPropName = "Liquid Density @ 298.15 K"
        Case LD
            GetPropName = "Liquid Density as f(T)"
        Case mp
            GetPropName = "Melting Point"
        Case NBP
            GetPropName = "Normal Boiling Point (NBP)"
        Case VP25
            GetPropName = "Vapor Pressure @ 298.15 K"
        Case VP
            GetPropName = "Vapor Pressure as f(T)"
        Case hfor
            GetPropName = "Heat of Formation"
        Case LHC
            GetPropName = "Liquid Heat Capacity"
        Case VHC
            GetPropName = "Vapor Heat Capacity"
        Case Hvap25
            GetPropName = "Heat of Vaporization @ 298.15 K"
        Case HvapNBP
            GetPropName = "Heat of Vaporization @ NBP"
        Case Hvap
            GetPropName = "Heat of Vaporization as f(T)"
        Case CT
            GetPropName = "Critical Temperature"
        Case CP
            GetPropName = "Critical Pressure"
        Case CV
            GetPropName = "Critical Volume"
        Case Dwater
            GetPropName = "Diffusivity in Water"
        Case Dair
            GetPropName = "Diffusivity in Air"
        Case ST25
            GetPropName = "Surface Tension @ 298.15 K"
        Case ST
            GetPropName = "Surface Tension as f(T)"
        Case VV
            GetPropName = "Vapor Viscosity as f(T)"
        Case LV
            GetPropName = "Liquid Viscosity as f(T)"
        Case LTC
            GetPropName = "Liquid Thermal Conductivity as f(T)"
        Case VTC
            GetPropName = "Vapor Thermal Conductivity as f(T)"
        Case UFL
            GetPropName = "Upper Flammibility Limit"
        Case LFL
            GetPropName = "Lower Flammibility Limit"
        Case FP
            GetPropName = "Flash Point"
        Case AIT
            GetPropName = "Autoignition Temperature"
        Case Hcomb
            GetPropName = "Heat of Combustion"
        Case ThODcarb
            GetPropName = "Carbonaceous ThOD"
        Case ThODcomb
            GetPropName = "Combined ThOD"
        Case COD
            GetPropName = "Chemical Oxygen Demand"
        Case BOD
            GetPropName = "Biochemical Oxygen Demand"
        Case ACwater
            GetPropName = "Activity Coefficient of Water in Chemical"
        Case HC
            GetPropName = "Henry's Constant"
        Case ACchem
            GetPropName = "Activity Coefficient of Chemical in Water"
        Case logKow
            GetPropName = "log Kow"
        Case logKoc
            GetPropName = "log Koc"
        Case BCF
            GetPropName = "Bioconcentration Factor"
        Case Schem
            GetPropName = "Solubility of Chemical in Water"
        Case Swater
            GetPropName = "Solubility of Water in Chemical"
        Case Fat48E
            GetPropName = "Fathead Minnow, 48h, EC50"
        Case Fat96E
            GetPropName = "Fathead Minnow, 96h, EC50"
        Case Fat24L
            GetPropName = "Fathead Minnow, 24h, LC50"
        Case Fat48L
            GetPropName = "Fathead Minnow, 48h, LC50"
        Case Fat96L
            GetPropName = "Fathead Minnow, 96h, LC50"
        Case Sal24L
            GetPropName = "Salmonidae, 24h, LC50"
        Case Sal48L
            GetPropName = "Salmonidae, 48h, LC50"
        Case Sal96L
            GetPropName = "Salmonidae, 96h, LC50"
        Case Daph24E
            GetPropName = "Daphnia Magna, 24h, EC50"
        Case Daph48E
            GetPropName = "Daphnia Magna, 48h, EC50"
        Case Daph24L
            GetPropName = "Daphnia Magna, 24h, LC50"
        Case Daph48L
            GetPropName = "Daphnia Magna, 48h, LC50"
        Case Mysid96L
            GetPropName = "Mysid, 96h, LC50"
        Case AltSpecies
            GetPropName = "Alternate Species"
    End Select

End Function


Function GetUserData() As Boolean

    Dim property_name As String
    Dim prop_index As Integer
    Dim i As Integer
    Dim DBTbl As Recordset
        
    On Error GoTo GetUserDataError
    
    Set DBTbl = DBJetUser.OpenRecordset("SaveTable1", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    If DBTbl.NoMatch Then
        GetUserData = False
        DBTbl.Close
        Exit Function
    End If
    
    GetUserData = True
    
    Cur_Info.OpT = DBTbl("OpT")
    Cur_Info.OpP = DBTbl("OpP")
    Cur_Info.OpTUnit = DBTbl("OpTUnit")
    Cur_Info.OpPUnit = DBTbl("OpPUnit")
    BIPIndex(1) = DBTbl("BIPIndex(ACchem)")
    BIPIndex(2) = DBTbl("BIPIndex(logKow)")
    BIPIndex(3) = DBTbl("BIPIndex(Schem)")
    
    Set DBTbl = DBJetUser.OpenRecordset("SaveTable2", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    For i = 0 To NumProperties
        InfoMethod(i).CurMethod = DBTbl("Current Method Index")
        InfoMethod(i).TFT = DBTbl("TFT")
        InfoMethod(i).value(10) = DBTbl("User Value")
        InfoMethod(i).EqNum(10) = DBTbl("User EqNum")
        InfoMethod(i).Coeff(10, 1) = DBTbl("User Coeff A")
        InfoMethod(i).Coeff(10, 2) = DBTbl("User Coeff B")
        InfoMethod(i).Coeff(10, 3) = DBTbl("User Coeff C")
        InfoMethod(i).Coeff(10, 4) = DBTbl("User Coeff D")
        InfoMethod(i).Coeff(10, 5) = DBTbl("User Coeff E")
        InfoMethod(i).MinT(10) = DBTbl("User MinT")
        InfoMethod(i).MaxT(10) = DBTbl("User MaxT")
        DBTbl.MoveNext
    Next i
    DBTbl.Close
        ' save block 5 preferences DENISE change this is erroring
    On Error GoTo after_B5_data ' the table's not here, skip the block 5 stuff
    Set DBTbl = DBJetUser.OpenRecordset("Block5pref", dbOpenTable)
    For prop_index = 1 To 4
        DBTbl.Index = "PrimaryKey"
        If prop_index = 1 Then
            property_name = "UFL"
        ElseIf prop_index = 2 Then
            property_name = "LFL"
        ElseIf prop_index = 3 Then
            property_name = "FP"
        ElseIf prop_index = 4 Then
            property_name = "AIT"
        Else
            GoTo next_B5_data
        End If
        DBTbl.Seek "=", property_name
        If DBTbl.NoMatch Then
            GoTo next_B5_data ' try the next property
        Else
            B5Preference(prop_index - 1, 0) = DBTbl("Method1")
            B5Preference(prop_index - 1, 1) = DBTbl("Method2")
            B5Preference(prop_index - 1, 2) = DBTbl("Method3")
            B5Preference(prop_index - 1, 3) = DBTbl("Method4")
            B5Preference(prop_index - 1, 4) = DBTbl("Method5")
            B5Preference(prop_index - 1, 5) = DBTbl("Method6")
            B5Preference(prop_index - 1, 6) = DBTbl("Method7")
        End If
next_B5_data:
    Next prop_index
    DBTbl.Close
after_B5_data:
        
    GetUserData = True
    Exit Function
    
GetUserDataError:
    
    MsgBox "Error loading data from user database", 48, "Error"
    DBTbl.Close
    
End Function


Sub SaveDataPrint()

    Dim i As Integer
    Dim J As Integer
    Dim DBTbl As Recordset
    
    On Error GoTo SaveUserDataError
    
    If Cur_Info.CAS = 0 Then Exit Sub
    
    Set DBTbl = DBJetUser.OpenRecordset("SaveTable1", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    If DBTbl.NoMatch Then
        DBTbl.AddNew
        DBTbl("CAS") = Cur_Info.CAS
        DBTbl("Name") = Cur_Info.name
        DBTbl("Formula") = Cur_Info.Formula
        DBTbl("SMILES") = Cur_Info.SMILES
        DBTbl("Source") = Cur_Info.source
        DBTbl("Family") = Cur_Info.Family
        DBTbl("OpT") = FormatVal(Cur_Info.OpT)
        DBTbl("OpP") = FormatVal(Cur_Info.OpP)
        DBTbl("OpTUnit") = Cur_Info.OpTUnit
        DBTbl("OpPUnit") = Cur_Info.OpPUnit
        DBTbl("BIPIndex(ACchem)") = BIPIndex(1)
        DBTbl("BIPIndex(logKow)") = BIPIndex(2)
        DBTbl("BIPIndex(Schem)") = BIPIndex(3)
        DBTbl.Update
    Else
        DBTbl.Edit
        DBTbl("CAS") = Cur_Info.CAS
        DBTbl("Name") = Cur_Info.name
        DBTbl("Formula") = Cur_Info.Formula
        DBTbl("SMILES") = Cur_Info.SMILES
        DBTbl("Source") = Cur_Info.source
        DBTbl("Family") = Cur_Info.Family
        DBTbl("OpT") = FormatVal(Cur_Info.OpT)
        DBTbl("OpP") = FormatVal(Cur_Info.OpP)
        DBTbl("OpTUnit") = Cur_Info.OpTUnit
        DBTbl("OpPUnit") = Cur_Info.OpPUnit
        DBTbl("BIPIndex(ACchem)") = BIPIndex(1)
        DBTbl("BIPIndex(logKow)") = BIPIndex(2)
        DBTbl("BIPIndex(Schem)") = BIPIndex(3)
        DBTbl.Update
    End If
    DBTbl.Close
            
    Set DBTbl = DBJetUser.OpenRecordset("SaveTable2", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    If DBTbl.NoMatch Then
        For i = 0 To NumProperties
            For J = 1 To NumMethods
                DBTbl.AddNew
                DBTbl("CAS") = Cur_Info.CAS
                DBTbl("Property Name") = GetPropName(i)
                DBTbl("Property Number") = i
                DBTbl("Property Group") = GetPropGroup(i)
                DBTbl("Method Name") = InfoMethod(i).MethodName(J)
                DBTbl("Method Number") = J
                DBTbl("Method Enabled") = InfoMethod(i).Enabled(J)
                DBTbl("Current Method Index") = InfoMethod(i).CurMethod
                DBTbl("Value") = FormatVal(InfoMethod(i).value(J))
                DBTbl("Unit") = InfoMethod(i).Unit
                DBTbl("TFT") = FormatVal(InfoMethod(i).TFT)
                DBTbl("TFTUnit") = InfoMethod(i).TFTUnit
                DBTbl("EqNum") = InfoMethod(i).EqNum(J)
                DBTbl("Coeff A") = FormatVal(InfoMethod(i).Coeff(J, 1))
                DBTbl("Coeff B") = FormatVal(InfoMethod(i).Coeff(J, 2))
                DBTbl("Coeff C") = FormatVal(InfoMethod(i).Coeff(J, 3))
                DBTbl("Coeff D") = FormatVal(InfoMethod(i).Coeff(J, 4))
                DBTbl("Coeff E") = FormatVal(InfoMethod(i).Coeff(J, 5))
                DBTbl("MinT") = FormatVal(InfoMethod(i).MinT(J))
                DBTbl("MaxT") = FormatVal(InfoMethod(i).MaxT(J))
                DBTbl.Update
            Next J
        Next i
        DBTbl.Close
        Exit Sub
    Else
        For i = 0 To NumProperties
            For J = 1 To NumMethods
                DBTbl.Edit
                DBTbl("CAS") = Cur_Info.CAS
                DBTbl("Property Name") = GetPropName(i)
                DBTbl("Property Number") = i
                DBTbl("Property Group") = GetPropGroup(i)
                DBTbl("Method Name") = InfoMethod(i).MethodName(J)
                DBTbl("Method Number") = J
                DBTbl("Method Enabled") = InfoMethod(i).Enabled(J)
                DBTbl("Current Method Index") = InfoMethod(i).CurMethod
                DBTbl("Value") = FormatVal(InfoMethod(i).value(J))
                DBTbl("Unit") = InfoMethod(i).Unit
                DBTbl("TFT") = FormatVal(InfoMethod(i).TFT)
                DBTbl("TFTUnit") = InfoMethod(i).TFTUnit
                DBTbl("EqNum") = InfoMethod(i).EqNum(J)
                DBTbl("Coeff A") = FormatVal(InfoMethod(i).Coeff(J, 1))
                DBTbl("Coeff B") = FormatVal(InfoMethod(i).Coeff(J, 2))
                DBTbl("Coeff C") = FormatVal(InfoMethod(i).Coeff(J, 3))
                DBTbl("Coeff D") = FormatVal(InfoMethod(i).Coeff(J, 4))
                DBTbl("Coeff E") = FormatVal(InfoMethod(i).Coeff(J, 5))
                DBTbl("MinT") = FormatVal(InfoMethod(i).MinT(J))
                DBTbl("MaxT") = FormatVal(InfoMethod(i).MaxT(J))
                DBTbl.Update
            Next J
        Next i
        DBTbl.Close
    End If
    
    Exit Sub
    
SaveUserDataError:
    
    MsgBox "Error saving data to user database", 48, "Error"
    DBTbl.Close

End Sub
Sub SaveUserData()

    Dim property_name As String
    Dim i As Integer
    Dim prop_index As Integer
    Dim DBTbl As Recordset
    
    
    ' DENISE fix
    If Cur_Info.CAS = 0 Then Exit Sub
    
    On Error GoTo SaveUserDataError_closed
    Set DBTbl = DBJetUser.OpenRecordset("SaveTable1", dbOpenTable)
    On Error GoTo SaveUserDataError_open
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    If DBTbl.NoMatch Then
        DBTbl.AddNew
        DBTbl("CAS") = Cur_Info.CAS
        DBTbl("OpT") = Cur_Info.OpT
        DBTbl("OpP") = Cur_Info.OpP
        DBTbl("OpTUnit") = Cur_Info.OpTUnit
        DBTbl("OpPUnit") = Cur_Info.OpPUnit
        DBTbl("BIPIndex(ACchem)") = BIPIndex(1)
        DBTbl("BIPIndex(logKow)") = BIPIndex(2)
        DBTbl("BIPIndex(Schem)") = BIPIndex(3)
        DBTbl.Update
    Else
        DBTbl.Edit
        DBTbl("CAS") = Cur_Info.CAS
        DBTbl("OpT") = Cur_Info.OpT
        DBTbl("OpP") = Cur_Info.OpP
        DBTbl("OpTUnit") = Cur_Info.OpTUnit
        DBTbl("OpPUnit") = Cur_Info.OpPUnit
        DBTbl("BIPIndex(ACchem)") = BIPIndex(1)
        DBTbl("BIPIndex(logKow)") = BIPIndex(2)
        DBTbl("BIPIndex(Schem)") = BIPIndex(3)
        DBTbl.Update
    End If
    DBTbl.Close
            
    On Error GoTo SaveUserDataError_closed
    Set DBTbl = DBJetUser.OpenRecordset("SaveTable2", dbOpenTable)
    On Error GoTo SaveUserDataError_open
    
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", Cur_Info.CAS
    
    If DBTbl.NoMatch Then
        For i = 0 To NumProperties
            DBTbl.AddNew
            DBTbl("CAS") = Cur_Info.CAS
            DBTbl("Current Method Index") = InfoMethod(i).CurMethod
            DBTbl("TFT") = InfoMethod(i).TFT
            DBTbl("User Value") = InfoMethod(i).value(10)
            DBTbl("User EqNum") = InfoMethod(i).EqNum(10)
            DBTbl("User Coeff A") = InfoMethod(i).Coeff(10, 1)
            DBTbl("User Coeff B") = InfoMethod(i).Coeff(10, 2)
            DBTbl("User Coeff C") = InfoMethod(i).Coeff(10, 3)
            DBTbl("User Coeff D") = InfoMethod(i).Coeff(10, 4)
            DBTbl("User Coeff E") = InfoMethod(i).Coeff(10, 5)
            DBTbl("User MinT") = InfoMethod(i).MinT(10)
            DBTbl("User MaxT") = InfoMethod(i).MaxT(10)
            DBTbl.Update
        Next i
        DBTbl.Close
    Else
        For i = 0 To NumProperties
            DBTbl.Edit
            DBTbl("CAS") = Cur_Info.CAS
            DBTbl("Current Method Index") = InfoMethod(i).CurMethod
            DBTbl("TFT") = InfoMethod(i).TFT
            DBTbl("User Value") = InfoMethod(i).value(10)
            DBTbl("User EqNum") = InfoMethod(i).EqNum(10)
            DBTbl("User Coeff A") = InfoMethod(i).Coeff(10, 1)
            DBTbl("User Coeff B") = InfoMethod(i).Coeff(10, 2)
            DBTbl("User Coeff C") = InfoMethod(i).Coeff(10, 3)
            DBTbl("User Coeff D") = InfoMethod(i).Coeff(10, 4)
            DBTbl("User Coeff E") = InfoMethod(i).Coeff(10, 5)
            DBTbl("User MinT") = InfoMethod(i).MinT(10)
            DBTbl("User MaxT") = InfoMethod(i).MaxT(10)
            DBTbl.Update
            DBTbl.MoveNext
        Next i
        DBTbl.Close
    End If
    
    On Error GoTo SaveUserDataError_closed
    Set DBTbl = DBJetUser.OpenRecordset("PrefBIPHierarchy", dbOpenTable)
    On Error GoTo SaveUserDataError_open
    
    DBTbl.MoveFirst
    DBTbl.Edit
    DBTbl("BIP 1") = BIPHierarchy(1, 1)
    DBTbl("BIP 2") = BIPHierarchy(1, 2)
    DBTbl("BIP 3") = BIPHierarchy(1, 3)
    DBTbl("BIP 4") = BIPHierarchy(1, 4)
    DBTbl.Update
    DBTbl.MoveNext
    DBTbl.Edit
    DBTbl("BIP 1") = BIPHierarchy(2, 1)
    DBTbl("BIP 2") = BIPHierarchy(2, 2)
    DBTbl("BIP 3") = BIPHierarchy(2, 3)
    DBTbl("BIP 4") = BIPHierarchy(2, 4)
    DBTbl.Update
    DBTbl.MoveNext
    DBTbl.Edit
    DBTbl("BIP 1") = BIPHierarchy(3, 1)
    DBTbl("BIP 2") = BIPHierarchy(3, 2)
    DBTbl("BIP 3") = BIPHierarchy(3, 3)
    DBTbl("BIP 4") = BIPHierarchy(3, 4)
    DBTbl.Update
    DBTbl.Close
    
        ' save block 5 preferences DENISE change erroring
    On Error GoTo SaveUserDataError_closed
    Set DBTbl = DBJetUser.OpenRecordset("Block5pref", dbOpenTable)
    On Error GoTo SaveUserDataError_open
    
For prop_index = 1 To 4
DBTbl.Index = "PrimaryKey"
If prop_index = 1 Then
    property_name = "UFL"
ElseIf prop_index = 2 Then
    property_name = "LFL"
ElseIf prop_index = 3 Then
    property_name = "FP"
ElseIf prop_index = 4 Then
    property_name = "AIT"
End If
DBTbl.Seek "=", property_name
If DBTbl.NoMatch Then
    
        DBTbl.AddNew
        DBTbl("Property") = property_name
        DBTbl("Method1") = B5Preference(prop_index - 1, 0)
        DBTbl("Method2") = B5Preference(prop_index - 1, 1)
        DBTbl("Method3") = B5Preference(prop_index - 1, 2)
        DBTbl("Method4") = B5Preference(prop_index - 1, 3)
        DBTbl("Method5") = B5Preference(prop_index - 1, 4)
        DBTbl("Method6") = B5Preference(prop_index - 1, 5)
        DBTbl("Method7") = B5Preference(prop_index - 1, 6)
        DBTbl.Update
    
Else
        DBTbl.Edit
        DBTbl("Method1") = B5Preference(prop_index - 1, 0)
        DBTbl("Method2") = B5Preference(prop_index - 1, 1)
        DBTbl("Method3") = B5Preference(prop_index - 1, 2)
        DBTbl("Method4") = B5Preference(prop_index - 1, 3)
        DBTbl("Method5") = B5Preference(prop_index - 1, 4)
        DBTbl("Method6") = B5Preference(prop_index - 1, 5)
        DBTbl("Method7") = B5Preference(prop_index - 1, 6)
        DBTbl.Update
        
    
End If
Next prop_index
    DBTbl.Close
    'Save current preferences
    
    On Error GoTo SaveUserDataError_closed
    Set DBTbl = DBJetUser.OpenRecordset("PrefFormatting", dbOpenTable)
    On Error GoTo SaveUserDataError_open
    
    DBTbl.MoveFirst
    DBTbl.Edit
    DBTbl("Setting") = FormatGT1000
    DBTbl.Update
    DBTbl.MoveNext
    DBTbl.Edit
    DBTbl("Setting") = FormatLT001
    DBTbl.Update
    DBTbl.MoveNext
    DBTbl.Edit
    DBTbl("Setting") = FormatGeneral
    DBTbl.Update
    DBTbl.Close
    
    On Error GoTo SaveUserDataError_closed
    Set DBTbl = DBJetUser.OpenRecordset("PrefDefaultUnits", dbOpenTable)
    On Error GoTo SaveUserDataError_open
    
    DBTbl.MoveFirst
    For i = 0 To NumProperties
        DBTbl.Edit
        DBTbl("Default Unit") = DefaultUnit(i)
        DBTbl.Update
        DBTbl.MoveNext
    Next i
        
    DBTbl.Edit
    DBTbl("Default Unit") = DefaultTFTUnit
    DBTbl.Update
    DBTbl.Close
           
    Exit Sub
SaveUserDataError_closed:
    MsgBox "Error saving data to user database", 48, "Error"
    Exit Sub
SaveUserDataError_open:
    
    MsgBox "Error saving data to user database", 48, "Error"
    DBTbl.Close
    Err = 1
End Sub

