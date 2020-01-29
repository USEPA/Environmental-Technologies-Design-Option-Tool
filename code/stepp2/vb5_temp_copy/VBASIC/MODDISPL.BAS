Attribute VB_Name = "moddisplay"
Option Explicit
Sub CenterForm(FRM As Form)
    
    'This takes a form and centers it on the screen
    
    'Check if screen is maximized or minimized
    If FRM.WindowState = 0 Then
       FRM.Move (Screen.Width - FRM.Width) / 2, (Screen.Height - FRM.Height) / 2
    End If

End Sub



Sub DisplayProps()
    
    Dim i As Integer
    
    'Update the values in the boxes
    For i = 0 To NumProperties
        FRMMain!PNLPropVal(i).caption = ""
        FRMMain!PNLPropUnits(i).caption = ""
        If DispMethod(i).Enabled(DispMethod(i).CurMethod) = True Then
            If FormatVal(DispMethod(i).value(DispMethod(i).CurMethod)) = "" And i >= 41 Then
                FRMMain!PNLPropVal(i).caption = "See Comment"
            Else
                FRMMain!PNLPropVal(i).caption = FormatVal(DispMethod(i).value(DispMethod(i).CurMethod))
            End If
            
            FRMMain!PNLPropUnits(i).caption = DispMethod(i).Unit
            Select Case i
                Case LD:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case VP:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case Hvap:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case ST:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case VV:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case LV:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case VTC:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case LTC:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case VHC:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case LHC:
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at " & DispMethod(i).TFT & " " & DispMethod(i).TFTUnit
                Case ACchem
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at inf. dilution"
                Case ACwater
                    FRMMain!PNLPropUnits(i).caption = RTrim$(FRMMain!PNLPropUnits(i).caption) & " at inf. dilution"
            End Select
        End If
    Next i

    'Display the operating temperature and pressure
    FRMMain!TXTOpT.Text = FormatVal(Cur_Disp.OpT)
    FRMMain!LBLOpTUnits.caption = Cur_Disp.OpTUnit
    FRMMain!TXTOpP.Text = FormatVal(Cur_Disp.OpP)
    FRMMain!LBLOpPUnits.caption = Cur_Disp.OpPUnit

End Sub
Sub DisplayOneProp(Code As Integer)
        
    If Code = OptTemp Then
        FRMMain!TXTOpT.Text = FormatVal(Cur_Disp.OpT)
        FRMMain!LBLOpTUnits.caption = Cur_Disp.OpTUnit
        Exit Sub
    ElseIf Code = OptPress Then
        FRMMain!TXTOpP.Text = FormatVal(Cur_Disp.OpP)
        FRMMain!LBLOpPUnits.caption = Cur_Disp.OpPUnit
        Exit Sub
    End If
                
    FRMMain!PNLPropVal(Code).caption = ""
    FRMMain!PNLPropUnits(Code).caption = ""
    FRMMain!PNLPropVal(Code).caption = FormatVal(DispMethod(Code).value(DispMethod(Code).CurMethod))
    FRMMain!PNLPropUnits(Code).caption = DispMethod(Code).Unit
    Select Case Code
        Case LD:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case VP:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case Hvap:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case ST:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case VV:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case LV:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case VTC:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case LTC:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case VHC:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case LHC:
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at " & DispMethod(Code).TFT & " " & DispMethod(Code).TFTUnit
        Case ACchem
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at inf. dilution"
        Case ACwater
            FRMMain!PNLPropUnits(Code).caption = RTrim$(FRMMain!PNLPropUnits(Code).caption) & " at inf. dilution"
    End Select
        
End Sub

Public Function FormatVal(Number As Double) As String

    If Number = 0 Then
        FormatVal = ""
    ElseIf Abs(Number) >= 1000 Then
        FormatVal = Format(Number, FormatGT1000)
    ElseIf Abs(Number) <= 0.001 Then
        FormatVal = Format(Number, FormatLT001)
    Else
        FormatVal = Format(Number, FormatGeneral)
    End If
    
End Function

Function GetFamilyGroup(Code As String) As String

    Select Case Code
        Case "AA"
            GetFamilyGroup = "AA - N-ALKANES"
        Case "MA"
            GetFamilyGroup = "MA - METHYLALKANES"
        Case "DA"
            GetFamilyGroup = "DA - DIMETHYLALKANES"
        Case "OA"
            GetFamilyGroup = "OA - OTHER ALKANES"
        Case "CA"
            GetFamilyGroup = "CA - CYCLOALKANES"
        Case "GA"
            GetFamilyGroup = "GA - ALKYLCYCLOPENTANES"
        Case "HA"
            GetFamilyGroup = "HA - ALKYLCYCLOHEXANES"
        Case "RA"
            GetFamilyGroup = "RA - MULTIRING CYCLOALKANES"
        Case "AE"
            GetFamilyGroup = "AE - 1-ALKENES"
        Case "BE"
            GetFamilyGroup = "BE - 2,3,4-ALKENES"
        Case "ME"
            GetFamilyGroup = "ME - METHYLALKENES"
        Case "EE"
            GetFamilyGroup = "EE - ETHYL & HIGHER ALKENES"
        Case "CE"
            GetFamilyGroup = "CE - CYCLOALKENES"
        Case "DE"
            GetFamilyGroup = "DE - DIALKENES"
        Case "YY"
            GetFamilyGroup = "YY - ALKYNES"
        Case "AR"
            GetFamilyGroup = "AR - N-ALKYLBENZENES"
        Case "BR"
            GetFamilyGroup = "BR - OTHER ALKYLBENZENES"
        Case "MR"
            GetFamilyGroup = "MR - OTHER MONOAROMATICS"
        Case "DR"
            GetFamilyGroup = "DR - NAPTHALENES"
        Case "OR"
            GetFamilyGroup = "OR - OTHER CONDENSED RINGS"
        Case "PR"
            GetFamilyGroup = "PR - DIPHENYL/POLYAROMATICS"
        Case "TR"
            GetFamilyGroup = "TR - TERPENES"
        Case "HR"
            GetFamilyGroup = "HR - OTHER HYDROCARBON RINGS"
        Case "IG"
            GetFamilyGroup = "IG - INORGANIC GASES"
        Case "LD"
            GetFamilyGroup = "LD - ALDEHYDES"
        Case "KK"
            GetFamilyGroup = "KK - KETONES"
        Case "AL"
            GetFamilyGroup = "AL - N-ALCOHOLS"
        Case "OL"
            GetFamilyGroup = "OL - OTHER ALIPHATIC ALCOHOLS"
        Case "CL"
            GetFamilyGroup = "CL - CYCLOALIPHATIC ALCOHOLS"
        Case "RL"
            GetFamilyGroup = "RL - AROMATIC ALCOHOLS"
        Case "PL"
            GetFamilyGroup = "PL - POLYOLS"
        Case "AC"
            GetFamilyGroup = "AC - N-ALIPHATIC ACIDS"
        Case "OC"
            GetFamilyGroup = "OC - OTHER ALIPHATIC ACIDS"
        Case "DC"
            GetFamilyGroup = "DC - DICARBOXYLIC ACIDS"
        Case "RC"
            GetFamilyGroup = "RC - AROMATIC CARBOXYLIC ACIDS"
        Case "HD"
            GetFamilyGroup = "HD - ANHYDRIDES"
        Case "FS"
            GetFamilyGroup = "FS - FORMATES"
        Case "ES"
            GetFamilyGroup = "ES - ACETATES"
        Case "BS"
            GetFamilyGroup = "BS - PROPIONATES AND BUTYRATES"
        Case "SS"
            GetFamilyGroup = "SS - OTHER SATURATED ALIPHATIC ESTERS"
        Case "US"
            GetFamilyGroup = "US - UNSATURATED ALIPHATIC ESTERS"
        Case "RS"
            GetFamilyGroup = "RS - AROMATIC ESTERS"
        Case "AT"
            GetFamilyGroup = "AT - ALIPHATIC ETHERS"
        Case "OT"
            GetFamilyGroup = "OT - OTHER ETHERS/DIETHERS"
        Case "CT"
            GetFamilyGroup = "CT - EPOXIDES"
        Case "TT"
            GetFamilyGroup = "TT - PEROXIDES"
        Case "AH"
            GetFamilyGroup = "AH - C1/C2 ALIPHATIC CHLORIDES"
        Case "HH"
            GetFamilyGroup = "HH - C3 & HIGHER ALIPHATIC CHLORIDES"
        Case "RH"
            GetFamilyGroup = "RH - AROMATIC CHLORIDES"
        Case "VH"
            GetFamilyGroup = "VH - C,H,BR COMPOUNDS"
        Case "WH"
            GetFamilyGroup = "WH - C,H,I COMPOUNDS"
        Case "FH"
            GetFamilyGroup = "FH - C,H,F COMPOUNDS"
        Case "PH"
            GetFamilyGroup = "PH - C,H, MULTIHALOGEN COMPOUNDS"
        Case "AM"
            GetFamilyGroup = "AM - N-ALIPHATIC PRIMARY AMINES"
        Case "BM"
            GetFamilyGroup = "BM - OTHER ALIPHATIC AMINES"
        Case "RM"
            GetFamilyGroup = "RM - AROMATIC AMINES"
        Case "OM"
            GetFamilyGroup = "OM - OTHER AMINES, IMINES"
        Case "NX"
            GetFamilyGroup = "NX - NITRILES"
        Case "TN"
            GetFamilyGroup = "TN - C,H,NO2 COMPOUNDS"
        Case "YN"
            GetFamilyGroup = "YN - ISOCYANATES/DIISOCYANATES"
        Case "SD"
            GetFamilyGroup = "SD - MERCAPTANS"
        Case "SF"
            GetFamilyGroup = "SF - SULFIDES/THIOPHENES"
        Case "PC"
            GetFamilyGroup = "PC - POLYFUNCTIONAL ACIDS"
        Case "PS"
            GetFamilyGroup = "PS - POLYFUNCTIONAL ESTERS"
        Case "PO"
            GetFamilyGroup = "PO - OTHER POLYFUNCTIONAL C,H,O"
        Case "NP"
            GetFamilyGroup = "NP - POLYFUNCTIONAL NITRILES"
        Case "TM"
            GetFamilyGroup = "TM - NITROAMINES"
        Case "PM"
            GetFamilyGroup = "PM - POLYFUNCTIONAL AMIDES/AMINES"
        Case "PN"
            GetFamilyGroup = "PN - POLYFUNCTIONAL C,H,O,N"
        Case "SP"
            GetFamilyGroup = "SP - POLYFUNCTIONAL C,H,O,S"
        Case "HP"
            GetFamilyGroup = "HP - POLYFUNCTIONAL C,H,O,HALIDE"
        Case "BP"
            GetFamilyGroup = "BP - POLYFUNCTIONAL C,H,N,HALIDE,(O)"
        Case "OP"
            GetFamilyGroup = "OP - OTHER POLYFUNCTIONAL ORGANICS"
        Case "LX"
            GetFamilyGroup = "LX - ELEMENTS"
        Case "SX"
            GetFamilyGroup = "SX - SILANES/SILOXANES"
        Case "GI"
            GetFamilyGroup = "GI - ORGANIC/INORGANIC COMPOUNDS"
        Case "IC"
            GetFamilyGroup = "IC - INORGANIC ACIDS"
        Case "IB"
            GetFamilyGroup = "IB - INORGANIC BASES"
        Case "GS"
            GetFamilyGroup = "GS - ORGANIC SALTS"
        Case "XS"
            GetFamilyGroup = "XS - SODIUM SALTS"
        Case "OS"
            GetFamilyGroup = "OS - OTHER INORGANIC SALTS"
        Case "IH"
            GetFamilyGroup = "IH - INORGANIC HALIDES"
        Case "OI"
            GetFamilyGroup = "OI - OTHER INORGANICS"
        Case Else
            GetFamilyGroup = ""
    End Select
    
End Function


