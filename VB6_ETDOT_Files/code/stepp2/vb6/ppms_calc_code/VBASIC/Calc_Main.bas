Attribute VB_Name = "Calc_Main"
Option Explicit








Const Calc_Main_decl_end = True


Function Recalculate_All() _
    As Boolean
On Error GoTo err_ThisFunc
Dim i As Integer
Dim UB As Integer
  UB = UBound(NowProj.UserChemicals)
  For i = 1 To UB
    Call Recalculate_OneChemical(i)
  Next i
exit_normally_ThisFunc:
  Recalculate_All = True
  Exit Function
exit_err_ThisFunc:
  Recalculate_All = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Recalculate_All")
  Resume exit_err_ThisFunc
End Function


Function Recalculate_OneChemical( _
    in_idx_Chem As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim UB_Tech1 As Integer
Dim UB_Tech2 As Integer
Dim i1 As Integer
Dim i2 As Integer
  '
  '
  ' TURN ON HOURGLASS POINTER.
  '
  frmMain.MousePointer = 11
  '
  '
  ' ZERO-OUT ALL TECHNIQUES, AND TAG VALID TECHNIQUES.
  '
  If (TagValidTechniques_OneChemical(in_idx_Chem) = False) Then
    GoTo exit_err_ThisFunc
  End If
  '
  '
  ' FIRST, HANDLE THE DIPPR801 DATABASE IMPORTS.  FOR EACH IMPORTED
  ' PROPERTY (WHETHER SUCCESSFUL OR NOT), TURN OFF THE .IsTagged PROPERTY
  ' SO THE CODE FURTHER DOWN DOES NOT ATTEMPT TO CALCULATE IT.
  '
  Call Calc_DIPPR801_DoImport(in_idx_Chem)
  '
  '
  ' SECOND, HANDLE THE DIPPR911 DATABASE IMPORTS.  FOR EACH IMPORTED
  ' PROPERTY (WHETHER SUCCESSFUL OR NOT), TURN OFF THE .IsTagged PROPERTY
  ' SO THE CODE FURTHER DOWN DOES NOT ATTEMPT TO CALCULATE IT.
  '
  Call Calc_DIPPR911_DoImport(in_idx_Chem)
  '
  '
  ' RE-DETERMINE TECHNIQUES TO USE.
  '
  Call ReDetermine_idx_Technique_Used(in_idx_Chem)
  '
  '
  ' CALCULATE THE T-DEPENDENT PROPERTIES.
  '
  UB_Tech1 = UBound(NowProj.UserChemicals(in_idx_Chem).PropertyData)
  For i1 = 1 To UB_Tech1
    If (NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(i1).Is_FofT = True) Then
      UB_Tech2 = UBound(NowProj.UserChemicals(in_idx_Chem). _
          PropertyData(i1).TechniqueData)
''''If (NowProj.UserChemicals(in_idx_Chem).PropertyData(i1).Property_Code = PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT) Then
''''  Debug.Print "blah"
''''End If
      For i2 = 1 To UB_Tech2
        NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(i1). _
            TechniqueData(i2).value = 0#
        Call Calc_FofT_Equation( _
            in_idx_Chem, _
            NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(i1). _
            TechniqueData(i2))
      Next i2
    End If
  Next i1
  '
  '
  ' RE-DETERMINE TECHNIQUES TO USE.
  '
  Call ReDetermine_idx_Technique_Used(in_idx_Chem)
  '
  '
  ' CALCULATE THE BLOCK5 VALUES.
  '
  Call Block5_Calculate_All_Block5_Values(in_idx_Chem)
  '
  '
  ' RE-DETERMINE TECHNIQUES TO USE.
  '
  Call ReDetermine_idx_Technique_Used(in_idx_Chem)
'  '
'  ' MAIN CALCULATION BLOCK.
'  '
'  UB_Tech1 = UBound(NowProj.UserChemicals(in_idx_Chem).PropertyData)
'  For i1 = 1 To UB_Tech1
'    UB_Tech2 = UBound(NowProj.UserChemicals(in_idx_Chem). _
'        PropertyData(i1).TechniqueData)
'    For i2 = 1 To UB_Tech2
'      '
'      ' IF TAGGED FOR CALCULATION, CALCULATE THIS TECHNIQUE.
'      '
'      If (NowProj.UserChemicals(in_idx_Chem). _
'          PropertyData(i1).TechniqueData(i2).IsTagged) Then
'        Call Recalculate_OneTechnique( _
'            in_idx_Chem, _
'            i1, _
'            i2)
'      End If
'    Next i2
'    '
'    ' CALCULATE THIS PROPERTY BASED ON TECHNIQUE VALUES
'    ' THAT HAVE ALREADY BEEN CALCULATED.
'    '
'    Call Recalculate_OneProperty( _
'        in_idx_Chem, _
'        i1)
'  Next i1
exit_normally_ThisFunc:
  Recalculate_OneChemical = True
  GoTo exit_ThisFunc
exit_err_ThisFunc:
  Recalculate_OneChemical = False
  GoTo exit_ThisFunc
exit_ThisFunc:
  '
  ' TURN OFF HOURGLASS POINTER.
  '
  frmMain.MousePointer = 0
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Recalculate_OneChemical")
  Resume exit_err_ThisFunc
End Function


Function ReDetermine_idx_Technique_Used( _
    in_idx_Chem As Integer _
    ) _
    As Boolean
On Error GoTo err_ThisFunc
Dim UB_Tech1 As Integer
Dim i1 As Integer
  UB_Tech1 = UBound(NowProj.UserChemicals(in_idx_Chem).PropertyData)
  For i1 = 1 To UB_Tech1
    Call Recalculate_OneProperty( _
        in_idx_Chem, _
        i1)
  Next i1
exit_normally_ThisFunc:
  ReDetermine_idx_Technique_Used = True
  Exit Function
exit_err_ThisFunc:
  ReDetermine_idx_Technique_Used = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("ReDetermine_idx_Technique_Used")
  GoTo exit_err_ThisFunc
End Function


Function Recalculate_OneProperty( _
    in_idx_Chem As Integer, _
    in_idx_PropertyData As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Use_Error_Code As String
Dim in_Property_Code As Long
Dim out_idx_PropertySheetOrder() As Integer
Dim out_idx_PropertyOrder() As Integer
Dim out_Size As Integer
Dim idx_PropertySheetOrder As Integer
Dim idx_PropertyOrder As Integer
Dim UB As Integer
Dim i As Integer
Dim in_Technique_Code As Long
Dim out_idx_PropertyData As Integer
Dim out_idx_TechniqueData As Integer
Dim idx_FirstValidTechnique As Integer
Dim Found_FirstValidTechnique As Boolean
Dim idx_OverrideTechnique As Integer
Dim Found_OverrideTechnique As Boolean
Dim This_Override_Technique_Code As Long
  '
  ' SET DEFAULTS.
  '
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(in_idx_PropertyData)
    .IsAvail = False
    .idx_Technique_Used = -1
    This_Override_Technique_Code = .Override_Technique_Code
  End With
  '
  ' SEARCH THROUGH HIERARCHY FOR FIRST OCCURENCE OF THIS PROPERTY.
  '
  in_Property_Code = NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(in_idx_PropertyData).Property_Code
  If (False = PropertyOrder_Property_Code_GetIndexes( _
      in_Property_Code, _
      out_idx_PropertySheetOrder(), _
      out_idx_PropertyOrder(), _
      out_Size)) Then
    GoTo exit_err_ThisFunc
  End If
  If (out_Size <= 0) Then
    GoTo exit_err_ThisFunc
  End If
  idx_PropertySheetOrder = out_idx_PropertySheetOrder(1)
  idx_PropertyOrder = out_idx_PropertyOrder(1)
  '
  ' SEARCH FOR FIRST VALID TECHNIQUE, AND FOR OVERRIDE TECHNIQUE (IF ANY).
  '
  UB = UBound(NowProj.UserHierarchy. _
      PropertySheetOrder(idx_PropertySheetOrder). _
      PropertyOrder(idx_PropertyOrder).Technique_Code)
  idx_FirstValidTechnique = -1
  Found_FirstValidTechnique = False
  idx_OverrideTechnique = -1
  Found_OverrideTechnique = False
  For i = 1 To UB
    in_Technique_Code = NowProj.UserHierarchy. _
        PropertySheetOrder(idx_PropertySheetOrder). _
        PropertyOrder(idx_PropertyOrder).Technique_Code(i)
    If (False = TechniqueData_GetIndex( _
        in_idx_Chem, _
        in_Property_Code, _
        in_Technique_Code, _
        out_idx_PropertyData, _
        out_idx_TechniqueData)) Then
      GoTo exit_err_ThisFunc
    End If
    With NowProj.UserChemicals(in_idx_Chem). _
        PropertyData(out_idx_PropertyData). _
        TechniqueData(out_idx_TechniqueData)
      If (.IsAvail = True) Then
        If (Found_FirstValidTechnique = False) Then
          Found_FirstValidTechnique = True
          idx_FirstValidTechnique = out_idx_TechniqueData
        End If
        'NowProj.UserChemicals(in_idx_Chem). _
        '    PropertyData(out_idx_PropertyData).IsAvail = True
        'NowProj.UserChemicals(in_idx_Chem). _
        '    PropertyData(out_idx_PropertyData).idx_Technique_Used = _
        '    out_idx_TechniqueData
        'Exit For
        If (This_Override_Technique_Code <> -1) Then
          If (.Technique_Code = This_Override_Technique_Code) Then
            Found_OverrideTechnique = True
            idx_OverrideTechnique = out_idx_TechniqueData
          End If
        End If
      End If
    End With
  Next i
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(in_idx_PropertyData)
    If (Found_OverrideTechnique = True) Then
      '
      ' IMPLEMENT OVERRIDE TECHNIQUE.
      '
      .IsAvail = True
      .idx_Technique_Used = idx_OverrideTechnique
    Else
      '
      ' CLEAR OVERRIDE TECHNIQUE (IF PRESENT) AND USE FIRST
      ' VALID TECHNIQUE (IF PRESENT).
      '
      .Override_Technique_Code = -1
      If (Found_FirstValidTechnique = True) Then
        '
        ' USE FIRST VALID TECHNIQUE.
        '
        .IsAvail = True
        .idx_Technique_Used = idx_FirstValidTechnique
      Else
        '
        ' PROPERTY IS UNAVAILABLE.
        '
        ' (KEEP DEFAULT VALUES SET AT THE TOP OF THIS FUNCTION.)
        '
      End If
    End If
  End With
exit_normally_ThisFunc:
  Recalculate_OneProperty = True
  Exit Function
exit_err_ThisFunc:
  Recalculate_OneProperty = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Recalculate_OneProperty")
  Resume exit_err_ThisFunc
End Function


Function Recalculate_OneTechnique( _
    in_idx_Chem As Integer, _
    in_idx_PropertyData As Integer, _
    in_idx_TechniqueData As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Use_Error_Code As String
Dim This_Technique_Code As Long
Dim Skip_It As Boolean
  Use_Error_Code = "Improperly set error code!"
  This_Technique_Code = NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(in_idx_PropertyData). _
      TechniqueData(in_idx_TechniqueData).Technique_Code
  Skip_It = False
  Select Case This_Technique_Code
    Case TECHCODE_ANY_000u_USER_INPUT:
    Case TECHCODE_ANY_991d_DB911:
      Skip_It = True
    Case TECHCODE_ANY_992d_DB801:
      Skip_It = True
    Case TECHCODE_000_002e_UNIFAC:
    Case TECHCODE_001_003e_BHIRUDS_1978:
      'If (False = Calc_Bhiruds_1978(Use_Error_Code)) Then
      '  GoTo exit_err_ThisFunc
      'End If
    Case TECHCODE_001_004e_RACKETT_1978:
          '
          '
          '
          ' MORE TO COME LATER .........
          '
          '
          '
  End Select
'  If (Skip_It = False) Then
'    With NowProj.UserChemicals(in_idx_Chem). _
'            PropertyData(in_idx_PropertyData). _
'            TechniqueData(in_idx_TechniqueData)
'      '
'      ' TEMPORARY TESTING CODE!
'      '
'      'If (This_Technique_Code <> TECHCODE_ANY_000u_USER_INPUT) Then
'      '  .IsAvail = True
'      'End If
'      .IsAvail = True
'      If (This_Technique_Code = TECHCODE_ANY_991d_DB911) Or _
'          (This_Technique_Code = TECHCODE_ANY_000u_USER_INPUT) Then
'        .IsAvail = False
'      End If
'      .Error_Code = ""
'      .Value = CDbl(This_Technique_Code)
'      '
'      ' TEMPORARY TESTING CODE!  (ENDS)
'      '
'    End With
'  End If
exit_normally_ThisFunc:
  Recalculate_OneTechnique = True
  Exit Function
exit_err_ThisFunc:
  Recalculate_OneTechnique = False
  With NowProj.UserChemicals(in_idx_Chem). _
          PropertyData(in_idx_PropertyData). _
          TechniqueData(in_idx_TechniqueData)
    .IsAvail = False
    .Error_Code = Use_Error_Code
    .value = 0#
  End With
  Exit Function
err_ThisFunc:
  Use_Error_Code = _
      Get_Trapped_Error_String("Recalculate_OneTechnique")
  Resume exit_err_ThisFunc
End Function


Function TagValidTechniques_OneChemical( _
    in_idx_Chem As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim UB_Tech1 As Integer
Dim UB_Tech2 As Integer
Dim i1 As Integer
Dim i2 As Integer
Dim in_Property_Code As Long
Dim in_Technique_Code As Long
Dim out_idx_PropertySheetOrder() As Integer
Dim out_idx_PropertyOrder() As Integer
Dim out_idx_Technique_Code() As Integer
Dim out_Size As Integer
  '
  ' ITERATE THROUGH EVERY PROPERTY AND TECHNIQUE IN THIS CHEMICAL.
  '
  UB_Tech1 = UBound(NowProj.UserChemicals(in_idx_Chem).PropertyData)
  For i1 = 1 To UB_Tech1
    UB_Tech2 = UBound(NowProj.UserChemicals(in_idx_Chem). _
        PropertyData(i1).TechniqueData)
    For i2 = 1 To UB_Tech2
      '
      ' ZERO-OUT THIS PROPERTY-TECHNIQUE.
      '
      With NowProj.UserChemicals(in_idx_Chem). _
          PropertyData(i1).TechniqueData(i2)
        .IsAvail = False
        .value = 0#
      End With
      '
      ' IS THIS PROPERTY-TECHNIQUE COMBINATION PRESENT IN
      ' THE GLOBAL HIERARCHY?
      '
      in_Property_Code = NowProj.UserChemicals(in_idx_Chem). _
          PropertyData(i1).Property_Code
      in_Technique_Code = NowProj.UserChemicals(in_idx_Chem). _
          PropertyData(i1).TechniqueData(i2).Technique_Code
      Call PropertyOrder_Technique_Code_GetIndexes( _
          in_Property_Code, _
          in_Technique_Code, _
          out_idx_PropertySheetOrder(), _
          out_idx_PropertyOrder(), _
          out_idx_Technique_Code(), _
          out_Size)
      '
      ' IF IT IS PRESENT IN THE HIERARCHY, TAG IT FOR CALCULATION.
      '
      With NowProj.UserChemicals(in_idx_Chem). _
          PropertyData(i1).TechniqueData(i2)
        If (out_Size > 0) Then
          .IsTagged = True
        Else
          .IsTagged = False
        End If
      End With
    Next i2
  Next i1
exit_normally_ThisFunc:
  TagValidTechniques_OneChemical = True
  Exit Function
exit_err_ThisFunc:
  TagValidTechniques_OneChemical = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("TagValidTechniques_OneChemical")
  Resume exit_err_ThisFunc
End Function






