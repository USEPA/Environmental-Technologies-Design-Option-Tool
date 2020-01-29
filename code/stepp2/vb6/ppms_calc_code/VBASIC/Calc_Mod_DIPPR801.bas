Attribute VB_Name = "Calc_Mod_DIPPR801"
Option Explicit






Const Calc_Mod_DIPPR801_decl_end = True


Function Calc_DIPPR801_DoImport( _
    in_idx_Chem As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Db1 As Database
Dim Rs1 As Recordset
Dim ThisCas As String
'Dim in_Property_Code As Long
'Dim in_Technique_Code As Long
'Dim out_idx_PropertyData As Integer
'Dim out_idx_TechniqueData As Integer
Dim RecFound As Boolean
  ThisCas = NowProj.UserChemicals(in_idx_Chem).CAS
  'NowProj.UserChemicals(in_idx_Chem).name
  Set Db1 = OpenDatabase(fn_Master_MDB)
  RecFound = Database_TestForExistingString00( _
      Db1, _
      Rs1, _
      "(n/a)", _
      "(n/a)", _
      "select * from DIPPR801 where [CAS]=" & _
      ThisCas)
  '
  ' READ IN PROPERTIES THAT LACK TEMPERATURE-DEPENDENT INFORMATION.
  '
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_000_MOLEC_WEIGHT, "MW")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_013_CRITICAL_T, "TC")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_014_CRITICAL_P, "PC")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_038_CRITICAL_V, "VC")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_003_MELTING_POINT, "MP")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_004_NBP, "NBP")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_007_HEAT_FORMATION, "HFOR")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_027_COMBUSTION_HEAT, "HCOM")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_025_FLASH_POINT, "FP")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_024_LF_LIMIT, "FLML")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_023_UF_LIMIT, "FLMU")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, False, PROPCODE_026_AUTOIGNITION_T, "AIT")
  '
  ' READ IN PROPERTIES THAT HAVE TEMPERATURE-DEPENDENT INFORMATION.
  '
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_002_LIQDENS_FOFT, "LDN")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT, "LCP")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT, "ICP")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT, "HVP")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_018_SURFACE_TENSION_FOFT, "ST")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_019_VAPOR_VISCOSITY_FOFT, "VVS")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_020_LIQUID_VISCOSITY_FOFT, "LVS")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT, "LTC")
  Call Calc_DIPPR801_DoImport_OneProperty(Db1, Rs1, RecFound, in_idx_Chem, True, PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT, "VTC")
  '
  ' CLOSE DATABASE AND EXIT.
  '
  Db1.Close
exit_normally_ThisFunc:
  Calc_DIPPR801_DoImport = True
  Exit Function
exit_err_ThisFunc:
  Calc_DIPPR801_DoImport = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Calc_DIPPR801_DoImport")
  GoTo exit_err_ThisFunc
End Function


Function Calc_DIPPR801_DoImport_OneProperty( _
    Db1 As Database, _
    Rs1 As Recordset, _
    in_RecFound As Boolean, _
    in_idx_Chem As Integer, _
    in_HasFofT As Boolean, _
    in_Property_Code As Long, _
    in_FieldName As String _
    ) _
    As Boolean
On Error GoTo err_ThisFunc
Dim in_Technique_Code As Long
Dim out_idx_PropertyData As Integer
Dim out_idx_TechniqueData As Integer
Dim Err_IndexNotFound As Boolean
Dim Err_RecNotFound As Boolean
Dim This_UnitBase As String
  in_Technique_Code = TECHCODE_ANY_992d_DB801
  Err_IndexNotFound = False
  Err_RecNotFound = False
  If (False = TechniqueData_GetIndex( _
      in_idx_Chem, _
      in_Property_Code, _
      in_Technique_Code, _
      out_idx_PropertyData, _
      out_idx_TechniqueData)) Then
    Err_IndexNotFound = True
  End If
  If (in_RecFound = False) Then
    Err_RecNotFound = True
  End If
  If (Err_IndexNotFound = True) Then
    ' PROPERTY NOT FOUND IN HIERARCHY !!!
    GoTo exit_err_ThisFunc
  End If
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(out_idx_PropertyData)
      This_UnitBase = .UnitBase
  End With
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(out_idx_PropertyData). _
      TechniqueData(out_idx_TechniqueData)
    If (Err_RecNotFound) Or (Err_IndexNotFound) Then
      .IsAvail = False
      If (Err_RecNotFound = True) Then
        .Error_Code = "Data not found in DIPPR801 database."
      Else
        .Error_Code = "Internal DIPPR801 technique index not found."
      End If
      .value = 0#
      .IsTagged = False
      .ReferenceText = ""
    Else
      .IsAvail = True
      .Error_Code = ""
      .IsTagged = False
      .ReferenceText = ""
      If (in_HasFofT = False) Then
        '
        ' SET VALUE FOR PROPERTY THAT LACKS T-DEPENDENT INFO.
        '
        .value = Database_Get_Double(Rs1, in_FieldName)
        ''''.DIPPR_REF = Database_Get_String(Rs1, in_FieldName & "REF")
        .DIPPR_REF = Database_Get_StringNoTrim(Rs1, in_FieldName & "REF")
        .DIPPR_R = Database_Get_String(Rs1, in_FieldName & "R")
        '
        ' LOOK UP THE REFERENCE, .DIPPR_REF, WITHIN THE DATABASE.
        '
Dim Ref_RecFound As Boolean
Dim Rs2 As Recordset
        Ref_RecFound = Database_TestForExistingString00( _
            Db1, _
            Rs2, _
            "(n/a)", _
            "(n/a)", _
            "select * from [801refs] where [RefNum]='" & _
            .DIPPR_REF & "'")
        If (Ref_RecFound = False) Then
          .ReferenceText = "( Error looking up reference for RefNum=`" & .DIPPR_REF & "` )"
        Else
          .ReferenceText = Database_Get_String(Rs2, "REFERENCE")
        End If
        Rs2.Close
      Else
        '
        ' LOAD IN TEMPERATURE-DEPENDENT INFO.
        '
        .FofT_EqForm = CInt(Database_Get_Long(Rs1, in_FieldName & "EQN"))
        ReDim .FofT_Coeffs(1 To 5)
        .FofT_Coeffs(1) = Database_Get_Double(Rs1, in_FieldName & "A")
        .FofT_Coeffs(2) = Database_Get_Double(Rs1, in_FieldName & "B")
        .FofT_Coeffs(3) = Database_Get_Double(Rs1, in_FieldName & "C")
        .FofT_Coeffs(4) = Database_Get_Double(Rs1, in_FieldName & "D")
        .FofT_Coeffs(5) = Database_Get_Double(Rs1, in_FieldName & "E")
        .FofT_Units_F = This_UnitBase
        .FofT_Units_T = "K"
        .FofT_Minimum_T = CInt(Database_Get_Long(Rs1, in_FieldName & "TMIN"))
        .FofT_Maximum_T = CInt(Database_Get_Long(Rs1, in_FieldName & "TMAX"))
        .DIPPR_REL = Database_Get_String(Rs1, in_FieldName & "REL")
''''
''''        '
''''        ' CALCULATE VALUE OF PROPERTY AT OPERATING TEMPERATURE.
''''        '
''''        .Value = 0#
''''        Call Calc_FofT_Equation( _
''''            in_idx_Chem, _
''''            NowProj.UserChemicals(in_idx_Chem). _
''''            PropertyData(out_idx_PropertyData). _
''''            TechniqueData(out_idx_TechniqueData))
''''
        .value = 0#
        '
        ' SET THE REFERENCE FOR THIS T-DEPENDENT PROPERTY.
        '
        If (False = Calc_Mod_GetRefText( _
            NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(out_idx_PropertyData). _
            TechniqueData(out_idx_TechniqueData))) Then
          .Error_Code = "Unable to set reference!"
          GoTo exit_err_ThisFunc
        End If
      End If
    End If
  End With
exit_normally_ThisFunc:
  Calc_DIPPR801_DoImport_OneProperty = True
  Exit Function
exit_err_ThisFunc:
  Calc_DIPPR801_DoImport_OneProperty = False
  If (Err_IndexNotFound = False) Then
    With NowProj.UserChemicals(in_idx_Chem). _
        PropertyData(out_idx_PropertyData). _
        TechniqueData(out_idx_TechniqueData)
      .IsAvail = False
      .value = 0#
      .IsTagged = False
      .ReferenceText = ""
    End With
  End If
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("Calc_DIPPR801_DoImport_OneProperty")
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(out_idx_PropertyData). _
      TechniqueData(out_idx_TechniqueData)
    .Error_Code = Get_Trapped_Error_String( _
        "Calc_DIPPR801_DoImport_OneProperty")
  End With
  Resume exit_err_ThisFunc
End Function



