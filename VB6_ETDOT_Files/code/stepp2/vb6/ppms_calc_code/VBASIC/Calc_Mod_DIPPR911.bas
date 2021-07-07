Attribute VB_Name = "Calc_Mod_DIPPR911"
Option Explicit






Const Calc_Mod_DIPPR911_decl_end = True


Function Calc_DIPPR911_DoImport( _
    in_idx_Chem As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Db1 As Database
Dim Rs1 As Recordset
'Dim in_Property_Code As Long
'Dim in_Technique_Code As Long
'Dim out_idx_PropertyData As Integer
'Dim out_idx_TechniqueData As Integer
Dim RecFound As Boolean
  'NowProj.UserChemicals(in_idx_Chem).name
  Set Db1 = OpenDatabase(fn_Master_MDB)
'  RecFound = Database_TestForExistingString00( _
      Db1, _
      Rs1, _
      "(n/a)", _
      "(n/a)", _
      "select * from DIPPR911 where [Cas #]=" & _
      ThisCas)
  '
  ' READ IN PROPERTIES THAT LACK TEMPERATURE-DEPENDENT INFORMATION.
  '
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_000_MOLEC_WEIGHT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_001_LIQDENS_298K)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_003_MELTING_POINT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_004_NBP)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_005_VP_298K)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_007_HEAT_FORMATION)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_010_HEAT_OF_VAPORIZATION_298K)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_011_HEAT_OF_VAPORIZATION_NBP)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_013_CRITICAL_T)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_014_CRITICAL_P)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_038_CRITICAL_V)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_015_DIFFUSIVITY_H2O)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_016_DIFFUSIVITY_AIR)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_017_SURFACE_TENSION_298K)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_034_AC_CHEM_IN_H2O)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_032_AC_H2O_IN_CHEM)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_033_HENRY_CONSTANT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_039_SOL_LIMIT_CHEM_IN_H2O)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_035_LOG_KOW)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_036_LOG_KOC)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_037_BIOCONC_FACTOR)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_023_UF_LIMIT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_024_LF_LIMIT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_025_FLASH_POINT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_026_AUTOIGNITION_T)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_027_COMBUSTION_HEAT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_028_CARBON_THOD)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_029_COMBINED_THOD)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_030_COD)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, False, PROPCODE_031_BCOD)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_041_FMINNOW_48H_EC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_042_FMINNOW_96H_EC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_043_FMINNOW_24H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_044_FMINNOW_48H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_045_FMINNOW_96H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_046_SALMONIDAE_24H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_047_SALMONIDAE_48H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_048_SALMONIDAE_96H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_049_DMAGNA_24H_EC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_050_DMAGNA_48H_EC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_051_DMAGNA_24H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_052_DMAGNA_48H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_053_MYSID_96H_LC50)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, False, True, PROPCODE_054_ALTERNATE_SPECIES)
  '
  ' READ IN PROPERTIES THAT HAVE TEMPERATURE-DEPENDENT INFORMATION.
  '
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_002_LIQDENS_FOFT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_018_SURFACE_TENSION_FOFT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_019_VAPOR_VISCOSITY_FOFT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_020_LIQUID_VISCOSITY_FOFT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT)
  Call Calc_DIPPR911_DoImport_OneProperty(Db1, Rs1, in_idx_Chem, True, False, PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT)
  '
  ' CLOSE DATABASE AND EXIT.
  '
  Db1.Close
exit_normally_ThisFunc:
  Calc_DIPPR911_DoImport = True
  Exit Function
exit_err_ThisFunc:
  Calc_DIPPR911_DoImport = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Calc_DIPPR911_DoImport")
  GoTo exit_err_ThisFunc
End Function


Function Calc_DIPPR911_DoImport_OneProperty( _
    Db1 As Database, _
    Rs1 As Recordset, _
    in_idx_Chem As Integer, _
    in_HasFofT As Boolean, _
    in_SeeComment As Boolean, _
    in_Property_Code As Long _
    ) _
    As Boolean
On Error GoTo err_ThisFunc
Dim in_Technique_Code As Long
Dim out_idx_PropertyData As Integer
Dim out_idx_TechniqueData As Integer
Dim Err_IndexNotFound As Boolean
Dim Err_RecNotFound As Boolean
Dim This_UnitBase As String
Dim This_UnitType As String
Dim in_RecFound As Boolean
Dim ThisCas As String
  ThisCas = NowProj.UserChemicals(in_idx_Chem).CAS
  in_Technique_Code = TECHCODE_ANY_991d_DB911
  '
  ' SEARCH FOR THIS RECORD IN THE DATABASE.
  '
  in_RecFound = Database_TestForExistingString00( _
      Db1, _
      Rs1, _
      "(n/a)", _
      "(n/a)", _
      "select * from DIPPR911 where [Cas #]=" & _
      ThisCas & " and [PEARLS Code]=" & _
      in_Property_Code)
  '
  ' LOOK UP THIS TECHNIQUE RECORD.
  '
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
    This_UnitType = .UnitType
    This_UnitBase = .UnitBase
  End With
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(out_idx_PropertyData). _
      TechniqueData(out_idx_TechniqueData)
    If (in_SeeComment = True) Then
      .Text_When_Blank = "See Comment"
    Else
      .Text_When_Blank = ""
    End If
    If (Err_RecNotFound) Or (Err_IndexNotFound) Then
      .IsAvail = False
      If (Err_RecNotFound = True) Then
        .Error_Code = "Data not found in DIPPR911 database."
      Else
        .Error_Code = "Internal DIPPR911 technique index not found."
      End If
      .value = 0#
      .IsTagged = False
      .ReferenceText = ""
    Else
      .IsAvail = True
      .Error_Code = ""
      .IsTagged = False
      .ReferenceText = ""
      '
      ' MAIN SET OF DATA IMPORTS.
      '
'Dim ThisValue As Double
Dim Ref_RecFound As Boolean
Dim Rs2 As Recordset
Dim out_Found As Integer
Dim Ret As String
      ''''OnError GoTo 0
      If (in_HasFofT = False) Then
        '
        ' SET VALUE FOR PROPERTY THAT LACKS T-DEPENDENT INFO.
        '
        .DIPPR_Value = Database_Get_Double(Rs1, "Value")
        .DIPPR_Units = Database_Get_String(Rs1, "Units")
        .DIPPR_R = Database_Get_Integer(Rs1, "Rating")
        If (.Text_When_Blank <> "") And (.DIPPR_Units = "") Then
          .value = 0#
        Else
          '
          ' CONVERT UNITS TO INTERNAL BASE UNITS.
          '
          Call unitsys_convert0( _
              This_UnitType, _
              .DIPPR_Units, _
              This_UnitBase, _
              .DIPPR_Value, _
              .value, _
              out_Found)
          If (out_Found = False) Then
            .Error_Code = "Unable to convert units; " & _
                ".DIPPR_Units=`" & .DIPPR_Units & "`, " & _
                "This_UnitType=`" & This_UnitType & "`, " & _
                "This_UnitBase=`" & This_UnitBase & "`."
            GoTo exit_err_ThisFunc
          End If
        End If
      Else
        '
        ' LOAD IN TEMPERATURE-DEPENDENT INFO.
        '
        .FofT_EqForm = CInt(Database_Get_Long(Rs1, "Equation"))
        ReDim .FofT_Coeffs(1 To 5)
        .FofT_Coeffs(1) = Database_Get_Double(Rs1, "Coef1")
        .FofT_Coeffs(2) = Database_Get_Double(Rs1, "Coef2")
        .FofT_Coeffs(3) = Database_Get_Double(Rs1, "Coef3")
        .FofT_Coeffs(4) = Database_Get_Double(Rs1, "Coef4")
        .FofT_Coeffs(5) = Database_Get_Double(Rs1, "Coef5")
        .FofT_Units_F = Database_Get_String(Rs1, "Units")
        .FofT_Units_T = "K"          '???? Is this correct?
        .FofT_Minimum_T = CInt(Database_Get_Long(Rs1, "Value"))
        .FofT_Maximum_T = CInt(Database_Get_Long(Rs1, "Temperature"))
        ''''.DIPPR_REL = Database_Get_String(Rs1, in_FieldName & "REL")
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
''''        .Value = 0#
''''
        .value = 0#
        .DIPPR_Value = .value
        '
        ' CONVERT UNITS TO INTERNAL BASE UNITS.
        '
        Call unitsys_convert0( _
            This_UnitType, _
            .FofT_Units_F, _
            This_UnitBase, _
            .DIPPR_Value, _
            .value, _
            out_Found)
        If (out_Found = False) Then
          .Error_Code = "Unable to convert units; " & _
              ".FofT_Units_F=`" & .FofT_Units_F & "`, " & _
              "This_UnitType=`" & This_UnitType & "`, " & _
              "This_UnitBase=`" & This_UnitBase & "`."
          GoTo exit_err_ThisFunc
        End If
      End If
      '
      ' HANDLE MISCELLANEOUS DATA IMPORTS THAT ARE COMMON TO
      ' BOTH T-DEPENDENT AND NON-T-DEPENDENT PROPERTIES.
      '
      .DIPPR_Pressure = Database_Get_String(Rs1, "Pressure")
      .DIPPR_DescMethod = Database_Get_String(Rs1, "Desc/Method")
      .DIPPR_Comment = Database_Get_String(Rs1, "Comment")
      .DIPPR_ArticleNumber = Database_Get_Long(Rs1, "Article #")
      '
      ' LOOK UP THE REFERENCE, .DIPPR_ArticleNumber, WITHIN THE DATABASE.
      '
      Ref_RecFound = Database_TestForExistingString00( _
          Db1, _
          Rs2, _
          "(n/a)", _
          "(n/a)", _
          "select * from [CITATION] where [Paper #]=" & _
          Trim$(Str$(.DIPPR_ArticleNumber)))
      If (Ref_RecFound = False) Then
        .ReferenceText = "( Error looking up reference for DIPPR_ArticleNumber=" & _
            .DIPPR_ArticleNumber & " )"
      Else
        Ret = ""
        .ReferenceText = _
            Database_Get_String(Rs2, "Author") & ", " & Ret & _
            Database_Get_String(Rs2, "Title") & ", " & Ret & _
            Database_Get_String(Rs2, "Journal") & ", " & _
            Database_Get_String(Rs2, "Date") & ", " & _
            Database_Get_String(Rs2, "Volume") & ", " & _
            Database_Get_String(Rs2, "Number") & ", " & _
            Database_Get_String(Rs2, "Pages")
      End If
      If (in_HasFofT = True) Then
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
  Calc_DIPPR911_DoImport_OneProperty = True
  Exit Function
exit_err_ThisFunc:
  Calc_DIPPR911_DoImport_OneProperty = False
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
  ''''Call Show_Trapped_Error("Calc_DIPPR911_DoImport_OneProperty")
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(out_idx_PropertyData). _
      TechniqueData(out_idx_TechniqueData)
    .Error_Code = Get_Trapped_Error_String( _
        "Calc_DIPPR911_DoImport_OneProperty")
  End With
  Resume exit_err_ThisFunc
End Function



